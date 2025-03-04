VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberBlockSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS         10/29/08      Creation
'---------------------------------------------------------------------------------------
Option Explicit

Private Const MODULE = "MemberBlockSym"
Private Const MYSYMBOL_PROGID = "SPSProgrammedShapes.MemberBlockSym"
Private Const MYSYMBOL_LIBNAME = "SPSProgrammedShapes_MemberBlockSym"

Implements IJDUserSymbolServices
Implements IStructInsulationGraphicInputFilter

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_GetDefinitionName = MYSYMBOL_PROGID
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    DefineEncasementSymbolInputs pSymbolDefinition
   
    DefineEncasementSymbolOutput MYSYMBOL_LIBNAME, MYSYMBOL_PROGID, pSymbolDefinition
   
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)
    Set oSymbolFactory = Nothing
    
    pSymbolDefinition.ProgId = MYSYMBOL_PROGID
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    pSymbolDefinition.SupportOnlyOption = igSYMBOL_SUPPORT_ONLY    'go away when no occurences
    
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub CMEvaluateEncasementCrossSection(pRepSCM As IJDRepresentationStdCustomMethod)
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
 
    Dim oRepDuringGame As IMSSymbolEntities.IJDRepresentationDuringGame
    Dim oInput As IJDInput
    
    Dim thickness As Double
    Dim iIJProxy As IJDProxy
    Dim iInsulationSpec As IStructInsulationSpec
    Dim iMemberPart As ISPSMemberPartPrismatic
    
    Dim ii As Long
    Dim gboxProfile As GBox
    Dim xlo As Double, ylo As Double, xhi As Double, yhi As Double
    Dim iRangeAlias As IJRangeAlias
    Dim elesProfiles As IJElements
    Dim XServices As New CrossSectionServices
    Dim iProfile As SectionProfile
    
    Dim oHoles As Object, elesProfileHoles As IJElements, elesAllHoles As IJElements
    Dim edgeIds(4) As Long

    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oLine As Object
    Dim elesLines As IJElements
    Dim oComplexString As ComplexString3d
    
    Dim outputColl As IJDOutputCollection

    Set oRepDuringGame = pRepSCM
    Set outputColl = oRepDuringGame.outputCollection

    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(1)
    Set iIJProxy = oInput.IJDInputDuringGame.Argument
    Set iInsulationSpec = iIJProxy.Source
    thickness = iInsulationSpec.thickness

    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(3)
    Set iMemberPart = oInput.IJDInputDuringGame.Argument
    
    XServices.GetProfiles iMemberPart.CrossSection.symbol, MEMBER_CROSSSECTION_REPNAME, elesProfiles

    For ii = 1 To elesProfiles.Count
        Set iProfile = elesProfiles(ii)
        Set iRangeAlias = iProfile.Geometry
        gboxProfile = iRangeAlias.GetRange
        If ii = 1 Then
            xlo = gboxProfile.m_low.x
            ylo = gboxProfile.m_low.y
            xhi = gboxProfile.m_high.x
            yhi = gboxProfile.m_high.y
        Else
            If gboxProfile.m_low.x < xlo Then xlo = gboxProfile.m_low.x
            If gboxProfile.m_low.y < ylo Then ylo = gboxProfile.m_low.y
            If gboxProfile.m_high.x > xhi Then xhi = gboxProfile.m_high.x
            If gboxProfile.m_high.y > yhi Then yhi = gboxProfile.m_high.y
        End If
        
        '' accum all holes into one list
        Set oHoles = iProfile.Holes
        If Not oHoles Is Nothing Then
            Set elesProfileHoles = oHoles
            If elesAllHoles Is Nothing Then
                Set elesAllHoles = elesProfileHoles
            Else
                elesAllHoles.AddElements elesProfileHoles
            End If
            Set elesProfileHoles = Nothing
        End If
    
    Next ii
    
    ' expand the range by thickness in all directions
    xlo = xlo - thickness
    ylo = ylo - thickness
    xhi = xhi + thickness
    yhi = yhi + thickness

    Set elesLines = New JObjectCollection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, xlo, ylo, 0#, xhi, ylo, 0#)       ' bottom
    elesLines.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, xhi, ylo, 0#, xhi, yhi, 0#)       ' right
    elesLines.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, xhi, yhi, 0#, xlo, yhi, 0#)       ' top
    elesLines.Add oLine
    Set oLine = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, xlo, yhi, 0#, xlo, ylo, 0#)       ' left
    elesLines.Add oLine
    Set oLine = Nothing
    Set oComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, elesLines)
    Set elesLines = Nothing

    '' edgeIDs are same as for member cross-sections.  For more info, consult documentation.
    '''''    JXSEC_LEFT_WEB   257
    '''''    JXSEC_RIGHT_WEB  258
    '''''
    '''''    JXSEC_BOTTOM_SIDE_OF_INFERIOR_FLANGE   513
    '''''    JXSEC_TOP_SIDE_OF_SUPERIOR_FLANGE      514

    edgeIds(0) = 513
    edgeIds(1) = 258
    edgeIds(2) = 514
    edgeIds(3) = 257

    BuildEncasementSymbolOutput oComplexString, edgeIds, elesAllHoles, outputColl

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    End Sub

Private Function IStructInsulationGraphicInputFilter_FilterGraphicInput(ByVal oGraphicObject As Object, ByVal oDefinition As Object) As Long
Const METHOD = "IStructInsulationGraphicInputFilter_FilterGraphicInput"
On Error GoTo ErrorHandler

    IStructInsulationGraphicInputFilter_FilterGraphicInput = FilterMembers(oGraphicObject, oDefinition)

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria() As String
    IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria = "[" & MYSYMBOL_PROGID & ",FilterMembers]"
End Function

Private Function IStructInsulationGraphicInputFilter_SingleGraphicInput() As Boolean
    IStructInsulationGraphicInputFilter_SingleGraphicInput = True
End Function

Public Function FilterMembers(oGraphicObject As Object, oDefinition As Object) As Integer
Const METHOD = "FilterMembers"
On Error GoTo ErrorHandler

    FilterMembers = ValidateMemberFPCriteria(oGraphicObject, oDefinition)
    If FilterMembers = StructInsulationInputHelper_Ok Then
        FilterMembers = 1
    Else
        FilterMembers = 0
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function
