VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberCirRectSym"
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

Private Const MODULE = "MemberCirRectSym"
Private Const MYSYMBOL_PROGID = "SPSProgrammedShapes.MemberCirRectSym"
Private Const MYSYMBOL_LIBNAME = "SPSProgrammedShapes_MemberCirRectSym"

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
    'See S:\CommonStruct\Middle\Services\StructGenericSymbol.cls for how to do this for a 2D graphic symbol
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    'Define the inputs

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
    Dim XServices As New CrossSectionServices
    Dim iProfile As SectionProfile
    
    Dim oHoles As Object, elesProfileHoles As IJElements, elesAllHoles As IJElements
    Dim edgeIds(2) As Long

    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim elesProfiles As IJElements
    
    'Dim pOutputs As IMSSymbolEntities.IJDOutputs
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
    
    Dim dDiameter As Double
    dDiameter = Sqr((xhi - xlo) * (xhi - xlo) + (yhi - ylo) * (yhi - ylo))
    
    Dim dCenX As Double, dCenY As Double
    
    dCenX = (xhi + xlo) * 0.5
    dCenY = (yhi + ylo) * 0.5
   
    Dim oCircle As Circle3d
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, dCenX, dCenY, 0#, _
                                                                    0#, 0#, 1#, dDiameter / 2 + thickness)
    edgeIds(0) = 3073       ' taken from HSSC CrossSection EdgeId attribute
    edgeIds(1) = 258
                                                                    
    BuildEncasementSymbolOutput oCircle, edgeIds, elesAllHoles, outputColl
    
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


