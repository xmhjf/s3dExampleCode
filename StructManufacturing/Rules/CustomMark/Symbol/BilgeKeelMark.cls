VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BilgeKeelMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: MfgGenericMarkSymbol
'
'File: BilgeKeelMark
'
'Abstract: implementation of BilgeKeel mark symbol
'
'Author: Siva
'
'Revision:
'*******************************************************************

Option Explicit

Implements IJDUserSymbolServices
Implements IJMfgGenericMarkInputs

Private Const MODULE = "MfgGenericMarkSymbol.BilgeKeelMark(BilgeKeelMark.cls)"
Private Const m_progID = "MfgGenericMarkSymbol.BilgeKeelMark"

'---------------------------------------------------
'   INPUT Names
'---------------------------------------------------
Private Const INPUT_1 = "PlateSystem"
Private Const INPUT_2 = "MoldedFormSystem"
Private Const INPUT_3 = "ReferenceSurface"
Private Const INPUT_4 = "MarkingSide"

'---------------------------------------------------
'   OUTPUT and REPRESENTATION Names
'---------------------------------------------------

Private Const REPRESENTATION1 = "Simple" ' used to named the simple physical representation

Private m_strINPUT_1                            As String
Private m_strINPUT_2                            As String

Private m_strPlateSysPrompt                     As String
Private m_strMFSysPrompt                        As String
Private m_strMarkingSide                        As String
Private m_strRefSurface                         As String
Private m_strRefSurfacePrompt                   As String

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    InitializeLocalizeVariables
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: EditOccurence
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the edition of the passed symbol occurrence
'********************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, _
                                                     ByVal pTransactionMgr As Object) As Boolean
    'TrackText "BilgeKeelMark::IJDUserSymbolServices_EditOccurence called", 0
    
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: GetDefinitionName
'
'Interface: IJDUserSymbolServices
'
'Abstract: get the symbol definition name (progID)
'********************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    'TrackText "BilgeKeelMark::IJDUserSymbolServices_GetDefinitionName called", 0
    
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
    
    IJDUserSymbolServices_GetDefinitionName = m_progID

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeSymbolDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: constructs the symbol definition by (re)defining the inputs, outputs,
'          representations, and representation
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(ppSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
    'TrackText "BilgeKeelMark::IJDUserSymbolServices_InitializeSymbolDefinition called", 0

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    'remove all previous Symbol Definition information
    ppSymbolDef.IJDInputs.RemoveAllInput
    ppSymbolDef.IJDRepresentations.RemoveAllRepresentation
    ppSymbolDef.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    '----------------
    'Define inputs
    '----------------
    
    'get the input interface of the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = ppSymbolDef

    '-----------------
    'PlateSystem Input
    '-----------------
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput

    oInput.name = INPUT_1
    oInput.Description = INPUT_1
    oInput.Properties = 0
    oInputs.SetInput oInput, 1
    oInput.Reset
    
    Set oInput = Nothing
    '----------------
    'MoldedFormSystem Input
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_2
    oInput.Description = INPUT_2
    oInput.Properties = 0
    oInputs.SetInput oInput, 2
    oInput.Reset

    Set oInput = Nothing
    
    '----------------
    'Reference surface Input
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_3
    oInput.Description = INPUT_3
    oInput.Properties = igINPUT_IS_A_PARAMETER
    
    'create an input parameter description
    Dim oPC As IJDParameterContent
    Set oPC = New IMSSymbolEntities.DParameterContent
  
    oPC.Type = igValue
    oPC.UomValue = 0

    oInput.DefaultParameterValue = oPC
    oInputs.SetInput oInput, 3
    oInput.Reset

    Set oInput = Nothing
    Set oPC = Nothing
    
    '----------------
    'Marking side Input
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_4
    oInput.Description = INPUT_4
    oInput.Properties = igINPUT_IS_A_PARAMETER
    
    'create an input parameter description
    Set oPC = New IMSSymbolEntities.DParameterContent
  
    oPC.Type = igValue
    oPC.UomValue = 0

    oInput.DefaultParameterValue = oPC
    oInputs.SetInput oInput, 4
    oInput.Reset

    Set oInput = Nothing
    Set oInputs = Nothing
    Set oPC = Nothing
    
    '-------------------------------------
    'Define outputs and representation
    '-------------------------------------
    'set the outputs
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = ppSymbolDef
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation
    
    oRep.RepresentationId = 1  'define a aspect 0 (Simple_physical)
    oRep.name = REPRESENTATION1
    oRep.Description = REPRESENTATION1
    oRep.Properties = igCOLLECTION_VARIABLE

    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep

    oOutputs.Property = igCOLLECTION_VARIABLE

    Set oOutputs = Nothing
    
    oReps.SetRepresentation oRep
    
    Set oRep = Nothing
    Set oReps = Nothing
    
    '------------------------
    'Define evaluations
    '------------------------
    Dim oRepEvals As IJDRepresentationEvaluations
    Set oRepEvals = ppSymbolDef
    
    Dim oRepEval As IJDRepresentationEvaluation
    Set oRepEval = New DRepresentationEvaluation
    
    oRepEval.name = REPRESENTATION1
    oRepEval.Description = "Evaluation function identification for the simple representation"
    oRepEval.Properties = igREPRESENTATION_HIDDEN
    oRepEval.Type = igREPRESENTATION_VBFUNCTION
    oRepEval.ProgId = m_progID

    oRepEvals.AddRepresentationEvaluation oRepEval

    Set oRepEval = Nothing
    Set oRepEvals = Nothing
    
CleanUp:

    Set oRep = Nothing
    Set oReps = Nothing
    Set oRepEval = Nothing
    Set oRepEvals = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InstanciateDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: instantiates and returns an initialized symbol definition
'********************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    'TrackText "IJDUserSymbolServices_InstanciateDefinition called", 0

    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IJDSymbolDefinition

    'instanciate the symbol definition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)

    'set the definition ProgId and CodeBase
    oSymbolDefinition.ProgId = m_progID
    oSymbolDefinition.CodeBase = CodeBase

    'feed the symbol definition with dynamic information (Inputs, Representation, Outputs)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    'set the definition name
    oSymbolDefinition.name = IJDUserSymbolServices_GetDefinitionName(defParams)

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition

CleanUp:
    Set oSymbolFactory = Nothing
    Set oSymbolDefinition = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InvokeRepresentation
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the evaluation function associated with the representation
'          with the set of passed inputs as arguments. the function has to compute
'          the symbol graphic outputs that will be attached to the specified
'          Output Collection object

'********************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, _
                                                       ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    'TrackText "BilgeKeelMark::IJDUserSymbolServices_InvokeRepresentation called", 0

    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    On Error GoTo ErrorHandler
    
    Dim lValue As Long
    lValue = GetJContext().GetVariable("MARKINGBYDEFINITION")

    If lValue = 0 Then
        ' lvalue = 0 indicates that the symbol is triggered because of modification to its inputs or copy-paste
            
        Dim pStructAssocCompute As IJStructAssocCompute
        Set pStructAssocCompute = New StructAssocTools
        
        Dim lAssocFlags As Long
        pStructAssocCompute.GetAssocFlags pSymbolOccurrence, lAssocFlags
        
        If ((lAssocFlags And RELATION_INSERTED_IN_TRANS) = RELATION_INSERTED_IN_TRANS) Then
            ' Do nothing as symbol called from Copy-Paste process
        Else
            ' Here the case is symbol input modification
            ' Reconnect the old outputs to symbol
            ReconnectOldOutputsToSymbol pSymbolOccurrence, pOutputColl
            Exit Sub
        End If
    End If
    
    Dim oPlateSystem    As Object
    Dim oMFSystem       As Object
    Dim lMarkingSide    As Long
    Dim lRefSurface     As Long
    
    '-------------------
    'gather all inputs
    '-------------------
    'TrackText "gather all inputs", 0
    
    On Error Resume Next
    'get the required inputs
    Set oPlateSystem = arrayOfInputs(1)
    
    If oPlateSystem Is Nothing Then
        'MsgBox "No inputs;  exit from symbol"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Set oMFSystem = arrayOfInputs(2)
    
    lRefSurface = arrayOfInputs(3)
    lMarkingSide = arrayOfInputs(4)
    
    If lRefSurface = 0 Or lMarkingSide = 0 Then
        'The symbol inputs MarkingSide and Reference Surface are not set yet
        Exit Sub
    End If
    
    '--------------------------------------------------
    '  compute the BilgeKeel mark for the given inputs
    '--------------------------------------------------
    Dim oResourceManager    As Object
    Dim oOutPutCollection   As IJDOutputCollection
    Set oOutPutCollection = pOutputColl
    
    Set oResourceManager = oOutPutCollection.ResourceManager
    
    CreateSymbolOuputs oPlateSystem, oMFSystem, lMarkingSide, lRefSurface, oResourceManager, pOutputColl
    
    'TrackText "clean up and return", 0

CleanUp:

    Set oPlateSystem = Nothing
    Set oMFSystem = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Sub

'--------------------------------------------------
'  compute the BilgeKeel mark for the given inputs
'--------------------------------------------------
Private Sub CreateSymbolOuputs(oPlateSystem As IJPlateSystem, oMFSystem As Object, lMarkingSide As Long, lRefSurface As Long, oResourceManager As Object, ByRef pOutputColl As IJDOutputCollection)
On Error GoTo ErrorHandler
Const METHOD = "CreateSymbolOuputs"

    Dim objCodeListMetaData     As IJDCodeListMetaData
    Set objCodeListMetaData = oResourceManager
    
    Dim strMarkingSideProgID    As String
    Dim strRefSurfaceProgID     As String
    Dim i, j                    As Long
    
    '---------------------------------------------------
    ' Get the long string(ProgID) of "MfgMarkingLinesSide"
    '---------------------------------------------------
    
    If Not objCodeListMetaData Is Nothing Then
        strMarkingSideProgID = objCodeListMetaData.LongStringValue("MfgMarkingLinesSide", lMarkingSide)
        strRefSurfaceProgID = objCodeListMetaData.LongStringValue("MfgMarkingLinesSide", lRefSurface)
    End If
    Dim ePlateSide      As enumPlateSide

    Dim oMfgMGhelper            As IJMfgMGHelper
    Set oMfgMGhelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oGeomHelper             As GSCADStrMfgUtilities.MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
       
    'Get all the parts of profile system
    'For each profile part
    '    Find the intersection of web surface and mounting face
    '    Find the plate part associated with profile part and get its surface
    '    Find the web direction
    '    project the wire on plate surface
    'Next
    
    Dim oParts As IJElements
    Set oParts = GetAllPartsFromSystem(oMFSystem)
    
    Dim oProfileSupport As IJProfilePartSupport
    Set oProfileSupport = New ProfilePartSupport
    
    Dim oPartSupport    As IJPartSupport
    Set oPartSupport = oProfileSupport
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oPOM As IJDPOM
    Set oPOM = GetPOM
    
    For i = 1 To oParts.Count
        Dim oPart As Object
        Set oPart = oParts.Item(i)
        
        Dim oPlateRefSurfRule       As IJDMfgPlateUpSideRule
        Dim oProfileRefSurfRule     As IJDMfgProfileUpsideRule

        Dim ConnectionsColl     As Collection
        Dim ConnectedObjColl    As Collection
    
        Dim oStructConnectable  As IJStructConnectable
        Set oStructConnectable = oPart

        Dim ConnectedObjIEUK    As IEnumUnknown
        Dim ConnectionsIEUK     As IEnumUnknown
        oStructConnectable.enumConnectedObjects2 ConnectionPhysical, ConnectedObjIEUK, ConnectionsIEUK
    
        Dim oConvertUtils As New CONVERTUTILITIESLib.CCollectionConversions
        oConvertUtils.CreateVBCollectionFromIEnumUnknown ConnectedObjIEUK, ConnectedObjColl
        
        Set oConvertUtils = Nothing
        Set ConnectedObjIEUK = Nothing
        Set ConnectionsIEUK = Nothing
        
        For j = 1 To ConnectedObjColl.Count
            Dim oConnPlate          As Object
            Set oConnPlate = ConnectedObjColl.Item(j)
            
            Dim oRootParent As Object
            Set oRootParent = oMfgRuleHelper.GetTopMostParentSystem(oConnPlate)
            
            If Not oRootParent Is oPlateSystem Then
                GoTo NextConnPart
            End If
            
            Dim oIntersectionWire   As IJWireBody
            Dim oDirVec             As IJDVector
            
            If TypeOf oPart Is IJPlatePart Then
                On Error Resume Next
                
                Dim ePlateRefSurface    As enumPlateSide
                
                If lRefSurface = 1 Or lRefSurface = 2 Then
                    If lRefSurface = 1 Then
                        ePlateRefSurface = BaseSide
                    Else
                        ePlateRefSurface = OffsetSide
                    End If
                Else
                
                    If TypeOf oMFSystem Is IJPlateSystem Then
                        Set oPlateRefSurfRule = SP3DCreateObject(strRefSurfaceProgID)
                        
                        If oPlateRefSurfRule Is Nothing Then
                            StrMfgLogError Err, MODULE, METHOD, "Please check the Reference surface input", "StrMfgSemanticErrorMessages", 8, , "Symbol"
                            Exit Sub
                        End If
            
                        ePlateRefSurface = oPlateRefSurfRule.GetPlateUpSide(oPart)
                
                    End If
                End If

       
                Dim oPlateSurface       As IJSurfaceBody
                Set oPlateSurface = oGeomHelper.GetSurfaceFromPlate(oPart, TRUE_PART, ePlateRefSurface, 0)
                
                ' First try intersection of plate surface and connected plate base surface
                Dim oConnPlateSurface   As IJSurfaceBody
                Set oConnPlateSurface = oGeomHelper.GetSurfaceFromPlate(oConnPlate, TRUE_PART, BaseSide, 0)
                
                Set oIntersectionWire = GetIntersection(oPlateSurface, oConnPlateSurface)
                If oIntersectionWire Is Nothing Then
                    ' Now try intersection of plate surface and connected plate offset surface
                    Set oConnPlateSurface = Nothing
                    Set oConnPlateSurface = oGeomHelper.GetSurfaceFromPlate(oConnPlate, TRUE_PART, OffsetSide, 0)
                    Set oIntersectionWire = GetIntersection(oPlateSurface, oConnPlateSurface)
                End If
                
            Else
            
                If Not TypeOf oConnPlate Is IJPlatePart Then
                    GoTo NextConnPart
                End If
                
                On Error Resume Next
                
                Dim eProfileFace As PrimaryProfileFace
                
                If lRefSurface = 1 Or lRefSurface = 2 Then
                    If lRefSurface = 1 Then
                        eProfileFace = WebLeftFace
                    Else
                        eProfileFace = WebRightFace
                    End If
                Else
                    If Not TypeOf oMFSystem Is IJPlateSystem Then
                        Set oProfileRefSurfRule = SP3DCreateObject(strRefSurfaceProgID)
                    End If
                    
                    If oProfileRefSurfRule Is Nothing Then
                        StrMfgLogError Err, MODULE, METHOD, "Please check the Reference surface input", "StrMfgSemanticErrorMessages", 8, , "Symbol"
                        Exit Sub
                    End If


                    On Error GoTo ErrorHandler
                    
                    Dim lProfileRefSurface  As Long
                    lProfileRefSurface = oProfileRefSurfRule.GetUpside(oPart) 'JXSEC_WEB_RIGHT (Or) JXSEC_WEB_LEFT
                    
                    If lProfileRefSurface = JXSEC_WEB_LEFT Then
                        eProfileFace = WebLeftFace
                    ElseIf lProfileRefSurface = JXSEC_WEB_RIGHT Then
                        eProfileFace = WebRightFace
                    Else
                        StrMfgLogError Err, MODULE, METHOD, "Please check the Reference surface input", "StrMfgSemanticErrorMessages", 8, , "Symbol"
                        Exit Sub
                    End If
                    
                End If
                
                Dim oStiffenerPart As IJStiffenerPart
                Set oStiffenerPart = oPart
                
                Dim pWebContourColl As Collection, pMonikerColl As Collection
                Dim pWebSB As IJSurfaceBody
                
                Set oPartSupport.Part = oStiffenerPart
                
                oProfileSupport.GetProfileContours eProfileFace, pWebSB, pWebContourColl, pMonikerColl
                
                If pWebSB Is Nothing Then
                    GoTo NextConnPart
                End If
                
                Dim oMountingFacePort As IJPort
                oStiffenerPart.GetMountingFacePort oMountingFacePort, vbNullString
                
                Set oIntersectionWire = GetIntersection(oMountingFacePort, pWebSB)
                
                If Not oIntersectionWire Is Nothing Then
                
                    ' Find endpoints of the LandingCurve
                    Dim oStartPos As IJDPosition
                    Dim oEndPos As IJDPosition
                    oIntersectionWire.GetEndPoints oStartPos, oEndPos
            
                    ' From the start point, get the CrossSection orientation
                    Dim oStartFlangeDirVec As IJDVector
                    Dim oStartWebDirVec As IJDVector
                    Dim oOrigin As IJDPosition
                    oProfileSupport.GetOrientation oStartPos, oStartFlangeDirVec, oStartWebDirVec, oOrigin
                    
                    ' From the end point, get the web direction
                    Dim oEndFlangeDirVec As IJDVector
                    Dim oEndWebDirVec As IJDVector
                    
                    oProfileSupport.GetOrientation oEndPos, oEndFlangeDirVec, oEndWebDirVec, oOrigin
                    
                    Set oDirVec = New DVector
                    
                    oDirVec.x = (oStartWebDirVec.x + oEndWebDirVec.x) / 2
                    oDirVec.y = (oStartWebDirVec.y + oEndWebDirVec.y) / 2
                    oDirVec.z = (oStartWebDirVec.z + oEndWebDirVec.z) / 2
                Else
                    
                    Dim oPortForIntersect As IJPort
                    
                    If eProfileFace = WebLeftFace Then
                        'Try to find intersection of web surface with "JXSEC_WEB_LEFT_BOTTOM_CORNER"
                        Set oPortForIntersect = GetPortGivenSectionID(oPart, JXSEC_WEB_LEFT_BOTTOM_CORNER)
                    ElseIf eProfileFace = WebRightFace Then
                        'Try to find intersection of web surface with "JXSEC_WEB_RIGHT_BOTTOM_CORNER"
                        Set oPortForIntersect = GetPortGivenSectionID(oPart, JXSEC_WEB_RIGHT_BOTTOM_CORNER)
                    End If
                    
                    Set oIntersectionWire = GetIntersection(oPortForIntersect, pWebSB)
                    Set oPortForIntersect = Nothing
                
                    If oIntersectionWire Is Nothing Then 'Try intersection with opposite mounting face port.
                
                        Dim oMountingFaceOppPort As IJPort
                        oStiffenerPart.GetPortOppositeMountingFace oMountingFaceOppPort
                                              
                        Set oIntersectionWire = GetIntersection(oMountingFaceOppPort, pWebSB)
                        Set oMountingFaceOppPort = Nothing
                        
                    End If
                
                End If
                
                Set oMountingFacePort = Nothing
                
            End If
            
        
            If Not oIntersectionWire Is Nothing Then
            
                Dim oCS             As IJComplexString
                oMfgMGhelper.WireBodyToComplexString oIntersectionWire, oCS
                
                Dim oPlatePart      As IJPlatePart
                Set oPlatePart = oConnPlate
                                
                If lMarkingSide = 1 Or lMarkingSide = 2 Then
                    If lMarkingSide = 1 Then
                        ePlateSide = BaseSide
                    Else
                        ePlateSide = OffsetSide
                    End If
                Else
                    Dim oPlateSideRule          As IJDMfgPlateUpSideRule
                    Set oPlateSideRule = SP3DCreateObject(strMarkingSideProgID)
                     
                    ePlateSide = oPlateSideRule.GetPlateUpSide(oPlatePart)
                End If
        
                Dim eSurfaceTypeToUse As eStrMfgSurfaceType
                eSurfaceTypeToUse = TRUE_PART
                
                '---------------------------------------------------------------------
                ' Get the Surface on which remarking line is needed and then project
                '---------------------------------------------------------------------
                '
                Dim oSurfaceBody As IJSurfaceBody
                Set oSurfaceBody = oGeomHelper.GetSurfaceFromPlate(oPlatePart, eSurfaceTypeToUse, ePlateSide, 0)
                
                Dim oProjCSElems As IJElements
                
                oMfgMGhelper.ProjectCSToSurface oCS, oSurfaceBody, oDirVec, oProjCSElems
                
                '--------------------------------------
                'set the output to the complex string
                '--------------------------------------
                
                'TrackText "set symbol output", 0
                
                If Not oProjCSElems Is Nothing Then
                    
                    'Add the complex string to symbol output collection with part OID as output name
                    'This is needed to eleminate the swapping of ref proxies with Marking line AEs
                    
                    Dim strPartOID As String
                    strPartOID = GetOID(oPOM, oPlatePart)
                    
                    'Remove the curly braces from the OID
                    strPartOID = Replace(strPartOID, "{", "")
                    strPartOID = Replace(strPartOID, "}", "")
                    
                    Dim iCount As Long
                    For iCount = 1 To oProjCSElems.Count
                        Dim oCurve As IJCurve
                        Set oCurve = oProjCSElems.Item(iCount)
                        If oCurve.Length > 0.0001 Then
                            Dim oPersistCS As IJComplexString
                            Set oPersistCS = CreatePersistableCS(oProjCSElems.Item(iCount), oResourceManager)
                            InitNewOutput pOutputColl, strPartOID
                            
                            pOutputColl.AddOutput strPartOID, oPersistCS
                        End If
                     Next
                End If
            End If
NextConnPart:
        Next
    Next
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'--------------------------------------------------------------------------------------------------
' Abstract : Find the maximum value of the given two double values
'--------------------------------------------------------------------------------------------------
Private Function MaximumValue(dValue1 As Double, dValue2 As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "MaximumValue"

    If (dValue1 > dValue2) Then
        MaximumValue = dValue1
    Else
        MaximumValue = dValue2
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Find the minimum value of the given two double values
'--------------------------------------------------------------------------------------------------
Private Function MinimumValue(dValue1 As Double, dValue2 As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "MinimumValue"
    If (dValue1 < dValue2) Then
        MinimumValue = dValue1
    Else
        MinimumValue = dValue2
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function


'--------------------------------------------------------------------------------------------------
' Abstract : This method returns the types of Input to the command
'--------------------------------------------------------------------------------------------------

Private Sub IJMfgGenericMarkInputs_GetInputs(InputTypes As JCmnShp_CollectionAlias, DisplayNames As JCmnShp_CollectionAlias, Prompts As JCmnShp_CollectionAlias, DefaultValues As JCmnShp_CollectionAlias, InputOptionalInfo As JCmnShp_CollectionAlias)
On Error GoTo ErrorHandler
Const METHOD = "IJMfgGenericMarkInputs_GetInputs"
    'This method returns 4 collections which give information about
        ' collection 1 - Type of the Input - object, codelistedvalue, simple data type, dimensioned value
        ' collection 2 - strings that are to be shown as prompt when the object is selected
        ' collection 3 - Default values/filter criteria for the input objects.
        ' collection 4 - Info about Optional input or not
        
    Set InputTypes = New Collection
    Set Prompts = New Collection
    Set DefaultValues = New Collection
    Set InputOptionalInfo = New Collection
    
    Dim strCriteria As String
    
    '------------
    'PlateSystem
    '------------
    InputTypes.Add ObjectType
    Prompts.Add m_strPlateSysPrompt
    DisplayNames.Add m_strINPUT_1
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_1)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '--------------------------
    'PlateSystem/ProfileSystem
    '--------------------------
    InputTypes.Add ObjectType
    Prompts.Add m_strMFSysPrompt
    DisplayNames.Add m_strINPUT_2
    'get filter criteria for the selection of connected Plate/Profile System
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_2)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '------------------
    'Reference Surface
    '------------------
    InputTypes.Add CodeListedValue
    Prompts.Add m_strRefSurfacePrompt
    DisplayNames.Add m_strRefSurface
    'get filter criteria for the selection of connected Plate/Profile System
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_3)
    DefaultValues.Add strCriteria '"MfgMarkingSide:2110"
    InputOptionalInfo.Add False 'False means Not optional
    
    '-------------
    'Marking Side
    '-------------
    InputTypes.Add CodeListedValue
    Prompts.Add m_strMarkingSide
    DisplayNames.Add m_strMarkingSide
    'get filter criteria for the selection of connected Plate/Profile System
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_4)
    DefaultValues.Add strCriteria '"MfgMarkingSide:1111"
    InputOptionalInfo.Add False 'False means Not optional
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'--------------------------------------------------------------------------------------------------
' Abstract : Method to get the filter criteria or the default value of an input
'--------------------------------------------------------------------------------------------------

Private Function GetFilterCriteriaOrDefaultvalue(strInputString As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetFilterCriteriaOrDefaultvalue"
    
    Select Case strInputString
        Case INPUT_1
            GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_2
           GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_3
            GetFilterCriteriaOrDefaultvalue = "MfgMarkingLinesSide:2110" 'Web Left
        Case INPUT_4
            GetFilterCriteriaOrDefaultvalue = "MfgMarkingLinesSide:1111" 'Molded Side
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Method to get the filter criteria for an input
'--------------------------------------------------------------------------------------------------
Private Function GetFilterCriteria(strInputString As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetFilterCriteria"

    Select Case strInputString
        Case INPUT_1
            'PlateSystem
            GetFilterCriteria = IJPlateSystem & " AND " & " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] "
        Case INPUT_2
            'Connected Plate/Profile System
            GetFilterCriteria = IJPlateSystem & " AND " & " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPlateSystem] " & " AND " & " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] " _
                                & " OR " & IJStiffenerSystem & " AND " & " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPlateSystem] " & " AND " & " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] "
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

Private Sub InitializeLocalizeVariables()
    Const METHOD = "InitializeLocalizeVariables"
    On Error GoTo ErrorHandler
    
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = InitializeLocalizer(App.Path & "\Resource\", App.EXEName)

    m_strINPUT_1 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_1, "Plate System")
    m_strINPUT_2 = oLocalizer.GetString(IDS_BILGEKEELMARK_SYMBOL_INPUT_2, "Plate/Profile System")
    m_strRefSurface = oLocalizer.GetString(IDS_BILGEKEELMARK_SYMBOL_REFSURFACE, "Reference Surface")
    m_strMarkingSide = oLocalizer.GetString(IDS_BILGEKEELMARK_SYMBOL_MARKINGSIDE, "Marking Side")
    
    m_strPlateSysPrompt = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_PLATESYSTEM_PROMPT, "Select a Plate System")
    m_strMFSysPrompt = oLocalizer.GetString(IDS_BILGEKEELMARK_SYMBOL_MF_SYSTEM_PROMPT, "Select a connected Plate or Profile System")
    m_strRefSurfacePrompt = oLocalizer.GetString(IDS_BILGEKEELMARK_SYMBOL_REFSURFACE_PROMPT, "Select connected part's reference surface")
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

Private Function InitializeLocalizer(strPath As String, strFileName As String) As IJLocalizer
    Const METHOD = "InitializeLocalizer"
    On Error GoTo ErrorHandler
     
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = CreateObject("IMSLocalizer.Localizer")
    
    If IsInDebugMode = True Then
      oLocalizer.Initialize "S:\StructManufacturing\Client\Bin\Resource\" & strFileName
    Else
       oLocalizer.Initialize strPath & strFileName
    End If
    
    Set InitializeLocalizer = oLocalizer
Exit Function
ErrorHandler:
   Err.Clear
End Function

Private Function IsInDebugMode() As Boolean
    On Error GoTo ErrorHandler
'     If the program is compiled, the following
'     Debug statement has been removed so it will
'     not generate an error.
    Debug.Print 1 / 0
    IsInDebugMode = False
    Exit Function
ErrorHandler:
'   We got an error so the Debug statement must
'   be working.
    IsInDebugMode = True
   Err.Clear
End Function

'******************************************************************************
' Routine: IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs
'
' Abstract: Get the parts or systems to be marked from the symbol inputs
'******************************************************************************
Private Function IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs(ByVal oSymbolInputColl As JCmnShp_CollectionAlias) As IJDObjectCollection
    Const METHOD = "IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs"
    On Error GoTo ErrorHandler

    Dim oPlateSystem As IJPlateSystem
    
    ' The first input in the symbol is plate system
    Set oPlateSystem = oSymbolInputColl.Item(1)
    
    If Not oPlateSystem Is Nothing Then
        Dim oConnPlatesColl         As IJDObjectCollection
        Set oConnPlatesColl = New JObjectCollection
        oConnPlatesColl.Add oPlateSystem
        
        Set IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs = oConnPlatesColl
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

Private Function IJMfgGenericMarkInputs_GetValuesInRange(ByVal pDispObj1 As Object, ByVal pDispObj2 As Object) As IJElements

End Function
Private Function GetPortGivenSectionID(Part As Object, lSectionID As Long) As Object

    Const METHOD = "GetPortGivenSectionID"
    On Error GoTo ErrorHandler
    
    Dim iCnt                As Integer
    Dim oPort               As IJPort
    Dim oStructPort         As IJStructPort
    Dim oEnumPorts          As IJElements
    Dim oStructConnectable  As IJStructConnectable
    
    Set oEnumPorts = New JObjectCollection
    Set oStructConnectable = Part
    
    oEnumPorts.AddElements oStructConnectable.enumAllPorts
    
    For iCnt = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(iCnt)
        Set oStructPort = oPort
         
        If oStructPort.SectionID = lSectionID Then
            Set GetPortGivenSectionID = oPort
            Exit For
        End If
    Next
    
     
Exit Function
ErrorHandler:
   Err.Clear
End Function

