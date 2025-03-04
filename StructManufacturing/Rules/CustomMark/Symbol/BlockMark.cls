VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: MfgGenericMarkSymbol
'
'File: BlockMark
'
'Abstract: Implementation of BlockMark mark symbol
'
'Author: Siva
'
'Creation date: July 20, 2010
'
'Revision:
'*******************************************************************

Option Explicit

Implements IJDUserSymbolServices
Implements IJMfgGenericMarkInputs

Private Const MODULE = "MfgGenericMarkSymbol.BlockMark(BlockMark.cls)"
Private Const m_progID = "MfgGenericMarkSymbol.BlockMark"

'---------------------------------------------------
'   INPUT Names
'---------------------------------------------------
Private Const INPUT_1 = "PinJig"
Private Const INPUT_2 = "SuppPlatesByUser"
Private Const INPUT_3 = "OffsetValue"
Private Const INPUT_4 = "SuppPlatesByType"

'---------------------------------------------------
'   OUTPUT and REPRESENTATION Names
'---------------------------------------------------

Private Const REPRESENTATION1 = "Simple" ' used to named the simple physical representation

Private m_strINPUT_1                 As String
Private m_strINPUT_2                 As String
Private m_strINPUT_3                 As String
Private m_strINPUT_4                 As String

Private m_strPinJigPrompt            As String
Private m_strOffsetValPrompt         As String
Private m_strSuppPlatesByType        As String
Private m_strSuppPlatesByUser        As String

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
    
    '-------------
    'Define inputs
    '-------------
    
    'get the input interface of the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = ppSymbolDef

    '------------
    'PinJig Input
    '------------
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput

    oInput.name = INPUT_1
    oInput.Description = INPUT_1
    oInput.Properties = 0
    oInputs.SetInput oInput, 1
    oInput.Reset
    
    Set oInput = Nothing
    
    '----------------------------------
    'Supported plates by User selection
    '----------------------------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_2
    oInput.Description = INPUT_2
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInputs.SetInput oInput, 2
    oInput.Reset

    Set oInput = Nothing
    
    '------------------
    'Offset value Input
    '------------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_3
    oInput.Description = INPUT_3
    oInput.Properties = igINPUT_IS_A_PARAMETER
    
    'create an input parameter description
    Dim oPC As IJDParameterContent
    
    'create an input parameter description
    Set oPC = New IMSSymbolEntities.DParameterContent
  
    oPC.Type = igValue
    oPC.UomValue = 0

    oInput.DefaultParameterValue = oPC
    oInputs.SetInput oInput, 3
    oInput.Reset

    Set oInput = Nothing
    Set oPC = Nothing
        
    '------------------------
    'Supported plates by Type
    '------------------------
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
    
    '------------------
    'Define evaluations
    '------------------
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
    
    Dim oSuppPlatesByUser   As IJDObjectCollection
    Dim oPinJig             As Object
    Dim oRefColl            As Object
    Dim dOffsetValue        As Double
    Dim lSuppPlatesByType   As Long
    
    '------------------
    'gather all inputs
    '------------------
    'TrackText "gather all inputs", 0
    
    'get the required inputs
    Set oPinJig = arrayOfInputs(1)
    
    If oPinJig Is Nothing Then
        'MsgBox "No inputs;  exit from symbol"
        Exit Sub
    End If
    
    If TypeOf arrayOfInputs(2) Is Object  Then
        Set oRefColl = arrayOfInputs(2)         ' collection of supported plates by user
        dOffsetValue = arrayOfInputs(3)         ' Offset value
        lSuppPlatesByType = arrayOfInputs(4)    ' Supported plates by type
    Else
        dOffsetValue = arrayOfInputs(2)         ' Offset value
        lSuppPlatesByType = arrayOfInputs(3)    ' Supported plates by type
        
        ' If the Supported plates by user selection is selected and if no plates are selected then exit with message logging
        If lSuppPlatesByType = 4 Then
            StrMfgLogError Err, MODULE, METHOD, "Supported Plates by Type is 'User Selection', Please select the plates to be marked manually", "StrMfgSemanticErrorMessages", 8, , "Symbol"
            Exit Sub
        End If
        
    End If
    
    If dOffsetValue = 0 Or lSuppPlatesByType = 0 Then
        'The symbol inputs OffsetValue and SuppPlatesByType are not set yet
        Exit Sub
    End If
    
    ' If the Supported plates by user selection is selected then take them into a collection
    If lSuppPlatesByType = 4 Then
        
        Dim oRefCollArgs            As IEnumJDArgument
        
        Set oSuppPlatesByUser = New JObjectCollection
        Set oRefCollArgs = oRefColl
        
        Dim arg1    As IJDArgument
        Dim found   As Long
        
        Do
            oRefCollArgs.Next 1, arg1, found
            If found = 0 Then Exit Do
            
            oSuppPlatesByUser.Add arg1.Entity, arg1.Index
            
            Set arg1 = Nothing
        Loop
    End If
    
    '--------------------------------------------------
    '  compute the Block mark for the given inputs
    '--------------------------------------------------
    Dim oResourceManager    As Object
    Dim oOutPutCollection   As IJDOutputCollection
    Set oOutPutCollection = pOutputColl

    Set oResourceManager = oOutPutCollection.ResourceManager
    
    CreateSymbolOuputs oPinJig, dOffsetValue, lSuppPlatesByType, oSuppPlatesByUser, oResourceManager, pOutputColl
    
    Set oPinJig = Nothing
    
    If Not oSuppPlatesByUser Is Nothing Then
        oSuppPlatesByUser.Clear
        Set oSuppPlatesByUser = Nothing
    End If
    
    'TrackText "clean up and return", 0

CleanUp:

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Sub

'--------------------------------------------------
'  compute the Block mark for the given inputs
'--------------------------------------------------
Private Sub CreateSymbolOuputs(oPinJig As IJPinJig, dOffsetValue As Double, lSuppPlatesByType As Long, oSuppPlatesByUser As IJDObjectCollection, oResourceManager As Object, ByRef pOutputColl As IJDOutputCollection)
On Error GoTo ErrorHandler
Const METHOD = "CreateSymbolOuputs"

    Dim strPlatesType       As String
    Dim iCount              As Long
    Dim ePlatesType         As StructPlateType
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oConnPlatesForMarking As IJElements
    Set oConnPlatesForMarking = New JObjectCollection
    
    If lSuppPlatesByType = 4 Then
        Dim oObject  As Object
        For Each oObject In oSuppPlatesByUser
            oConnPlatesForMarking.Add oObject
        Next
    Else
        Dim oPinJigSuppPlates As IJElements
        Set oPinJigSuppPlates = oPinJig.SupportedPlates
        
        ' Get the eligible connected plates from SuppPlatesByType option
        For iCount = 1 To oPinJigSuppPlates.Count
            Dim oTempColl As IJElements
            Set oTempColl = GetPlateConnectedObjectsByType(oPinJigSuppPlates.Item(iCount), lSuppPlatesByType)
            oConnPlatesForMarking.AddElements oTempColl
        Next
        
    End If
    
    ' Get the Pin Jig base plane information i.e., root point and normal
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    oPinJig.GetBasePlane dNormalX, dNormalY, dNormalZ, dRootX, dRootY, dRootZ
    
    Dim dAdjRootX As Double, dAdjRootY As Double, dAdjRootZ As Double
    
    ' Adjust the Base plane by adding offset to the root point in normal direction
    dAdjRootX = dRootX + dNormalX * dOffsetValue
    dAdjRootY = dRootY + dNormalY * dOffsetValue
    dAdjRootZ = dRootZ + dNormalZ * dOffsetValue
    
    Dim oOffsetBasePlane   As IJPlane
    Set oOffsetBasePlane = New Plane3d
    
    ' Set the new root point and normal on Base plane
    oOffsetBasePlane.SetRootPoint dAdjRootX, dAdjRootY, dAdjRootZ
    oOffsetBasePlane.SetNormal dNormalX, dNormalY, dNormalZ
    
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    
    Dim oPOM As IJDPOM
    Set oPOM = GetPOM
    
    For iCount = 1 To oConnPlatesForMarking.Count
    
        ' Get the base surface of the plate part
        Dim oSurfaceBody As IJSurfaceBody
        Set oSurfaceBody = oMfgGeomHelper.GetSurfaceFromPlate(oConnPlatesForMarking.Item(iCount), TRUE_PART, PlateBaseSide, 0)
        
        Dim oIntersectionCurve As Object
        
        ' Intersect each plate base surface with the adjusted base plane
        oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oOffsetBasePlane, oIntersectionCurve, oStartPos, oEndPos
    
        If Not oIntersectionCurve Is Nothing Then
        
            ' Convert the intersection wire body into complex strings
            Dim oCSElems As IJElements
            Dim oMfgMGhelper As New MfgMGHelper
            oMfgMGhelper.WireBodyToComplexStrings oIntersectionCurve, oCSElems
            
            '--------------------------------------
            'set the output to the complex string
            '--------------------------------------
            
            'TrackText "set symbol output", 0
            
            ' Add each complex string to the symbol output if it's length is more than 1 mm
            If Not oCSElems Is Nothing Then
                
                'Add the complex string to symbol output collection with part OID as output name
                'This is needed to eleminate the swapping of ref proxies with Marking line AEs
                Dim strPartOID As String
                strPartOID = GetOID(oPOM, oConnPlatesForMarking.Item(iCount))
                
                'Remove the curly braces from the OID
                strPartOID = Replace(strPartOID, "{", "")
                strPartOID = Replace(strPartOID, "}", "")
                
                Dim jCount As Long
                For jCount = 1 To oCSElems.Count
                    Dim oCurve As IJCurve
                    Set oCurve = oCSElems.Item(jCount)
                    If oCurve.Length > 0.0001 Then
                        Dim oPersistCS As IJComplexString
                        Set oPersistCS = CreatePersistableCS(oCSElems.Item(jCount), oResourceManager)
                        InitNewOutput pOutputColl, strPartOID
                        
                        pOutputColl.AddOutput strPartOID, oPersistCS
                    End If
                Next
            End If
        End If
        
    Next
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'******************************************************************************
' Routine: GetPlateConnectedObjects
'
' Abstract: Used to get the connected objects to the plate part based on input plateType given
'******************************************************************************
Private Function GetPlateConnectedObjectsByType(oElement As Object, lPlatesType As Long) As IJElements
    Const METHOD = "GetPlateConnectedObjects"
    On Error GoTo ErrorHandler
    
    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oPartSupport As IJPartSupport
    Dim oConnPlatesColl As IJElements

    Set oPlatePartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport = oPlatePartSupport
    Set oPartSupport.Part = oElement
    
    Set oConnPlatesColl = New JObjectCollection
    
    ' Get the PC based connected objects
    Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol
    
    Dim lNoOfObj As Long, iCount As Long
    lNoOfObj = oConObjCol.Count
    
    ' If the connected object satisfies that plate type given then add it to a collection and return
    For iCount = 1 To lNoOfObj
        If TypeOf oConObjCol.Item(iCount) Is IJPlatePart Then
        
            Dim oSDPlatePart As New StructDetailObjects.PlatePart
            Set oSDPlatePart.object = oConObjCol.Item(iCount)
            
            If lPlatesType = 1 And oSDPlatePart.plateType = TBulkheadPlate Then  'lPlatesType = 1 means Transversal
                oConnPlatesColl.Add oConObjCol.Item(iCount)
            ElseIf lPlatesType = 2 And oSDPlatePart.plateType = LBulkheadPlate Then  'lPlatesType = 2 means Longitudinal
                oConnPlatesColl.Add oConObjCol.Item(iCount)
            ElseIf lPlatesType = 3 And (oSDPlatePart.plateType = TBulkheadPlate Or oSDPlatePart.plateType = LBulkheadPlate) Then 'lPlatesType = 3 means Transversal and Longitudinal
                oConnPlatesColl.Add oConObjCol.Item(iCount)
            End If
        End If
    Next
    
    Set GetPlateConnectedObjectsByType = oConnPlatesColl
    
    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    Set oConObjCol = Nothing
    Set oConnCol = Nothing
    Set oThisPortCol = Nothing
    Set oOtherPortCol = Nothing
    
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
    
    ' Add the inputs so that same order is respected in showing on the marking by definition form
    '--------
    'Pin Jig
    '--------
    InputTypes.Add ObjectType
    Prompts.Add m_strPinJigPrompt
    DisplayNames.Add m_strINPUT_1
    'get filter criteria for the selection of Pin Jig
    strCriteria = GetFilterCriteriaOrDefaultvalue(m_strINPUT_1)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '-------------
    'Offset value
    '-------------
    InputTypes.Add DimensionedValue
    Prompts.Add m_strOffsetValPrompt
    DisplayNames.Add m_strINPUT_2
    strCriteria = ""
    DefaultValues.Add "2.5"
    InputOptionalInfo.Add False 'False means Not optional
    
    '-------------------------
    'Supported Plates by Type
    '-------------------------
    InputTypes.Add CodeListedValue
    Prompts.Add m_strSuppPlatesByType
    DisplayNames.Add m_strINPUT_3
    strCriteria = GetFilterCriteriaOrDefaultvalue(m_strINPUT_3) '"StrMfgMarkingConnPlatesType:1" 'Transversal
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '-----------------------------------
    'Supported Plates by User Selection
    '-----------------------------------
    InputTypes.Add MultiObjectType
    Prompts.Add m_strSuppPlatesByUser
    DisplayNames.Add m_strINPUT_4
    'get filter criteria for the selection of Supported Plates (by User)
    strCriteria = GetFilterCriteriaOrDefaultvalue(m_strINPUT_4)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add True 'True means optional
    
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
        Case m_strINPUT_1
            GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case m_strINPUT_3
            GetFilterCriteriaOrDefaultvalue = "StrMfgMarkingConnPlatesType:1" 'Transversal
        Case m_strINPUT_4
            GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        
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
        Case m_strINPUT_1
            'Pin Jig
            GetFilterCriteria = IJPinJig
        Case m_strINPUT_4
            'Connected Plates to the supported plates of Pin Jig
            GetFilterCriteria = IJPlatePart & " AND " & " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPinJigSuppPlates] "
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
    
    m_strINPUT_1 = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_INPUT_1, "Pin Jig")
    m_strINPUT_2 = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_INPUT_2, "Offset Value")
    m_strINPUT_3 = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_INPUT_3, "Supported Plates by Type")
    m_strINPUT_4 = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_INPUT_4, "Supported Plates by User")

    m_strPinJigPrompt = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_PINJIG_PROMPT, "Select a Pin Jig")
    m_strOffsetValPrompt = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_OFFSETVALUE_PROMPT, "Enter Offset Value")
    m_strSuppPlatesByType = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_SUPPPLATESBYTYPE_PROMPT, "Select Supported Plates by Type")
    m_strSuppPlatesByUser = oLocalizer.GetString(IDS_BLOCKMARK_SYMBOL_SUPPPLATESBYUSER_PROMPT, "Select Plates manually: Allowed only if Supported Plates by Type is 'User Selection'")
    
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

Private Function IJMfgGenericMarkInputs_GetValuesInRange(ByVal pDispObj1 As Object, ByVal pDispObj2 As Object) As IJElements

End Function

'******************************************************************************
' Routine: IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs
'
' Abstract: Get the parts or systems to be marked from the symbol inputs
'******************************************************************************
Private Function IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs(ByVal oSymbolInputColl As JCmnShp_CollectionAlias) As IJDObjectCollection
    Const METHOD = "IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs"
    On Error GoTo ErrorHandler

    Dim oConnPlatesColl         As IJDObjectCollection
    Dim strPlatesType           As String
    
    ' The third input in the symbol indicates the "Supported plates by Type"
    strPlatesType = oSymbolInputColl.Item(3)
    
    If strPlatesType = "User Selection" Then
        Set oConnPlatesColl = oSymbolInputColl.Item(4)
    Else
    
        Dim oPinJig             As IJPinJig
        Set oPinJig = oSymbolInputColl.Item(1)
        
        Set oConnPlatesColl = New JObjectCollection
        
        Dim oPinJigSuppPlates   As IJElements
        Set oPinJigSuppPlates = oPinJig.SupportedPlates
        
        Dim lSuppPlatesByType As Long
        
        If strPlatesType = "Transversal" Then
            lSuppPlatesByType = 1
        ElseIf strPlatesType = "Longitudinal" Then
            lSuppPlatesByType = 2
        Else 'Transversal and Longitudinal
            lSuppPlatesByType = 3
        End If
           
        Dim iCount As Long
        For iCount = 1 To oPinJigSuppPlates.Count
            Dim oTempColl As IJElements
            
            ' Get the connected plates based on the SuppPlatesByType value
            Set oTempColl = GetPlateConnectedObjectsByType(oPinJigSuppPlates.Item(iCount), lSuppPlatesByType)
            
            Dim jCount As Long
            For jCount = 1 To oTempColl.Count
                If Not oConnPlatesColl.Contains(oTempColl.Item(jCount)) Then
                    oConnPlatesColl.Add oTempColl.Item(jCount)
                End If
            Next
        Next
        
        Set oPinJig = Nothing
        
    End If
    
    If oConnPlatesColl.Count > 0 Then
        Set IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs = oConnPlatesColl
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function
