VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToEndChPCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID

Const m_CAName As String = "EndToEndChPCDef"                   ' - TODO
Const m_CAProjectName As String = CUSTOMERID + "EndCutRules"            ' - TODO

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

' Special user defined Edge ID used in the WebCut Symbol files
Const JXSEC_FLANGE_CUT_BY_WEBCUT = 8194

Private Const MODULE = "\StructDetail\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToEndChPCDef.cls"

'Newly created member objects
Private m_oPhysConn As Object

'Error management
Private sError As String
Private sMETHOD As String

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    pAD.UserTypeClsid = CA_WEBCUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("PC1", 1, "CMConstructPC1", _
                                  imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "ConstructPC1_Con"
    Set pMemDesc = Nothing
    
    ' Used for Second Physical Connection for SeamAngleCase#
    ' (between Flange cut by Web)
    Set pMemDesc = pMDs.AddMember("PC2", 2, "CMConstructPC2", _
                                  imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "ConstructPC2_Con"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("EndToEndPhyConnFET", 3, "CMConstructFET", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMGetEdgeTreatment"
    Set pMemDesc = Nothing
    ' Added new item meber for creating chamfers
    Set pMemDesc = pMDs.AddMember("ChamferOnER", 4, "CMConstructChamferOnER", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetChamferType"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMIsChamferOnERNeeded"
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructWebCut(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Dim clsWebCutCM As New WebCutDefCM
    
    clsWebCutCM.WebCut_FinalConstruct pAggregatorDescription
    
    Set clsWebCutCM = Nothing
    Set m_oPhysConn = Nothing
End Sub

Public Sub CMConstructPC1(ByVal pMemberDescription As IJDMemberDescription, _
                          ByVal pResourceManager As IUnknown, _
                          ByRef pObject As Object)
    
    Dim clsWebCutCM As WebCutDefCM
    Set clsWebCutCM = New WebCutDefCM
            
zMsgBox "EndToEndChPCDef::CMConstructPC1"

    Set m_oPhysConn = clsWebCutCM.CAConstruct_PhysConn(pMemberDescription, _
                                                pResourceManager, _
                                                "ButtWeld", _
                                                JXSEC_IDEALIZED_BOUNDARY, _
                                                JXSEC_IDEALIZED_BOUNDARY)
    Set pObject = m_oPhysConn
    Set clsWebCutCM = Nothing
End Sub

Public Sub CMConstructPC2(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    
    Dim clsWebCutCM As WebCutDefCM
    Set clsWebCutCM = New WebCutDefCM
            
zMsgBox "EndToEndChPCDef::CMConstructPC2"
    ' The Flange Portion cut by the Web is represented by "8194"
    Set m_oPhysConn = clsWebCutCM.CAConstruct_PhysConn(pMemberDescription, _
                                                pResourceManager, _
                                                "ButtWeld", _
                                                JXSEC_IDEALIZED_BOUNDARY, _
                                                JXSEC_FLANGE_CUT_BY_WEBCUT)
    Set pObject = m_oPhysConn
    Set clsWebCutCM = Nothing
End Sub


' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_CAProgid
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation
    
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers pMDs
  
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_CAProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'***********************************************************************
' METHOD:  CMSetChamferType
'
' DESCRIPTION:  Method to copy Chamfer Type Question from AC to Chamfer Rules
'
'***********************************************************************

Public Sub CMSetChamferType(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String

    Dim sChamferType As String
    Dim oSmartParent As Object

    GetSmartOccurrenceParent pMemberDescription.CAO, oSmartParent
    
    GetSelectorAnswer oSmartParent, "ChamferType", sChamferType

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.PutAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "ChamferType", sChamferType

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetChamferType", strError).Number
End Sub

'***********************************************************************
' METHOD:  CMIsChamferOnERNeeded
'
' DESCRIPTION:  Conditional that determines if the Chamfer is required
'
'***********************************************************************

Public Sub CMIsChamferOnERNeeded(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Chamfer) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oWebCutDefCM As New WebCutDefCM
    bIsNeeded = oWebCutDefCM.IsChamferExpected(pMD)
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsChamferOnERNeeded", strError).Number
End Sub

'***********************************************************************
' METHOD:  CMConstructChamferOnER
'
' DESCRIPTION:  Method to construct the chamfer creation on Edge Reinforcements
'
' Thid definition file is used for creating chamfers on end cuts with end cut edge id 8193 in symbol files
' If the edge id is other than 8193 this definition file is not used.
'***********************************************************************

Public Sub CMConstructChamferOnER(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    Dim oCreateChamfer As GSCADSDCreateModifyUtilities.SDChamferUtils
    Set oCreateChamfer = New GSCADSDCreateModifyUtilities.SDChamferUtils
    
    Dim pWebCut As StructDetailObjects.WebCut
    Set pWebCut = New StructDetailObjects.WebCut
    Set pWebCut.object = pMemberDescription.CAO
    
    Set m_oPhysConn = oCreateChamfer.CreateChamferAtEndCut(pResourceManager, pWebCut.object, 8193, "RootChamfer", pMemberDescription.CAO)
    Set pObject = m_oPhysConn
    
    InitializeStructFeatureProperties pObject
    
    Set oCreateChamfer = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructChamferOnER", strError).Number
End Sub



