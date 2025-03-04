VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffenerThruPlateDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.StiffenerThruPlateDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.StiffenerThruPlateDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffenerThruPlateDef.bas"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs", strError).Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim pMD As IJDMemberDescription
    
    strError = "Constructing slot."
    Set pMD = pMDs.AddMember("StiffThruPlateSlot1", 1, "CMConstructSlot", imsCOOKIE_ID_USS_LIB)
    pMD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsSlot"
    pMD.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateSlot"
    Set pMD = Nothing

    ' Manual Slot
    ' add Conditional create for Primary Clip
    strError = "Constructing collar1."
    Set pMD = pMDs.AddMember("StiffThruPlateCollar1", 2, "CMConstructCollar", imsCOOKIE_ID_USS_LIB)
    pMD.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateClip_Primary"
    pMD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsCollar"
    Set pMD = Nothing
    
   
    ' Manual Slot
    ' add Conditional create for Secondary Clip
    strError = "Constructing collar2"
    Set pMD = pMDs.AddMember("StiffThruPlateCollar2", 3, "CMConstructCollar", imsCOOKIE_ID_USS_LIB)
    pMD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetCollar2Inputs"
    pMD.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateClip_Secondary"
    Set pMD = Nothing
   
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
    
End Sub

' ** Start CM **
Public Sub CMConstructCollar(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    Dim pMemberObjects As IJDMemberObjects
    
    strError = "Retrieving CAO and constructing collar."
    Set pMemberObjects = pMemberDescription.CAO
    Set pObject = clsAssyConnCM.CMConstruct_Collar(pMemberObjects.Item(1), "RootClip")
    
    Set pMemberObjects = pMemberDescription.CAO
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructCollar", strError).Number
End Sub
Public Sub CMSetCollar2Inputs(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    strError = "Copying questions for collar 2"
    CMCopyQuestionsCollar pMemberDescription
      
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set Collar order on the collar"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer CUSTOMERID + "CollarRules.RootClipSel", "CollarCreationOrder", "Secondary"

    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetCollar2Inputs", strError).Number
End Sub

Public Sub CMConstructSlot(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Retrieving CAO and constructing slot."
    Set pObject = clsAssyConnCM.CMConstruct_Slot(pMemberDescription, pResourceManager, "RootSlot")
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructSlot", strError).Number
End Sub

Public Sub CMUpdateThickness(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating Thickness property."
    Call clsAssyConnCM.CMCopyParentThickness(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMUpdateThickness", strError).Number
End Sub

Public Sub CMUpdateMatAndGrade(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating MaterialAndGrade property."
    Call clsAssyConnCM.CMCopyParentMatAndGrade(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMUpdateMatAndGrade", strError).Number
End Sub

Public Sub CMSetSideOfPlate(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating SideOfPlate property."
    Call clsAssyConnCM.CMSetSideOfPlate(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetSideOfPlate", strError).Number
End Sub

Public Sub CMCopyQuestionsCollar(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Calling Copy answer to copy assembly questions to collar"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "CollarRules.RootClipSel", "StressLevel", _
                                 CUSTOMERID + "AssyConRul.PenetrationConnSel", "StressLevel"

                                
    Set pCopyAnswerHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMemebers", strError).Number
End Sub
Public Sub CMCopyQuestionsSlot(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Calling Copy answer to copy assembly questions to Slot"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "SlotRules.SlotRootSel", "AssyMethod", _
                                 CUSTOMERID + "AssyConRul.PenetrationConnSel", "AssyMethod"
                                 
    Set pCopyAnswerHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsSlot", strError).Number
End Sub


' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
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
    pIH.Definition = pDefinition
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
    pDefinition.ProgId = m_ItemProgid
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

