VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotIDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgId

Const m_CAName As String = "SlotIDef"                   ' - TODO
Const m_CAProjectName As String = "SlotRules"            ' - TODO

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\SlotIDef.cls"

'Newly created member objects
Private m_oPhysConn As Object

'Error management
Private sERROR As String
Private sMETHOD As String

Public Sub ItemInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
  On Error GoTo ErrorHandler
  
  pAD.SetCMFinalConstruct "SlotRules.SlotIDef", "CMFinalConstructSlot"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler

    pMDs.AddMember "SlotIPC1", 1, "CMConstructLeftPC", "SlotRules.SlotIDef"
    pMDs.AddMember "SlotIPC2", 2, "CMConstructRightPC", "SlotRules.SlotIDef"

    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("SlotIFET1", 3, "CMConstructFET1", CMLIBRARY_SLOTDEFCM)
    pMemDesc.SetCMConditional CMLIBRARY_SLOTDEFCM, "CMGetEdgeTreatment"

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructSlot(ByVal pAggregatorDescription As IJDAggregatorDescription)
  Dim clsSlotCM As New SlotDefCM
  clsSlotCM.Slot_FinalConstruct pAggregatorDescription
  Set clsSlotCM = Nothing
  
  Set m_oPhysConn = Nothing
End Sub
Public Sub CMConstructLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsSlotCM As New SlotDefCM
  Set m_oPhysConn = clsSlotCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, "TeeWeld", _
                                                    JXSEC_WEB_LEFT, JXSEC_WEB_LEFT)
  Set pObject = m_oPhysConn
  Set clsSlotCM = Nothing
End Sub
Public Sub CMConstructRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsSlotCM As New SlotDefCM
  Set m_oPhysConn = clsSlotCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, "TeeWeld", _
                                                   JXSEC_WEB_RIGHT, JXSEC_WEB_RIGHT)
  Set pObject = m_oPhysConn
  Set clsSlotCM = Nothing
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
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

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
  
  Exit Function
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "IJDUserSymbolServices_InstanciateDefinition").Number
End Function
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
  On Error GoTo ErrorHandler
  sMETHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
  sERROR = ""
  
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
  
  ' define the aggregator
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  ItemAggregator pAD
   
  ' define the members
  Dim pMDs As IJDMemberDescriptions
  Set pMDs = pDefinition
  ItemMembers pMDs
  
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function
