VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToLongBoxWebCutDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID
Const m_CAName As String = "EndToLongBoxWebCutDef"
Const m_CAProjectName As String = "EndCutRules"

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\EndToLongBoxWebCutDef.cls"

'Newly created member objects
Private m_oPhysConn As Object
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    pAD.UserTypeClsid = CA_WEBCUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    ' Set Q/A to create free edge treatments
    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("WebCutsFET", 1, "CMConstructFET", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMGetEdgeTreatment"

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

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                    ByVal defParams As Variant, _
                                    ByVal ActiveConnection As Object) As Object
Const sMETHOD = "IJDUserSymbolServices_InstanciateDefinition"
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
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
Const sMETHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
      
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
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function



