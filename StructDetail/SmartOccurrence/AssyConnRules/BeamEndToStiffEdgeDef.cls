VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BeamEndToStiffEdgeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "AssyConnRules.BeamEndToStiffEdgeDef"
Const m_ItemName As String = "AssyConnRules.BeamEndToStiffEdgeDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\BeamEndToStiffEdgeDef.cls"

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

    Dim pMemDesc As IJDMemberDescription
    strError = "Constructing web cut."
    Set pMemDesc = pMDs.AddMember("BeamEndToStiffEdgeWebCut1", 1, _
                                  "CMConstruct_WebCut", CMLIBRARY_ASSYCONNRULES)
    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToMembers"
    Set pMemDesc = Nothing
    
    strError = "Constructing flange cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToStiffEdgeFlangeCut1", 2, _
                                  "CMConstruct_FlangeCut", CMLIBRARY_ASSYCONNRULES)
        
    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyFCQuestionsToMembers"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateTopFlangeCut"
    
    strError = "Constructing flange cut 2."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToStiffEdgeFlangeCut2", 3, _
                                  "CMConstruct_BottomFlangeCut", CMLIBRARY_ASSYCONNRULES)
        
    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetBottomFlange"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateBottomFlangeCut"
    
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Constructing physical connection."
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnFromACPorts(pMemberDescription, pResourceManager, "ButtWeld")
    
    Set clsAssyConnCM = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC", strError).Number
End Sub

Public Sub CMCopyQuestionsToMembers(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim pCopyAnswerHelper As CopyAnswerHelper
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer "EndCutRules.WebCutSel", QUES_ENDCUTTYPE, _
                                 "AssyConnRules.BeamByStiffenerSel", QUES_ENDCUTTYPE
    
    UpdateWebCutForFlangeCuts pMemberDescription
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMembers", strError).Number
End Sub

Public Sub CMCopyFCQuestionsToMembers(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim pCopyAnswerHelper As CopyAnswerHelper
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.CopyAnswer "EndCutRules.FlangeCutSel", QUES_ENDCUTTYPE, _
                                 "AssyConnRules.BeamByStiffenerSel", QUES_ENDCUTTYPE
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMembers", strError).Number
End Sub
Public Sub CMSetBottomFlange(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
          
    CMCopyFCQuestionsToMembers pMemberDescription
    
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.FlangeCutSel", "TheBottomFlange", "Yes"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetBottomFlange", strError).Number
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
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
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

