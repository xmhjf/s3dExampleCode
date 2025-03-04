VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BeamEndToBeamEndDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "AssyConnRules.BeamEndToBeamEndDef"
Const m_ItemName As String = "AssyConnRules.BeamEndToBeamEndDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\BeamEndToBeamEndDef.cls"

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
       
    '--------------------------
    'Web Cut for Part 1
    strError = "Constructing web cut."
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndWebCut1", 1, _
                                  "CMConstruct_WebCutEndToEndPart1", CMLIBRARY_ASSYCONNRULES)

    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToMembers"

    '--------------------------
    'Top flange cut for part 1
    strError = "Constructing flange cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndFlangeCut1", 2, _
                                  "CMConstruct_FlangeCutEndToEndPart1", CMLIBRARY_ASSYCONNRULES)

    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyFCQuestionsToMembers"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateTopFlangeCut"
    
    '--------------------------
    'Bottom flange cut for part 1
    strError = "Constructing flange cut 2."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndFlangeCut2", 3, _
                                  "CMConstruct_BottomFlangeCutEndToEndPart1", CMLIBRARY_ASSYCONNRULES)

    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetBottomFlange"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateBottomFlangeCut"
    
    '--------------------------
    'Web cut for part 2
    strError = "Constructing web cut for part 2."
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndWebCutPart2", 4, _
                                  "CMConstruct_WebCutEndToEndPart2", CMLIBRARY_ASSYCONNRULES)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetWeldPartNumberWeb"

    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    'pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToMembers"

    '--------------------------
    'Top flange cut for part 2
    strError = "Constructing flange cut for part 2."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndFlangeCut1Part1", 5, _
                                  "CMConstruct_FlangeCutEndToEndPart2", CMLIBRARY_ASSYCONNRULES)
    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetWeldPartNumberFlange"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateTopFlangeCut2"

    '--------------------------
    'Bottom flange cut for part 2
    strError = "Constructing flange cut 2 for part 2."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("BeamEndToBeamEndFlangeCut2Part2", 6, _
                                  "CMConstruct_BottomFlangeCutEndToEndPart2", CMLIBRARY_ASSYCONNRULES)

    'Copy Q/A in the assembly connection selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetWeldPartNumberFlange2"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateBottomFlangeCut2"
    
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
Public Sub CMSetWeldPartNumberWeb(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    strError = "Copying questions for end cut end to end weld"
    CMCopyQuestionsToMembers pMemberDescription
      
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set weld number"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.EndToEndWebCutSel", "WeldPart", "Second"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetWeldPartNumberWeb", strError).Number
End Sub

Public Sub CMSetWeldPartNumberFlange(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    strError = "Copying questions for end cut end to end weld"
    CMCopyFCQuestionsToMembers pMemberDescription
      
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set weld number"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.EndToEndFlangeCutSel", "WeldPart", "Second"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetWeldPartNumberFlange", strError).Number
End Sub
Public Sub CMSetWeldPartNumberFlange2(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
    
    strError = "Copying questions for end cut end to end weld"
    CMCopyFCQuestionsToMembers pMemberDescription
    
    CMSetBottomFlange2 pMemberDescription
      
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set weld number"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.EndToEndFlangeCutSel", "WeldPart", "Second"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetWeldPartNumberFlange2", strError).Number
End Sub


Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM

    strError = "Constructing physical connection."
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnFromACPorts(pMemberDescription, pResourceManager, "TeeWeld")
    
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
    pCopyAnswerHelper.CopyAnswer "EndCutRules.EndToEndWebCutSel", QUES_ENDCUTTYPE, _
                                 "AssyConnRules.BeamByBeamSel", QUES_ENDCUTTYPE

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
    pCopyAnswerHelper.CopyAnswer "EndCutRules.EndToEndFlangeCutSel", QUES_ENDCUTTYPE, _
                                 "AssyConnRules.BeamByBeamSel", QUES_ENDCUTTYPE

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyFCQuestionsToMembers", strError).Number
End Sub
Public Sub CMSetBottomFlange(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
          
    CMCopyFCQuestionsToMembers pMemberDescription
    
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.EndToEndFlangeCutSel", "TheBottomFlange", "Yes"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetBottomFlange", strError).Number
End Sub


Public Sub CMSetBottomFlange2(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
            
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.EndToEndFlangeCutSel", "TheBottomFlange", "Yes"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetBottomFlange2", strError).Number
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



