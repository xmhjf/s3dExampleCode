VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToMiterFlangeCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID
Const m_SelectorProgid As String = "EndCutRules.EndToMiterFlangeCutSel"
Const m_SelectorName As String = "EndCutRules.EndToMiterFlangeCutSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\EndToMiterFlangeCutSel.cls"
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    'Define questions.
    sError = "Setting questions."
    SetQuestionEndCutType pQH
    pQH.SetQuestion gsWeldPart, "First"
    pQH.SetQuestion "KnuckleEndCut", "Miter", "ProfileKnuckleEndCutTypes"
    pQH.SetQuestion gsBottomFlange, "No", "BooleanCol"
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", sError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    Dim sEndCutType As String
    Dim sSectionType As String
    Dim sBottomFlange As String
    Dim sKnuckleEndCut As String
    Dim sWeldPartNumber As String
    Dim sApplyTreatment As String
    
    Dim bRightFlangeOnly As Boolean
    
    Dim oBoundedPort As IJPort
    Dim oBoundedPart As Object
    Dim oBoundedInput As Object
    
    'verify Bounded Object implements IJPort interface
    Set oBoundedInput = pSLH.InputObject(INPUT_BOUNDED)
    If oBoundedInput Is Nothing Then
        sError = "pSLH.InputObject(INPUT_BOUNDED) is not valid object (Nothing)"
        GoTo ErrorHandler
    ElseIf Not TypeOf oBoundedInput Is IJPort Then
        sError = "pSLH.InputObject(INPUT_BOUNDED) is not valid IJPort object"
        GoTo ErrorHandler
    End If
    
    Set oBoundedPort = oBoundedInput
    Set oBoundedPart = oBoundedPort.Connectable
    sSectionType = GetCrossSectionType(oBoundedPart)
    If (sSectionType = "EA") Or _
       (sSectionType = "UA") Or _
       (sSectionType = "B") Then
       bRightFlangeOnly = True
    Else
       bRightFlangeOnly = False
    End If
    
    sEndCutType = pSLH.Answer(QUES_ENDCUTTYPE)
    sWeldPartNumber = pSLH.Answer(gsWeldPart)
    sKnuckleEndCut = pSLH.Answer("KnuckleEndCut")
    sBottomFlange = pSLH.Answer(gsBottomFlange)
    sApplyTreatment = pSLH.Answer(gsApplyTreatment)

    Select Case sEndCutType
        Case gsW
            If LCase(Trim(sWeldPartNumber)) = LCase("First") Then
                If bRightFlangeOnly Then
                   pSLH.Add "FlangeCut_W1MiterPC"
                Else
                    pSLH.Add "FlangeCut_W4MiterPC"
                End If
            Else
                If bRightFlangeOnly Then
                    pSLH.Add "FlangeCut_W1MiterEnd"
                Else
                    pSLH.Add "FlangeCut_W4MiterEnd"
                End If
            End If
        
        Case Else
            ' "W" is only valid type for Miter Cut:: No Flange Cut applied
        
    End Select
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", sError).Number
End Sub
  
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
   
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    pSLH.Representation = pRep
    SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




