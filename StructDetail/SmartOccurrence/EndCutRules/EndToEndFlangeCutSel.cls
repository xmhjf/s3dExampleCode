VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToEndFlangeCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID
Const m_SelectorProgid As String = "EndCutRules.EndToEndFlangeCutSel"   ' - TODO
Const m_SelectorName As String = "EndCutRules.EndToEndFlangeCutSel"                   ' - TODO
Const m_FamilyProgid As String = ""       ' TODO

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\EndToEndFlangeCutSel.cls"


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
zMsgBox "EndToEndFlangeCutSel::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions.  !! Note !!  This default is only used if flange cut is NOT created by an assembly connection.
    strError = "Setting questions."
    SetQuestionEndCutType pQH
    
    pQH.SetQuestion gsWeldPart, "First"
    pQH.SetQuestion gsBottomFlange, "No", "BooleanCol"
    
    ''' Can NOT Delete (Remove) a Question from the Selector
    ''' (Can Not change the order of Questions in a Selector)
    ''' to remove a Question, set Name and Default value to NULL
    ''' Question will not be displayed on Property Pages
    ''' the Question Name can not be re-used
    pQH.SetQuestion "", ""
'''"    pQH.SetQuestion "EndToEndCase", "NoAngle"
    
    ' Add new Questions after all existing Questions
    pQH.SetQuestion "SplitEndToEndCase", "NoAngle", "SplitEndCutTypes"
    pQH.SetQuestion "FlipEndToEndAngle", "NoFlip", "FlipSplitEndCutAngle"
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim oBounded  As Object
    Dim oPort As IJPort
    Dim strError As String
    
    Set oPort = pSLH.InputObject(INPUT_BOUNDED)
    If (oPort Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    Else
        Set oBounded = oPort.Connectable
    End If
    Set oPort = Nothing
    
    Dim strEndCutType As String
    Dim strWeldPartNumber As String
    strEndCutType = pSLH.Answer(QUES_ENDCUTTYPE)
    strWeldPartNumber = pSLH.Answer(gsWeldPart)
            
    'get the bottom flange answer and see if this is top or bottom
    Dim sBottomFlange As String
    sBottomFlange = pSLH.Answer(gsBottomFlange)
            
    Dim sSplitEndToEndCase As String
    sSplitEndToEndCase = pSLH.Answer("SplitEndToEndCase")
            
zMsgBox "EndToEndFlangeCutSel::SelectorLogic" & vbCrLf & _
        "strEndCutType     : " & strEndCutType & vbCrLf & _
        "strWeldPartNumber : " & strWeldPartNumber & vbCrLf & _
        "sBottomFlange     : " & sBottomFlange & vbCrLf & _
        "sSplitEndToEndCase: " & sSplitEndToEndCase
            
    Call FlangeCutNoFlangeEndToEnd(oBounded, True, strEndCutType, strWeldPartNumber, pSLH)
   
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
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


