VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToEndWebCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgId
Const m_SelectorProgid As String = CUSTOMERID + "EndCutRules.EndToEndWebCutSel"   ' - TODO
Const m_SelectorName As String = CUSTOMERID + "EndCutRules.EndToEndWebCutSel"                   ' - TODO
Const m_FamilyProgid As String = ""       ' TODO

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToEndWebCutSel.cls"

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
zMsgBox "EndToEndWebCutSel::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions.  !! Note !!  This default is only used if web cut is NOT created by an assembly connection.
    strError = "Setting questions."
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList"
    
    pQH.SetQuestion gsWeldPart, "First"
    
    ''' Can NOT Delete (Remove) a Question from the Selector
    ''' (Can Not change the order of Questions in a Selector)
    ''' to remove a Question, set Name and Default value to NULL
    ''' Question will not be displayed on Property Pages
    ''' the Question Name can not be re-used
    pQH.SetQuestion "", ""
''''    pQH.SetQuestion "EndToEndCase", "NoAngle"
    
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
    
    Dim oPort As IJPort
    Dim oBounded  As Object
    Dim strError As String
    
    Set oPort = pSLH.InputObject(INPUT_BOUNDED)
    If (oPort Is Nothing) Then
        strError = "pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    Else
        Set oBounded = oPort.Connectable
    End If
    Set oPort = Nothing
    
    Dim strEndCutType As String
    Dim strWeldPartNumber As String
    strEndCutType = pSLH.Answer(QUES_ENDCUTTYPE)
    strWeldPartNumber = pSLH.Answer(gsWeldPart)
                        
    Dim sSplitEndToEndCase As String
    sSplitEndToEndCase = pSLH.Answer("SplitEndToEndCase")
    
    Dim sFlipEndToEndAngle As String
    sFlipEndToEndAngle = pSLH.Answer("FlipEndToEndAngle")
    
zMsgBox "EndToEndWebCutSel::SelectorLogic" & vbCrLf & _
        "strEndCutType     : " & strEndCutType & vbCrLf & _
        "strWeldPartNumber : " & strWeldPartNumber & vbCrLf & _
        "sSplitEndToEndCase: " & sSplitEndToEndCase & vbCrLf & _
        "sFlipEndToEndAngle: " & sFlipEndToEndAngle
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    Dim oBoundedPart As Object
    Dim oBoundedPort As IJPort
    
    Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = "pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
    
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedPart = oBoundedPort.Connectable
    Else
        ' error, BoundedObject MUST be a IJPort object
        strError = "BoundedObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
        
    If LCase(Trim(sSplitEndToEndCase)) = LCase("NoAngle") Then
        Call WebCutNoFlangeEndToEnd(oBounded, False, strEndCutType, strWeldPartNumber, pSLH)
    ElseIf LCase(Trim(sSplitEndToEndCase)) = LCase("AlongGlobalAxis") Then
        Call SelectAlongAxisEndToEndWebCut(strEndCutType, strWeldPartNumber, pSLH)
    Else
        Call WebCutEndToEndAngle(oBounded, False, strEndCutType, _
                                 strWeldPartNumber, _
                                 sSplitEndToEndCase, sFlipEndToEndAngle, pSLH)
    End If
    
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


