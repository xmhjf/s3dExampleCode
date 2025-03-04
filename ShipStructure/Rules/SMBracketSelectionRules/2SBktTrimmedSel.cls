VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt2STrimmedSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgid As String = "BktSelRules.Bkt2STrimmedSel"
Const m_SelectorName As String = "BktSelRules.Bkt2STrimmedSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktSelRules\2SBktTrimmedSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    
    pIH.SetInput INPUT_SUPPORTS

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
   
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketContourType As String
    sBracketContourType = pSL.SelectorAnswer("BktSelRules.Bkt2SSel", QUESTION_BracketContourType)
            
        Select Case sBracketContourType
            Case "Linear"
                pSL.Add "BPS_2SL HxWxN1xN2"
                pSL.Add "BPS_2SL 400x400x50x50"
                pSL.Add "BPS_2SL 600x600x75x75"
                pSL.Add "BPS_2SL 800x800x75x75"
                pSL.Add "BPS_2SL 1000x1000x100x100"
                pSL.Add "BPS_2SL 2000x2000x150x150"
                 pSL.Add "BPS_2SLT HxWxN1xN2"

            Case "Curved"
                pSL.Add "BPS_2SC HxWxN1xN2"
                pSL.Add "BPS_2SC 400x400x50x50"
                pSL.Add "BPS_2SC 600x600x75x75"
                pSL.Add "BPS_2SC 800x800x75x75"
                pSL.Add "BPS_2SC 1000x1000x100x100"
                pSL.Add "BPS_2SC 2000x2000x150x150"
                pSL.Add "BPS_2SCT HxWxN1xN2"
        End Select
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub
'Public Sub BracketType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
'    ByRef ppArgument As Object)
'    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'    ' Get Symbol Rep so Selector Logic object can be created
'    ' Get Symbol Definition
'    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
'    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
'    Set oInputDG = pInput
'    Set oSymbolDefinition = oInputDG.definition
'
'    ' Create/Initialize the Selector Logic Object from the symbol definition
'    Dim pSL As IJDSelectorLogic
'    Set pSL = New SelectorLogic
'    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
'
'    pSL.Answer(QUESTION_BracketType) = "2SLinearTrim"
'    Exit Sub
    
'End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **



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
    pIH.definition = pSelector
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
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



