VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt3SLapSup1Sel"
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

Const m_SelectorProgid As String = "BktUnifiedRules.Bkt3SLapSup1Sel"
Const m_SelectorName As String = "BktUnifiedRules.Bkt3SLapSup1Sel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BracketRules\Bkt3SLapSup1Sel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
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

    'Get answer from parent...
    sBracketContourType = pSL.SelectorAnswer("BktUnifiedRules.Bkt3SSel", QUESTION_BracketContourType)
   
    Select Case sBracketContourType
       Case "Linear"
           ' Additional Method could be added to check the 3rd support type
           ' and conditions to filter selection set...
           pSL.Add "BKT_3SLLS1_SB"
           pSL.Add "BKT_3SLLS1"
           pSL.Add "BKT_3SLLS1_OF"
           pSL.Add "BKT_3SLLS1_IF"
           pSL.Add "BKT_3SLLS1_FT"
           pSL.Add "BKT_3SLLS1_OT"
           pSL.Add "BKT_3SLCLS1_FB"
           pSL.Add "BKT_3SLLS1_SB_Far"
           pSL.Add "BKT_3SLCLS1_SB_Far"
       Case "Curved"
           pSL.Add "BKT_3SCLS1_SB"
           pSL.Add "BKT_3SCLS1"
           pSL.Add "BKT_3SCLS1_OT"
           pSL.Add "BKT_3SCLS1_Sym"
            
     End Select
     
Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub


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
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************










