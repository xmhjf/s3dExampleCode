VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt3SSel"
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

Const m_SelectorProgid As String = "BktUnifiedRules.Bkt3SSel"
Const m_SelectorName As String = "BktUnifiedRules.Bkt3SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BracketRules\3SBktSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' inputs:
    '   Plane that bracket is on
    '   Bracket Plate System
    '   Group of supports

    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
    pIH.SetInput INPUT_SUPPORTS
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler

    pQH.SetQuestion QUESTION_BracketContourType, _
                    DEFAULT_BracketContourType_LINEAR, _
                    CODELIST_BracketContourType

    pQH.SetQuestion QUESTION_BracketSupport1Type, _
                    DEFAULT_BracketSupport1or2Type_TRIMMED, _
                    CODELIST_BracketSupport1or2Type
    
    pQH.SetQuestion QUESTION_BracketSupport2Type, _
                    DEFAULT_BracketSupport1or2Type_TRIMMED, _
                    CODELIST_BracketSupport1or2Type
    
    pQH.SetQuestion QUESTION_BracketSupport3Type, _
                    DEFAULT_BracketSupport3or4or5Type_TRIMMED, _
                    CODELIST_BracketSupport3or4or5Type
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sBracketSupport1Type As String
    Dim sBracketSupport2Type As String
    Dim sBracketSupport3Type As String
    sBracketSupport1Type = pSL.Answer(QUESTION_BracketSupport1Type)
    sBracketSupport2Type = pSL.Answer(QUESTION_BracketSupport2Type)
    sBracketSupport3Type = pSL.Answer(QUESTION_BracketSupport3Type)
    Dim bLapSup1 As Boolean
    Dim bLapSup2 As Boolean
    Dim bLapSup3 As Boolean
    
    If (sBracketSupport1Type = "Lapped") Then
        bLapSup1 = True
    Else
        bLapSup1 = False
    End If
    
    If (sBracketSupport2Type = "Lapped") Then
        bLapSup2 = True
    Else
        bLapSup2 = False
    End If
    
    If (sBracketSupport3Type = "Lapped") Then
        bLapSup3 = True
    Else
        bLapSup3 = False
    End If
            
        
    If bLapSup1 And bLapSup2 And bLapSup3 Then
             pSL.Add "BKT_3S_LapAllSupports"
    ElseIf bLapSup1 And bLapSup2 Then
             pSL.Add "BKT_3S_LapSup1Sup2"
    ElseIf bLapSup1 And bLapSup3 Then
             pSL.Add "BKT_3S_LapSup1Sup3"
    ElseIf bLapSup1 Then
             pSL.Add "BKT_3S_LapSup1"
    ElseIf bLapSup2 And bLapSup3 Then
             pSL.Add "BKT_3S_LapSup2Sup3"
    ElseIf bLapSup2 Then
             pSL.Add "BKT_3S_LapSup2"
    ElseIf bLapSup3 Then
             pSL.Add "BKT_3S_LapSup3"
    Else
             pSL.Add "BKT_3S_Trimmed"
    End If

    pSL.Add "BKT_3S_Intergrated"

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



