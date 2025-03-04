VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GrindSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "FETRules.GrindSel"
Const m_SelectorName As String = CUSTOMERID + "FETRules.GrindSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "FreeEdgeTreatmentRules\GrindSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
'    'Add port inputs
    pIH.SetInput INPUT_TREATMENT_EDGE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
    ''' Can NOT Delete (Remove) a Question from the Selector
    ''' (Can Not change the order of Questions in a Selector)
    ''' to remove a Question, set Name and Default value to NULL
    ''' Question will not be displayed on Property Pages
    ''' the Question Name can not be re-used
    
''''    pQH.SetQuestion "GrindMethod", "Flat", "GrindMethod", "FETRules.EdgeTreatmentSelCM"
''''    pQH.SetQuestion "Flip", "No",  "BooleanCol"
    
    pQH.SetQuestion "", ""
    pQH.SetQuestion "", ""
    
    ' Add new Questions after all existing Questions
    pQH.SetQuestion "Grind Shape", "Flat", "GrindMethod", "GrindMethodCM", CUSTOMERID + "FETRules.EdgeTreatmentSelCM"
    pQH.SetQuestion "Category", "TwoSided", "ButtWeldCategory"
    pQH.SetQuestion "Grind Side", "Molded", "FirstWeldingSideCol"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  
    ' Get the Question answers
    Dim sGrind As String
    sGrind = pSLH.Answer("Grind Shape")
    Dim sCategory As String
    sCategory = pSLH.Answer("Category")
    Dim sGrindSide As String
    sGrindSide = pSLH.Answer("Grind Side")
            
    If sCategory = "OneSided" Then
        If sGrind = "Flat" Then
            If sGrindSide = "Molded" Then
                pSLH.Add "GrindOnMolded"
            Else
                pSLH.Add "GrindOnAntiMolded"
            End If
        Else
            If sGrindSide = "Molded" Then
                pSLH.Add "RoundGrindOnMolded"
            Else
                pSLH.Add "RoundGrindOnAntiMolded"
            End If
        End If
    Else
        If sGrind = "Flat" Then
            If sGrindSide = "Molded" Then
                pSLH.Add "Flat1"
            Else
                pSLH.Add "Flat2"
            End If
        Else
            If sGrindSide = "Molded" Then
                pSLH.Add "Round1"
            Else
                pSLH.Add "Round2"
            End If
        End If
    End If
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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
  
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
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
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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


