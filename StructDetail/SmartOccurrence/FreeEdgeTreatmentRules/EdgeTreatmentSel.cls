VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdgeTreatmentSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "FETRules.EdgeTreatmentSel"
Const m_SelectorName As String = "FETRules.EdgeTreatmentSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\FreeEdgeTreatmentRules\EdgeTreatmentSel.cls"

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

    pQH.SetQuestion "EdgeTreatment", "Grind", "EdgeTreatmentType", "EdgeTreatmentCM", "FETRules.EdgeTreatmentSelCM"
 
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  
    ' Get the Question answers
    Dim sETType As String
    sETType = pSLH.Answer("EdgeTreatment")
  
    If sETType = "Grind" Then
        pSLH.Add "Grind"
    Else
        pSLH.Add "Bevel"
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


