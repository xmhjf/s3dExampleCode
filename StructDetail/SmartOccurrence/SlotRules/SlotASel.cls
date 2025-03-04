VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotASel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "SlotRules.SlotASel"
Const m_SelectorName As String = "SlotRules.SlotASel"
Const m_FamilyProgid As String = ""

Const INPUT_PENETRATED = "Penetrated"
Const INPUT_PENETRATING = "Penetrating"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\SlotASel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
  'Define questions
  pQH.SetQuestion "Clearance", 0.1, , "GenerateClearance", LIBRARY_SOURCE_ID
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  
  ' ************* Collect the Arguments ************
  
  'Get Penetrating object
  Dim oProfilePart As StructDetailObjects.ProfilePart
  Set oProfilePart = New StructDetailObjects.ProfilePart
  Set oProfilePart.object = pSLH.InputObject(INPUT_PENETRATING)

    ' Get the Question answers
   Dim clearance As Double
   clearance = pSLH.Answer("Clearance")

  ' ********************* Selection code ****************
  
  Select Case oProfilePart.SectionType
     Case "EA", "UA", "C_SS", "CSType"
        pSLH.Add "SlotAA"
     Case "B"
        pSLH.Add "SlotBA"
     Case "T_XType", "TSType", "BUT", "BUTL2", "I", "ISType"
        pSLH.Add "SlotTA"
     Case "F", "SB", "ST", "FB"
        pSLH.Add "SlotFA"
     Case "BUTL3"
        pSLH.Add "SlotLA"
     'Case Else
     '   sERROR = "Invalid cross section type ('" & oProfilePart.SectionType & "') specified for " & m_SelectorProgid
     '   GoTo ErrorHandler
  End Select
  
  ' *********************************************************
  
  Set oProfilePart = Nothing
  
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

