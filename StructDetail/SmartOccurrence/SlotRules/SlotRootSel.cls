VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotRootSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "SlotRules.SlotRootSel"
Const m_SelectorName As String = "SlotRules.SlotRootSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\SlotRootSel.cls"

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
  
'  'Define allowed answers to questions
'  Dim codeCol As New Collection
'  codeCol.Add gsDrop
'  codeCol.Add gsSlide
'  pQH.DefineCodeList "AssyMethodCol", codeCol
    
  'Define questions
    pQH.SetQuestion "AssyMethod", gsDrop, "AssyMethodCol"
    pQH.SetQuestion "ApplyTreatment", "No", "BooleanCol"

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
  On Error GoTo ErrorHandler
  Dim sERROR As String
  
  sERROR = "Picking Slot"
  
  'Get Penetrating object
  Dim oProfilePart As StructDetailObjects.ProfilePart
  Set oProfilePart = New StructDetailObjects.ProfilePart
  Set oProfilePart.object = pSLH.InputObject(INPUT_PENETRATING)
  
  'Get Slot object
  Dim oSlot As StructDetailObjects.Slot
  Set oSlot = New Slot
  Set oSlot.object = pSLH.SmartOccurrence

  ' Get the Question answers
  Dim sAssyMethod As String
  sAssyMethod = pSLH.Answer("AssyMethod")

  ' ********************* Selection code - TODO  ****************
  
  Select Case oProfilePart.SectionType
     Case "FB", "EA", "UA", "BUTL3"
        If oSlot.IsPerpendicular = True And sAssyMethod = gsSlide Then
            pSLH.Add "SlotI"
            pSLH.Add "SlotC"
            pSLH.Add "SlotA"
        ElseIf oSlot.IsPerpendicular = True And sAssyMethod = gsDrop Then
            pSLH.Add "SlotC"
            pSLH.Add "SlotA"
        ElseIf oSlot.IsPerpendicular = False And sAssyMethod = gsSlide Then
            pSLH.Add "SlotC"
            pSLH.Add "SlotA"
        ElseIf oSlot.IsPerpendicular = False And sAssyMethod = gsDrop Then
            pSLH.Add "SlotA"
        End If
     Case "T_XType", "TSType", "BUT", "BUTL2"
        If oSlot.IsPerpendicular = True And sAssyMethod = gsSlide Then
            pSLH.Add "SlotI"
            pSLH.Add "SlotA"
        Else
            pSLH.Add "SlotA"
        End If
     Case "B"
        If oSlot.IsPerpendicular = False And sAssyMethod = gsDrop Then
            pSLH.Add "SlotA"
        Else
            pSLH.Add "SlotC"
            pSLH.Add "SlotA"
        End If
     Case "C_SS", "CSType", "I", "ISType"
        pSLH.Add "SlotA"
     Case "HalfR", "P", "R"
        sERROR = "No Slot defined for this profile type ('" & oProfilePart.SectionType & "'"
        GoTo ErrorHandler
     Case Else
        sERROR = "Invalid cross section type ('" & oProfilePart.SectionType & "') specified for " & m_SelectorProgid
        GoTo ErrorHandler
  End Select
  
  ' *********************************************************
  
  Set oProfilePart = Nothing
  Set oSlot = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic", sERROR).Number
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

