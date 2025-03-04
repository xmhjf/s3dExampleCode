VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotC2Sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotC2Sel"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

Private Sub SelectorInputs(oIH As IJDInputsHelper)
   On Error GoTo ErrorHandler
    
   oIH.SetInput INPUT_PENETRATING
   oIH.SetInput INPUT_PENETRATED

   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Private Sub SelectorQuestions(oQH As IJDQuestionsHelper)
   On Error GoTo ErrorHandler
 
   'Define questions
   oQH.SetQuestion "Clearance", 0.1, , "GenerateClearance", LIBRARY_SOURCE_ID
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Private Sub SelectorLogic(oSLH As IJDSelectorLogic)
   On Error GoTo ErrorHandler
   Dim sERROR As String
   
  'Get the Penetrating Section
    Dim sXSectionType As String
    If TypeOf oSLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
        'Penetrating Object is a Profile
        Dim oProfilePart As StructDetailObjects.ProfilePart
        Set oProfilePart = New StructDetailObjects.ProfilePart
        Set oProfilePart.object = oSLH.InputObject(INPUT_PENETRATING)
        sXSectionType = oProfilePart.sectionType
        
        Set oProfilePart = Nothing
    ElseIf TypeOf oSLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
        'Penetrating Object is a Plate
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oSlot As New StructDetailObjects.Slot
        Set oSlot.object = oSLH.SmartOccurrence
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
        'Unknown Penetrating Object Type
        Exit Sub
    End If

    ' Get the Question answers
   Dim clearance As Double
   
   clearance = oSLH.Answer("Clearance")
   
   ' ********************* Selection code ****************
   ' Top Connected
   Select Case sXSectionType
      Case "BUTL2"
         oSLH.Add "SlotL2C2_T_PTT_STR"
         oSLH.Add "SlotL2C2_T_PTT_STR2"
      Case "BUT"
         oSLH.Add "SlotTC2_T_PTT_STR"
         oSLH.Add "SlotTC2_T_PTT_STR2"
         oSLH.Add "SlotTC2_T_PAA_STR"
      Case Else
         'sERROR = "Invalid cross section type ('" & sProfileSectionType & "') specified for " & m_sRuleProgID
         sERROR = "Invalid cross section type ('" & sXSectionType & "') specified for " & m_sRuleProgID
         GoTo ErrorHandler
   End Select
 
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
  IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
  
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oSelector.IJDInputs.RemoveAllInput
   oSelector.IJDRepresentations.RemoveAllRepresentation
  
   Dim oSelectorFact As New DefinitionFactory
   
   oSelectorFact.InitAbstractSelector oSelector
   Dim oIH As IJDInputsHelper
   
   Set oIH = New InputHelper
   oIH.definition = oSelector
   oIH.InitAs m_sFamilyProgID
   SelectorInputs oIH
   Set oIH = Nothing
   
   Dim oQH As IJDQuestionsHelper
   
   Set oQH = New QuestionHelper
   oQH.Selector = oSelector
   SelectorQuestions oQH
   Set oQH = Nothing
   
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal sCB As String, ByVal vDP As Variant, ByVal oRM As Object) As Object
   Dim oSelectorFact As New DefinitionFactory
   
   Set IJDUserSymbolServices_InstanciateDefinition = oSelectorFact.InstanciateSelector( _
                                          m_sRuleProgID, _
                                          sCB, _
                                          IJDUserSymbolServices_GetDefinitionName(vDP), _
                                          oRM)
   Set oSelectorFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, _
                                                       ByVal sRepName As String, _
                                                       ByVal oOutputColl As Object, _
                                                       vArrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, _
                                                     ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(oRep As IJDRepresentation)
   Dim oSLH As IJDSelectorLogic
   
   Set oSLH = New SelectorLogic
   oSLH.Representation = oRep
   SelectorLogic oSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
