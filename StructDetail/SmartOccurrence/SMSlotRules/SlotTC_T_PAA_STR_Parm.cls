VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotTC_T_PAA_STR_Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Parameter rule for Top Connected SlotTC:
' This slot has: Point between Bottom and Left
'                Arc between Left and Top
'                Arc between Top and Right
'                Same Top arc Radii (STR)

' It applies to following cross section type:
' T

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotTC_T_PAA_STR_Parm"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

Private Sub ParameterRuleInputs(oIH As IJDInputsHelper)
   On Error GoTo ErrorHandler
    
   oIH.SetInput INPUT_PENETRATING
   oIH.SetInput INPUT_PENETRATED
    
   Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Private Sub ParameterRuleLogic(oPLH As IJDParameterLogic)
   On Error GoTo ErrorHandler

   ' Set parameters
   Dim dSlotAngle As Double
   Dim dCornerRadius As Double
   Dim dProfilePartWidth As Double
   dSlotAngle = GetSlotAngle(oPLH)
   If TypeOf oPLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
        dCornerRadius = 0.03
   ElseIf TypeOf oPLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oSlot As New StructDetailObjects.Slot
      Set oSlot.object = oPLH.SmartOccurrence
      oSlotMappingRule.GetSectionWidth oSlot.Penetrating, oSlot.Penetrated, dProfilePartWidth
      dProfilePartWidth = Round(dProfilePartWidth, 3)
      If dProfilePartWidth >= 0.05 And dProfilePartWidth < 0.065 Then
        dCornerRadius = 0.02
      Else
        dCornerRadius = 0.03
      End If
   Else
       'unknown section
   End If
   oPLH.Add PARAM_SLOT_ANGLE, dSlotAngle
   oPLH.Add PARAM_CORNER_RADIUS, dCornerRadius
  
   SetSlotAssyOrientation oPLH
   Set oPLH = Nothing
  
   Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal vDefinitionParameters As Variant) As String
   IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oSelector.IJDInputs.RemoveAllInput
   oSelector.IJDRepresentations.RemoveAllRepresentation
 
   Dim oSelectorFact As New DefinitionFactory
   Dim oIH As IJDInputsHelper
  
   oSelectorFact.InitAbstractParameterRule oSelector
 
   Set oIH = New InputHelper
   oIH.Definition = oSelector
   oIH.InitAs m_sFamilyProgID
   ParameterRuleInputs oIH
   Set oIH = Nothing
  
   '----------- Define Output and Add it to the representation ------------------
   Dim oOH As IJDOutputsHelper
  
   Set oOH = New OutputHelper
   oOH.Representation = oSelector.IJDRepresentations.Item(1)
   oOH.SetOutput PARAM_SLOT_ANGLE
   oOH.SetOutput PARAM_CORNER_RADIUS
   Set oOH = Nothing
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal sCB As String, ByVal vDP As Variant, ByVal oRM As Object) As Object
   Dim oSelectorFact As New DefinitionFactory
  
   Set IJDUserSymbolServices_InstanciateDefinition = oSelectorFact.InstanciateParameterRule(m_sRuleProgID, sCB, IJDUserSymbolServices_GetDefinitionName(vDP), oRM)
   Set oSelectorFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal sRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(oRep As IJDRepresentation)
   Dim oPLH As IJDParameterLogic
  
   Set oPLH = New ParameterLogic
   oPLH.Representation = oRep
   ParameterRuleLogic oPLH
   Set oPLH = Nothing

End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
