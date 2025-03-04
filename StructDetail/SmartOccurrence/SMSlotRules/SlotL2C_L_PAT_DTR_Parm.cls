VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotL2C_L_PAT_DTR_Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Parameter rule for Web Left Connected SlotL2C:
' This slot has: Point between Bottom and Left
'                Arc between Left and Top
'                Tangent Arc between Top and Right
'                Different Top Arc Radii (DTR)
'
' This slot TopClearance is independent from TopFlangeRightTopCorner radius
'
' Modify class name to particular rule
Private Const m_sClassName As String = "SlotL2C_L_PAT_DTR_Parm"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
'
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
  
  'Get the Penetrating Section
  Dim dProfilePartHeight As Double
  Dim sXSectionType As String
  If TypeOf oPLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
      'Penetrating Object is a Profile
      Dim oProfilePart As StructDetailObjects.ProfilePart
      Set oProfilePart = New StructDetailObjects.ProfilePart
      Set oProfilePart.object = oPLH.InputObject(INPUT_PENETRATING)
      sXSectionType = oProfilePart.sectionType
      dProfilePartHeight = oProfilePart.Height
      dProfilePartHeight = Round(dProfilePartHeight, 3)
      Set oProfilePart = Nothing
  ElseIf TypeOf oPLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
      'Penetrating Object is a Plate
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oSlot As New StructDetailObjects.Slot
      Set oSlot.object = oPLH.SmartOccurrence
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      
      oSlotMappingRule.GetSectionDepth oSlot.Penetrating, oSlot.Penetrated, dProfilePartHeight
      dProfilePartHeight = Round(dProfilePartHeight, 3)
  Else
      'Unknown Penetrating Object Type
      Exit Sub
  End If
   
   ' Set parameters
   Dim dSlotAngle As Double
   Dim dTopFlangeLeftTopCornerRadius As Double
   Dim dTopClearance As Double
   Dim dTopFlangeRightTopCornerRadius As Double
   
   dSlotAngle = GetSlotAngle(oPLH)
   
   If dProfilePartHeight > 0.1 And dProfilePartHeight <= 0.2 Then
      dTopFlangeLeftTopCornerRadius = 0.04
      dTopClearance = 0.025
      dTopFlangeRightTopCornerRadius = 0.04
   ElseIf dProfilePartHeight > 0.2 And dProfilePartHeight <= 0.3 Then
      dTopFlangeLeftTopCornerRadius = 0.05
      dTopClearance = 0.04
      dTopFlangeRightTopCornerRadius = 0.05
   ElseIf dProfilePartHeight > 0.3 And dProfilePartHeight <= 0.4 Then
      dTopFlangeLeftTopCornerRadius = 0.05
      dTopClearance = 0.05
      dTopFlangeRightTopCornerRadius = 0.065
   ElseIf dProfilePartHeight > 0.4 And dProfilePartHeight <= 0.6 Then
      dTopFlangeLeftTopCornerRadius = 0.065
      dTopClearance = 0.05
      dTopFlangeRightTopCornerRadius = 0.1
   ElseIf dProfilePartHeight > 0.6 And dProfilePartHeight <= 0.95 Then
      dTopFlangeLeftTopCornerRadius = 0.1
      dTopClearance = 0.05
      dTopFlangeRightTopCornerRadius = 0.1
   Else
 
      ' To be removed
      dTopFlangeLeftTopCornerRadius = 0.065
      dTopClearance = 0.05
      dTopFlangeRightTopCornerRadius = 0.1
   End If
          
   oPLH.Add PARAM_TOP_FLANGE_LEFT_TOP_CORNER_RADIUS, dTopFlangeLeftTopCornerRadius
   oPLH.Add PARAM_TOP_CLEARANCE, dTopClearance
   oPLH.Add PARAM_TOP_FLANGE_RIGHT_TOP_CORNER_RADIUS, dTopFlangeRightTopCornerRadius
   oPLH.Add PARAM_SLOT_ANGLE, dSlotAngle
  
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
   oIH.definition = oSelector
   oIH.InitAs m_sFamilyProgID
   ParameterRuleInputs oIH
   Set oIH = Nothing
  
   '----------- Define Output and Add it to the representation ------------------
   Dim oOH As IJDOutputsHelper
  
   Set oOH = New OutputHelper
   oOH.Representation = oSelector.IJDRepresentations.Item(1)
   oOH.SetOutput PARAM_TOP_FLANGE_LEFT_TOP_CORNER_RADIUS
   oOH.SetOutput PARAM_TOP_FLANGE_RIGHT_TOP_CORNER_RADIUS
   oOH.SetOutput PARAM_SLOT_ANGLE
   oOH.SetOutput PARAM_TOP_CLEARANCE
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

