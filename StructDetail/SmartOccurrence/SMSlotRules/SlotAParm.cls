VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotAParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Parameter rule for open slots
' Removed height restrictions which restricts selection of smart item
Private Const m_sClassName As String = "SlotAParm" ' Modify this to particular rule

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\" + m_sClassName + ".cls"

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
  
   Dim bSetValue As Boolean
   Dim dSlotAngle As Double
   Dim dSlotAngleLeft As Double
   Dim dCornerRadius As Double
   Dim dTopClearance As Double
   Dim dFlangeClearance As Double
   Dim dBackClearance As Double
    
   bSetValue = False
   dSlotAngle = GetSlotAngle(oPLH)
   dSlotAngleLeft = dSlotAngle
  
   ' Set corner radius,clearances
   Select Case sXSectionType
      Case "EA", "UA", "FB", "BUT", "BUTL2"
      
         If dProfilePartHeight <= 0.3 Then
            dCornerRadius = 0.05
            dTopClearance = 0.05
            dFlangeClearance = 0.05
            dBackClearance = 0.05
            bSetValue = True
            
         Else
            dCornerRadius = 0.075
            dTopClearance = 0.075
            dFlangeClearance = 0.075
            dBackClearance = 0.075
            bSetValue = True
         End If
         
      Case "B"
         ' Tsuneishi specs is not quite clear...
         If dProfilePartHeight <= 0.3 Then
            dCornerRadius = 0.05
            dTopClearance = 0.05
            dFlangeClearance = 0.06
            dBackClearance = 0.05
            bSetValue = True
            
         Else
            dCornerRadius = 0.075
            dTopClearance = 0.075
            dFlangeClearance = 0.085
            dBackClearance = 0.075
            bSetValue = True
         End If
         
      Case Else
         MsgBox m_sClassName & " should not be fired for " & sXSectionType & ", check selection rule."
   End Select
  
   If bSetValue = True Then
      oPLH.Add PARAM_CORNER_RADIUS, dCornerRadius
      oPLH.Add PARAM_TOP_CLEARANCE, dTopClearance
      oPLH.Add PARAM_FLANGE_CLEARANCE, dFlangeClearance
      oPLH.Add PARAM_BACK_CLEARANCE, dBackClearance
      oPLH.Add PARAM_SLOT_ANGLE, dSlotAngle
      oPLH.Add PARAM_SLOT_ANGLE_LEFT, dSlotAngleLeft
  
      SetSlotAssyOrientation oPLH
   End If
   
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
  IJDUserSymbolServices_GetDefinitionName = m_sRuleProgID
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSelector As IJDSymbolDefinition)
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oSelector.IJDInputs.RemoveAllInput
   oSelector.IJDRepresentations.RemoveAllRepresentation
 
   Dim oSelectorFact As New DefinitionFactory
  
   oSelectorFact.InitAbstractParameterRule oSelector
   Set oSelectorFact = Nothing
  
   Dim oIH As IJDInputsHelper
  
   Set oIH = New InputHelper
   oIH.definition = oSelector
   oIH.InitAs m_sFamilyProgID
   ParameterRuleInputs oIH
   Set oIH = Nothing
  
   '----------- Define Output and Add it to the representation ------------------
   Dim oOH As IJDOutputsHelper
  
   Set oOH = New OutputHelper
   oOH.Representation = oSelector.IJDRepresentations.Item(1)
   oOH.SetOutput PARAM_CORNER_RADIUS
   oOH.SetOutput PARAM_TOP_CLEARANCE
   oOH.SetOutput PARAM_FLANGE_CLEARANCE
   oOH.SetOutput PARAM_BACK_CLEARANCE
   oOH.SetOutput PARAM_SLOT_ANGLE
   oOH.SetOutput PARAM_SLOT_ANGLE_LEFT
   Set oOH = Nothing

End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal sCB As String, ByVal vDP As Variant, ByVal oRM As Object) As Object
   Dim oSelectorFact As New DefinitionFactory
  
   Set IJDUserSymbolServices_InstanciateDefinition = oSelectorFact.InstanciateParameterRule(m_sRuleProgID, sCB, IJDUserSymbolServices_GetDefinitionName(vDP), oRM)
   Set oSelectorFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal oRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
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
