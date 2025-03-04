VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotAIParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "SlotRules.SlotAIParm"
Const m_ParameterRuleName As String = CUSTOMERID + "SlotRules.SlotAIParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "SlotRules\SlotAIParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
  'Get the Penetrating Section
  Dim dProfilePartHeight As Double
  Dim sXSectionType As String
  If TypeOf pPLH.InputObject(INPUT_PENETRATING) Is IJProfile Then
      'Penetrating Object is a Profile
      Dim oProfilePart As StructDetailObjects.ProfilePart
      Set oProfilePart = New StructDetailObjects.ProfilePart
      Set oProfilePart.object = pPLH.InputObject(INPUT_PENETRATING)
      sXSectionType = oProfilePart.sectionType
      dProfilePartHeight = oProfilePart.Height
      dProfilePartHeight = Round(dProfilePartHeight, 3)
      Set oProfilePart = Nothing
  ElseIf TypeOf pPLH.InputObject(INPUT_PENETRATING) Is IJPlate Then
      'Penetrating Object is a Plate
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oSlot As New StructDetailObjects.Slot
      Set oSlot.object = pPLH.SmartOccurrence
      
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
  
  ' Parameter rule:  Set slot snipe size based on height of profile
  If dProfilePartHeight <= 0.07 Then
     pPLH.Add "SnipeSize", 0.001
  ElseIf dProfilePartHeight > 0.07 And dProfilePartHeight <= 0.12 Then
     pPLH.Add "SnipeSize", 0.01
  Else ' oProfilePart.Height > 120mm, so...
     pPLH.Add "SnipeSize", 0.015
  End If
  
  Set oProfilePart = Nothing
  
  SetSlotAssyOrientation pPLH
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub


' This method sets the values of the IJUASlotAssyOrientation custom
' atrributes that have been placed on the slot.
' Version 1
' Note: The IJAssemblyOrientation interface is in Level 4 (Planning)
' while this rule is in level 3.  Therefore the reference to the
' interface cannot be made in code that will be delivered with
' IntelliShip.  This code is provided as an example only and requires
' a reference to 'Ingr GSCAD Planning Assembly 1.0 Type Library'.
Private Sub SetSlotAssyOrientation(pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler

  ' Get one of the inputs on the slot, either the penetrated plate or the
  ' penetrating profile that can be used to obtain the parent Assembly.
  Dim oAssemblyChild As IJAssemblyChild
  Set oAssemblyChild = pPLH.InputObject(INPUT_PENETRATING)

  ' Get the Assembly in which the slot exists.
  Dim oAssembly As IJAssembly
  Set oAssembly = oAssemblyChild.Parent
  
  If oAssembly Is Nothing Then
     ' The part has not been assigned to an assembly yet(so is the slot)
     GoTo Cleanup
  End If
  
  If Not TypeOf oAssembly Is IJLocalCoordinateSystem Then
      'This assembly does not support IJLocalCoordinateSystem,
      'no orientation associated with it,so it won't affect slot
      GoTo Cleanup
  End If
  
  ' Get the orientation of the Assembly.
  Dim oAssemblyOrientation As IJLocalCoordinateSystem
  Set oAssemblyOrientation = oAssembly
  
  Dim vector As IJDVector
  Set vector = oAssemblyOrientation.ZAxis
  
  ' Set the values on the slot attributes.
  Dim oSlot As IJSmartOccurrence
  Set oSlot = pPLH.SmartOccurrence
  
  Dim oAttrHelper As IJDAttributes
  Set oAttrHelper = oSlot

  Dim oAttributeCol As IMSAttributes.IJDAttributesCol
  Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUASlotAssyOrientation")
  
  Dim oAttr As IJDAttribute
  Set oAttr = oAttributeCol.Item("AssyOrientationX")
  oAttr.Value = vector.x
  Set oAttr = oAttributeCol.Item("AssyOrientationY")
  oAttr.Value = vector.y
  Set oAttr = oAttributeCol.Item("AssyOrientationZ")
  oAttr.Value = vector.z
    
  GoTo Cleanup
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
  
Cleanup:
  Set oAssemblyChild = Nothing
  Set oAssembly = Nothing
  Set oAssemblyOrientation = Nothing
  Set vector = Nothing
  Set oSlot = Nothing
  Set oAttrHelper = Nothing
  Set oAttributeCol = Nothing
  Set oAttr = Nothing
  
End Sub


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pSelector.IJDInputs.RemoveAllInput
  pSelector.IJDRepresentations.RemoveAllRepresentation
 
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractParameterRule pSelector
 
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)
  pOH.SetOutput "SnipeSize"
  
  Set pOH = Nothing

End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)

  Dim pPLH As IJDParameterLogic
  Set pPLH = New ParameterLogic
  pPLH.Representation = pRep
  ParameterRuleLogic pPLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

