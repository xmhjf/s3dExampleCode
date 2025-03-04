VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarClipParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "CollarRules.CollarClipParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CollarRules.CollarClipParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CollarRules\CollarClipParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_BOUNDINGPLATE
  pIH.SetInput INPUT_SLOT
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
  'Get slot object
  Dim oSlotClass As StructDetailObjects.Slot
  Set oSlotClass = New Slot
  Set oSlotClass.object = pPLH.InputObject(INPUT_SLOT)
  
  'Get Penetrated Part; temporarily use the slot base plate for the default stiffener side, in case
  'the penetrated part is a profile
  On Error Resume Next
  Dim oPlate As IJPlate
  Dim sStiffenrSide As String
  Dim oPlateClass As New StructDetailObjects.PlatePart
  Set oPlate = oSlotClass.Penetrated
  If oPlate Is Nothing Then
    sStiffenrSide = "Base"
  Else
    Set oPlateClass.object = oSlotClass.Penetrated
    sStiffenrSide = oPlateClass.StiffenerSide
  End If

  Set oPlate = Nothing
  Set oPlateClass = Nothing
  Set oSlotClass = Nothing
  
  'Get Penetrating object
  Dim dProfilePartHeight As Double
  If TypeOf oSlotClass.Penetrating Is IJProfile Then
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlotClass.Penetrating
    dProfilePartHeight = oProfilePart.Height
    Set oProfilePart = Nothing
  ElseIf TypeOf oSlotClass.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance

    oSlotMappingRule.GetSectionDepth oSlotClass.Penetrating, oSlotClass.Penetrated, dProfilePartHeight
    Set oSlotMappingRule = Nothing
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  
  Set oSlotClass = Nothing
  dProfilePartHeight = Round(dProfilePartHeight, 3)
  
  'Get Collar SideOfPart Answer
  Dim strCollarSide As String
   'TR-169311 Trying to retrieve the question's answer
    'without harcoding the path
    GetSelectorAnswer pPLH, gsCollarSideOfPart, strCollarSide
    
  ' Parameter rule:  Set Collar SideOfPart based on Answer
  ' StiffenerSide Base and Offset are internally set and not exposed to the user.
  ' IJDParameterLogic calls PenetrationSymbol, where:
  ' 0 = BASE
  ' 1 = OFFSET
  ' 2 = PLATE_CENTERED
  Select Case strCollarSide
    Case gsNoFlip
        If sStiffenrSide = "Base" Then
            pPLH.Add "SideOfPart", 1
        ElseIf sStiffenrSide = "Offset" Then
            pPLH.Add "SideOfPart", 0
        End If
    Case gsFlip
        If sStiffenrSide = "Base" Then
            pPLH.Add "SideOfPart", 0
        ElseIf sStiffenrSide = "Offset" Then
            pPLH.Add "SideOfPart", 1
        End If
    Case gsCentered
        pPLH.Add "SideOfPart", 2
 End Select
  
  'Parameter rule:  Set Clip Distance
  Dim dClipDistance As Double
  
  dClipDistance = dProfilePartHeight
  
  '  Define Clip Bottom Distance
  If dClipDistance < 0.15 Then
     pPLH.Add "BottomClearance", 0.02
  ElseIf dClipDistance >= 0.15 And dClipDistance < 0.2 Then
     pPLH.Add "BottomClearance", 0.035
  ElseIf dClipDistance >= 0.2 And dClipDistance < 0.35 Then
     pPLH.Add "BottomClearance", 0.05
  ElseIf dClipDistance >= 0.35 And dClipDistance < 0.45 Then
     pPLH.Add "BottomClearance", 0.075
  ElseIf dClipDistance >= 0.45 Then
     pPLH.Add "BottomClearance", 0.1
  End If
    
  '  Distance from Flange to Clip top
  If dClipDistance <= 0.25 Then
     pPLH.Add "TopClearance", 0.05
  Else
     pPLH.Add "TopClearance", 0.07
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
  pOH.SetOutput "SideOfPart"
  pOH.SetOutput "TopClearance"
  pOH.SetOutput "BottomClearance"
  
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


