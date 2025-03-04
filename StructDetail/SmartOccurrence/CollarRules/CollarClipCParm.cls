VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarClipCParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_ParameterRuleProgid As String = "CollarRules.CollarClipCParm"
Const m_ParameterRuleName As String = "CollarRules.CollarClipCParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarClipCParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    oIH.SetInput INPUT_PENETRATING
    oIH.SetInput INPUT_BOUNDINGPLATE
    oIH.SetInput INPUT_SLOT
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleLogic(oPLH As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    'Get slot object
    Dim oSlotClass As New StructDetailObjects.Slot
    
    Set oSlotClass.object = oPLH.InputObject(INPUT_SLOT)
  
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
  
    'Get Penetrating object
    Dim oProfileClass As StructDetailObjects.ProfilePart
    Set oProfileClass = New ProfilePart
    Set oProfileClass.object = oPLH.InputObject(INPUT_PENETRATING)
  
    'Get Collar SideOfPart Answer
    Dim strCollarSide As String
    strCollarSide = oPLH.SelectorAnswer("CollarRules.RootClipSel", gsCollarSideOfPart)
  
    ' Parameter rule:  Set Collar SideOfPart based on Answer
    ' StiffenerSide Base and Offset are internally set and not exposed to the user.
    ' IJDParameterLogic calls PenetrationSymbol, where:
    ' 0 = PLATE_MOLDED (offset)
    ' 1 = PLATE_ANTI-MOLDED (base)
    ' 2 = PLATE_CENTERED
    Select Case strCollarSide
        Case gsNoFlip
            If sStiffenrSide = "Base" Then
                oPLH.Add "SideOfPart", 0
            ElseIf sStiffenrSide = "Offset" Then
                oPLH.Add "SideOfPart", 1
            End If
        Case gsFlip
            If sStiffenrSide = "Base" Then
                oPLH.Add "SideOfPart", 1
            ElseIf sStiffenrSide = "Offset" Then
                oPLH.Add "SideOfPart", 0
            End If
        Case gsCentered
            oPLH.Add "SideOfPart", 2
    End Select
  
    'Parameter rule:  Set Clip Distance
    Dim dClipDistance As Double
    
    dClipDistance = oProfileClass.Height

    '  Distance from Flange to Clip top
    If dClipDistance <= 0.25 Then
       oPLH.Add "TopClearance", 0.05
    Else
       oPLH.Add "TopClearance", 0.07
    End If

    Set oPlateClass = Nothing
    Set oSlotClass = Nothing
  
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
    oIH.Definition = oSelector
    oIH.InitAs m_FamilyProgid
    ParameterRuleInputs oIH
    Set oIH = Nothing
  
    '----------- Define Output and Add it to the representation ------------------
    Dim oOH As IJDOutputsHelper
    
    Set oOH = New OutputHelper
    oOH.Representation = oSelector.IJDRepresentations.Item(1)
    oOH.SetOutput "SideOfPart"
    oOH.SetOutput "TopClearance"
    Set oOH = Nothing
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal oRM As Object) As Object
    Dim oSelectorFact As New DefinitionFactory
    
    Set IJDUserSymbolServices_InstanciateDefinition = oSelectorFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), oRM)
    Set oSelectorFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal pRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
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
