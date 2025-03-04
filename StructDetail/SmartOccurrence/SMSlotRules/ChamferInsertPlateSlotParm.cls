VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChamferInsertPlateSlotParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Parameter rule for open slots
'
Private Const m_sClassName As String = "ChamferInsertPlateSlotParm" ' Modify this to particular rule

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\" + m_sClassName + ".cls"
Const PI As Double = 3.141592654

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
       
    Dim dSetBack As Double
    dSetBack = 0.00001
    
    Dim dAngle As Double
    Dim strChamferMeasurement As String
    
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = oPLH.SmartOccurrence
     
    Dim dThickDifference As Double
    Dim oParentSmartOcc As Object
     
    'Get distance between plate and Member Flange
    GetThicknessDiffOfPlateOverMbrFlange oSlot.Penetrated, oSlot.Penetrating, dThickDifference
    
    'Get Parent Smart Occurence and Selector Question Answers
    GetSmartOccurrenceParent oPLH.SmartOccurrence, oParentSmartOcc
    GetSelectorAnswer oParentSmartOcc, "ChamferValue", dAngle
    GetSelectorAnswer oParentSmartOcc, "ChamferMeasurement", strChamferMeasurement

    'Calculate Slot Angle Based on Chamfer Value and ChamferMeasurement
    If dAngle < 0.01 Then dAngle = 0.01
    If strChamferMeasurement = "Slope" Then
        oPLH.Add "Angle", Atn(dAngle)
    Else
        oPLH.Add "Angle", dAngle * PI / 180
    End If
    
    oPLH.Add PARAM_SET_BACK, dSetBack
    oPLH.Add "Depth", dThickDifference '10mm
    
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
   oOH.SetOutput PARAM_SET_BACK
   oOH.SetOutput "Angle"
   oOH.SetOutput "Depth"
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
