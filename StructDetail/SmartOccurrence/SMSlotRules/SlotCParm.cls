VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotCParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_ParameterRuleProgid As String = "SlotRules.SlotCParm"
Const m_ParameterRuleName As String = "SlotRules.SlotCParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SlotRules\SlotCParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "SlotAngle"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub


Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler
    
    'Get the answer to the assembly method question
    Dim sAssyMethod As String
    sAssyMethod = pPLH.SelectorAnswer("SlotRules.SlotRootSel", "AssyMethodInSlot")
      
    Select Case sAssyMethod
        Case "Slide", "Default value"
            'do nothing to set the angle
            pPLH.Add "SlotAngle", 0.0001

        Case "Drop", "Drop at angle", "Vertical drop"
            'first, get the First Meet assembly
            Dim oAssembly As IJAssembly
            Dim oObject1 As Object
            Dim oObject2 As Object
            
            Dim oSlot As New StructDetailObjects.Slot
            Set oSlot.object = pPLH.SmartOccurrence
            Set oObject1 = oSlot.Penetrated
            Set oObject2 = oSlot.Penetrating
        
            Dim pCHelper As New StructDetailObjects.Helper
            Set oAssembly = pCHelper.FirstMeet(oObject1, oObject2)
        
            Dim oPlanWrapper As New PlanningObjects.PlnAssembly
            Set oPlanWrapper.object = oAssembly
            
            'then, get the angle
            Dim dAssyAngle As Double
            dAssyAngle = oPlanWrapper.SlotOpenAngle(pPLH.SmartOccurrence)
            
            'MsgBox "the angle is:  " & dAssyAngle
            
            If dAssyAngle < 0.0001 And dAssyAngle > -0.0001 Then
                dAssyAngle = 0.00001
            ElseIf dAssyAngle >= 1.570796323 Then
                dAssyAngle = 0.00001
            End If
    
            pPLH.Add "SlotAngle", dAssyAngle
    End Select
 
  
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
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


