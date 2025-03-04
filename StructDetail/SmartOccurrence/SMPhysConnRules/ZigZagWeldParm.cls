VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZigZagWeldParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "PhysConnRules.ZigZagWeldParm"
Const m_ParameterRuleName As String = CUSTOMERID + "PhysConnRules.ZigZagWeldParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\ZigZagWeldParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
  
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NoseMethod"
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "Pitch"
  pOH.SetOutput "Length"
  pOH.SetOutput "MoldedFillet"
  pOH.SetOutput "AntiMoldedFillet"
  pOH.SetOutput "FilletMeasureMethod"
  pOH.SetOutput "Category"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  
  ' call sub to declare the computed TEE outputs for IJWeldSymbol
  AddWeldParmRuleOutputs ZIG_ZAG_WELD, pOH
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
  ' Get Class Arguments
  Dim oPhysConn As New StructDetailObjects.PhysicalConn
  Set oPhysConn.object = pPRL.SmartOccurrence
  
    'Get data required for Parameter Rule
  Dim sCategory As String
  Dim sBevelMethod As String
  Dim sClassSociety As String
  Dim dThickness1 As Double
  Dim dThickness2 As Double
   
  Get_ParameterRuleData pPRL, "ZigZagWeld", _
                        sClassSociety, sCategory, sBevelMethod, _
                        dThickness1, dThickness2

  pPRL.Add "Nose", dThickness1
  pPRL.Add "NoseOrientationAngle", 1.570796
  
  'Calculate Pitch and Length based on Class Society
  Dim dPitch As Double
  Dim dLength As Double
  Select Case sClassSociety
    Case gsLloyds
      dPitch = 0.3
      dLength = 0.075
    Case gsABS
      dPitch = 0.31
      dLength = 0.08
    Case gsDNV
      dPitch = 0.32
      dLength = 0.09
  End Select
            
  ' store the values, for now.  May be overridden by a user defined value later
  pPRL.Add "Pitch", dPitch
  pPRL.Add "Length", dLength
  
  'Calculate the fillet size
  Dim dFillet As Double
  If dThickness1 > dThickness2 Then
    dFillet = dThickness2 * 0.82
  Else
    dFillet = dThickness1 * 0.82
  End If

  'Get the part names
  Dim sPartName1 As String
  Dim oNamedItem1 As IJNamedItem
  Set oNamedItem1 = oPhysConn.ConnectedObject1
  sPartName1 = oNamedItem1.Name
  
  pPRL.Add "RefPartName", sPartName1

  'set the reference side
  Dim sRefSide As String
  sRefSide = GetRefSide(oPhysConn.object)
  Dim bRefIsMolded As Boolean
  If sRefSide = "Base" Then
    pPRL.Add "ReferenceSide", "molded"
    bRefIsMolded = True
  ElseIf sRefSide = "Offset" Then
    pPRL.Add "ReferenceSide", "antimolded"
    bRefIsMolded = False
  Else
    pPRL.Add "ReferenceSide", sRefSide
    bRefIsMolded = True
  End If

  pPRL.Add "MoldedFillet", dFillet
  pPRL.Add "AntiMoldedFillet", dFillet
  pPRL.Add "FilletMeasureMethod", 65536
  
  pPRL.Add "Category", 65541
  pPRL.Add "NoseMethod", 65536
  
  ' fill in the correct values for IJWeldSymbol
  ' this method will include a check for any of the input parameters from the bevel
  ' that have been overridden by the user
  SetCalculatedTeeWeldParams pPRL, ZIG_ZAG_WELD, bRefIsMolded, 0#, 0#, 0#, 0#, pcr_WG_None, _
                             dFillet, dFillet, 0#, dLength, dPitch

  Set oPhysConn = Nothing

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
     
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation
 
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************






