VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TeeWeldSquareParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "PhysConnRules.TeeWeldSquareParm"
Const m_ParameterRuleName As String = CUSTOMERID + "PhysConnRules.TeeWeldSquareParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\TeeWeldSquareParm.cls"

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
  
  pOH.SetOutput "RefSideFirstBevelDepth"        'User code "A"
  pOH.SetOutput "RefSideFirstBevelMethod"
  pOH.SetOutput "RefSideFirstBevelAngle"        'User code "a"
  pOH.SetOutput "AntiRefSideFirstBevelDepth"    'User code "B"
  pOH.SetOutput "AntiRefSideFirstBevelMethod"
  pOH.SetOutput "AntiRefSideFirstBevelAngle"    'User code "b"
  
  pOH.SetOutput "MoldedFillet"
  pOH.SetOutput "AntiMoldedFillet"
  pOH.SetOutput "FilletMeasureMethod"
  pOH.SetOutput "Category"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  
  AddWeldParmRuleOutputs TEE_WELD_SQUARE, pOH
  
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
   
  Get_ParameterRuleData pPRL, TEE_WELD_SQUARE, _
                        sClassSociety, sCategory, sBevelMethod, _
                        dThickness1, dThickness2

  'calculate NoseOrientationAngle, depending on the bevel method
  pPRL.Add "NoseOrientationAngle", 1.570796
  pPRL.Add "NoseMethod", 65536
  pPRL.Add "Nose", dThickness1
  
  'Get the part names
  Dim sPartName1 As String
  Dim oNamedItem1 As IJNamedItem
  Set oNamedItem1 = oPhysConn.ConnectedObject1
  sPartName1 = oNamedItem1.Name
  
  pPRL.Add "RefPartName", sPartName1
  
  'Calculate the fillet size
  Dim dFillet As Double
  dFillet = dThickness1 * 0.82

  'get the reference side
  Dim sRefSide As String
  Dim bRefIsMolded As Boolean
  bRefIsMolded = True
  sRefSide = GetRefSide(oPhysConn.object)
  If sRefSide = "Base" Then
        pPRL.Add "ReferenceSide", "molded"
  ElseIf sRefSide = "Offset" Then
        bRefIsMolded = False
        pPRL.Add "ReferenceSide", "antimolded"
  Else
        pPRL.Add "ReferenceSide", sRefSide
  End If

  pPRL.Add "MoldedFillet", dFillet
  pPRL.Add "AntiMoldedFillet", dFillet
  pPRL.Add "FilletMeasureMethod", 65536
          
  If sCategory = gsNormal Then
        pPRL.Add "Category", 65537
  ElseIf sCategory = gsDeep Then
        pPRL.Add "Category", 65538
  ElseIf sCategory = gsFull Then
        pPRL.Add "Category", 65539
  ElseIf sCategory = gsChain Then
        pPRL.Add "Category", 65540
  ElseIf sCategory = gsZigZag Then
        pPRL.Add "Category", 65541
  Else
        pPRL.Add "Category", 65536
  End If
  
    ' variables to compute the values for the angle and depth on the two sides
    Dim dRefSideFirstBevelAngle As Double
    Dim dRefSideFirstBevelDepth As Double
    Dim dAntiRefSideFirstBevelAngle As Double
    Dim dAntiRefSideFirstBevelDepth As Double
    
    pPRL.Add "RefSideFirstBevelDepth", 0#
    pPRL.Add "RefSideFirstBevelAngle", 0#
    pPRL.Add "RefSideFirstBevelMethod", 65536
    
    pPRL.Add "AntiRefSideFirstBevelDepth", 0#
    pPRL.Add "AntiRefSideFirstBevelAngle", 0#
    pPRL.Add "AntiRefSideFirstBevelMethod", 65536
  
    ' fill in the correct values for IJWeldSymbol
    ' this method will include a check for any of the input parameters from the bevel
    ' that have been overridden by the user
    SetCalculatedTeeWeldParams pPRL, TEE_WELD_SQUARE, bRefIsMolded, _
                               0#, 0#, 0#, 0#, 0#, _
                               dFillet, dFillet, 0#, 0#, 0#
  
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
