VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrFreeWebAngledParm"
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

Const m_sClassName As String = "MbrFreeWebAngledParm"
Const m_FamilyProgid As String = ""
Const m_ParameterRuleProgid As String = m_sProjectName + "." + m_sClassName
Const m_ParameterRuleName As String = m_ParameterRuleProgid


Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "Mbr_EndCutRules\MbrWebSnipeParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
    
  pOH.SetOutput "CutDepth"
  pOH.SetOutput "SetBack"
  pOH.SetOutput "TopNose"
  pOH.SetOutput "BottomNose"
  pOH.SetOutput "Slope"
  pOH.SetOutput "CutVang"

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
     On Error GoTo ErrorHandler

     Dim oPortBounded As IJPort
     Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    
    'define the cutting depth as four(4) times the Web Thickness
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
    pPRL.Add "CutDepth", dCuttingDepth
    
    Dim oBoundedObject As StructDetailObjects.MemberPart
    Set oBoundedObject = New StructDetailObjects.MemberPart
    Set oBoundedObject.object = oPortBounded.Connectable
    
    Dim bTFL As Boolean, bTFR As Boolean, bBFL As Boolean, bBFR As Boolean
    CrossSection_Flanges oBoundedObject.object, bTFL, bBFL, bTFR, bBFR
    
    Dim dSetBack As Double
    dSetBack = 0.035
    
    'set the slope value
    Dim dSlope As Double
    ' No requirements specified for now
    dSlope = 2 * Atn(1) '90 degrees
    
    pPRL.Add "Slope", dSlope
    pPRL.Add "SetBack", dSetBack
    
    Dim dTopNose As Double
    Dim dBtmNose As Double
    
    If (bTFL) Or (bTFR) Then
        dTopNose = oBoundedObject.FlangeThickness
    Else
        dTopNose = 0.015
    End If
    
    If (bBFL) Or (bBFR) Then
        dBtmNose = oBoundedObject.FlangeThickness
    Else
        dBtmNose = 0.015
    End If
    
    pPRL.Add "TopNose", dTopNose
    pPRL.Add "BottomNose", dBtmNose
    pPRL.Add "CutVang", (2 * Atn(1)) / 9 '10degrees
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **
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









