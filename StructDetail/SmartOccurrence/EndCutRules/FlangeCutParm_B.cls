VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeCutParm_B"
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
'****
'   This is for the "B" Cross Section type
'   Because the "B" Cross Section type does not have a Flange Right Edge
'   To calculate the Flange Length, the width of the Bulb is an input to the Flange Cut Symbol
'
Option Explicit

Const m_ParameterRuleProgid As String = "EndCutRules.FlangeCutParm_B"
Const m_ParameterRuleName As String = "EndCutRules.FlangeCutParm_B"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\FlangeCutParm_B.cls"

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
  pOH.SetOutput "BottomFlange"
  pOH.SetOutput "LengthAdjustment"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

     Dim oPortBounded As IJPort
     Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
     Dim oBounded As Object
     Dim phelper As New Structdetailobjects.Helper
     
    'Get Bounded part
    If phelper.ObjectType(oPortBounded.Connectable) = SDOBJECT_STIFFENER Then
        Dim oBoundedProfile As New Structdetailobjects.ProfilePart
        Set oBoundedProfile.object = oPortBounded.Connectable
        Set oBounded = oBoundedProfile
    ElseIf phelper.ObjectType(oPortBounded.Connectable) = SDOBJECT_BEAM Then
        Dim oBoundedBeam As New Structdetailobjects.BeamPart
        Set oBoundedBeam.object = oPortBounded.Connectable
        Set oBounded = oBoundedBeam
    End If
    
    ' Define the cutting depth as one-half the Web Length
    ' The Flange's Thickness is never over Half the Web Length
    Dim dCuttingDepth As Double
    dCuttingDepth = oBounded.WebLength * 0.6
    pPRL.Add "CutDepth", dCuttingDepth

    ' This is for the "B" Cross Section type, Bottom Flange is always "No"
    pPRL.Add "BottomFlange", 0

    ' The total "B" Cross Section Flange width is defined by:
    '   FlangeLength + CornerRadius
    Dim dLengthAdjustment As Double
    dLengthAdjustment = oBounded.FlangeLength + oBounded.CrossSectionParameter("CornerRadius")
    pPRL.Add "LengthAdjustment", dLengthAdjustment

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




