VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongScallop1Parm"
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

Const m_ParameterRuleProgid As String = "CornerFeatureRules.LongScallop1Parm"
Const m_ParameterRuleName As String = "CornerFeatureRules.LongScallop1Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CornerFeatureRules\LongScallop1Parm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
pOH.SetOutput "Ulength"
pOH.SetOutput "Radius"
pOH.SetOutput "Flip"

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
'Get corner feature part object
Dim oCornerFeature As New Structdetailobjects.CornerFeature
Set oCornerFeature.object = pPRL.SmartOccurrence

Dim oPart As Object
Set oPart = oCornerFeature.GetPartObject

'Get Corner Feature Orientation Answer
Dim strCornerFlip As String
strCornerFlip = pPRL.SelectorAnswer("CornerFeatureRules.RootCornerSel", gsCornerFlip)

'Parameter rule:  Set general corner size
'**********************************************
 
'first, check the corner gap.  If gap found, it drives the parms
Dim oGapLengthGap As Double
Dim oGapHeightGap As Double
Dim oGapTolerance As Double
Dim dULength As Double
Dim dRadius As Double

oGapTolerance = 0.001 'will not find gaps < 1 mm
If oCornerFeature.MeasureCornerGap(oGapTolerance, _
                                   oGapLengthGap, _
                                   oGapHeightGap) Then
   
    If oGapLengthGap > 0# Then
        'the Ulength value will be increased, no flip
        dULength = oGapLengthGap + 0.025
        'do not flip the corner feature
        pPRL.Add "Flip", 0
    ElseIf oGapHeightGap > 0# Then
        'the Vlength value will be increased, flip
        dULength = oGapHeightGap + 0.025
        'flip the corner feature
        pPRL.Add "Flip", 1
        
    Else
        ' default to No Flip
        pPRL.Add "Flip", 0
    End If
    
    dRadius = 0.035 'default to 35 mm
Else
 
    'Parameter rule: Set corner feature input orientation based on Answer
    'IJDParameterlogic calls FeatureSymbol.ASCornerFeature with parameter Flip, where:
    '0 = NoFlip
    '1 = Flip
    Select Case strCornerFlip
        Case "NoFlip"
            pPRL.Add "Flip", 0
        Case "Flip"
            pPRL.Add "Flip", 1
    End Select
    
    ' no gap found, user selected this feature.  Process based on part type
    If TypeOf oPart Is IJPlate Then
        dULength = 0.1
        dRadius = 0.05
    
    ElseIf TypeOf oPart Is IJStiffener Then 'oPart is a profile
        Dim oProfile As New Structdetailobjects.ProfilePart
        Set oProfile.object = oPart
        Select Case oProfile.WebLength     'Is Select Case (or different parm class) needed for each different types of symbols which will be in General Corner class?
            Case 0 To 0.2
                dULength = 0.07
                dRadius = 0.035
            Case Is > 0.2, Is < 0.4
                dULength = 0.1
                dRadius = 0.05
            Case Is >= 0.4
                dULength = 0.15
                dRadius = 0.075
        End Select
    ElseIf TypeOf oPart Is IJBeam Then 'oPart is a beam
        Dim oBeam As New Structdetailobjects.BeamPart
        Set oBeam.object = oPart
        Select Case oBeam.WebLength     'Is Select Case (or different parm class) needed for each different types of symbols which will be in General Corner class?
            Case 0 To 0.2
                dULength = 0.07
                dRadius = 0.035
            Case Is > 0.2, Is < 0.4
                dULength = 0.1
                dRadius = 0.05
            Case Is >= 0.4
                dULength = 0.15
                dRadius = 0.075
        End Select
    ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
        Dim oMember As New Structdetailobjects.MemberPart
        Set oMember.object = oPart
        Select Case oMember.WebLength
            Case 0 To 0.2
                dULength = 0.07
                dRadius = 0.035
            Case Is > 0.2, Is < 0.4
                dULength = 0.1
                dRadius = 0.05
            Case Is >= 0.4
                dULength = 0.15
                dRadius = 0.075
        End Select
    End If
End If
        
pPRL.Add "Ulength", dULength
pPRL.Add "Radius", dRadius
        
    '************************************************
  
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

