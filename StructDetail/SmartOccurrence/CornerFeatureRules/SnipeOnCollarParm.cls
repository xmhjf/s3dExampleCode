VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnipeOnCollarParm"
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

Const m_ParameterRuleProgid As String = "CornerFeatureRules.SnipeOnCollarParm"
Const m_ParameterRuleName As String = "CornerFeatureRules.SnipeOnCollarParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CornerFeatureRules\SnipeOnCollarParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    oIH.SetInput INPUT_PORT1FACE
    oIH.SetInput INPUT_PORT2EDGE
    oIH.SetInput INPUT_PORT3EDGE
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
    oOH.SetOutput "Ulength"
    oOH.SetOutput "Vlength"
    oOH.SetOutput "Flip"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
     
    '0 = NoFlip
    '1 = Flip
    oPRL.Add "Flip", 0

    'Parameter rule:  Set snipe length
    oPRL.Add "Ulength", 0.008
    oPRL.Add "Vlength", 0.008

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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymbolDefinition As IJDSymbolDefinition)
    On Error Resume Next
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oSymbolDefinition.IJDInputs.RemoveAllInput
    oSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    
    Dim oDefFactory As New DefinitionFactory
    
    oDefFactory.InitAbstractParameterRule oSymbolDefinition
    Set oDefFactory = Nothing
    
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.Definition = oSymbolDefinition
    oIH.InitAs m_FamilyProgid
    ParameterRuleInputs oIH
    Set oIH = Nothing
    
    Dim oOH As IJDOutputsHelper
    
    Set oOH = New OutputHelper
    oOH.Representation = oSymbolDefinition.IJDRepresentations.Item(1)
    oOH.InitAs m_FamilyProgid
    ParameterRuleOutputs oOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal oRM As Object) As Object
    Dim oDefFactory As New DefinitionFactory
    
    Set IJDUserSymbolServices_InstanciateDefinition = oDefFactory.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), oRM)
    Set oDefFactory = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, ByVal pRepName As String, ByVal oOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(oRep As IJDRepresentation)
    Dim oPRL As IJDParameterLogic
    
    Set oPRL = New ParameterLogic
    oPRL.Representation = oRep
    ParameterRuleLogic oPRL
    Set oPRL = Nothing
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
