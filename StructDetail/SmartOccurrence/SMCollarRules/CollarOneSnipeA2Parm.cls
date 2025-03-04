VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOneSnipeA2Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Parameter rule for CollarOneSnipeA2(B2)

' Modify class name to particular rule
Private Const m_sClassName As String = "CollarOneSnipeA2Parm"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
'
Implements IJDUserSymbolServices

Private Sub ParameterRuleInputs(oIH As IJDInputsHelper)
   On Error GoTo ErrorHandler

   oIH.SetInput INPUT_PENETRATING
   oIH.SetInput INPUT_BOUNDINGPLATE
   oIH.SetInput INPUT_SLOT
  
   Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Private Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
   On Error GoTo ErrorHandler

   oOH.SetOutput PARAM_SIDE_OF_PART
   oOH.SetOutput PARAM_OUTER_CORNER_RADIUS
   
   Exit Sub
  
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Private Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
   On Error GoTo ErrorHandler

   ' This collar should always be centered
   Dim dOuterCornerRadius As Double
   Dim nSideOfPart As Integer
   
   nSideOfPart = 2 ' Centered
   dOuterCornerRadius = 0.05
   
   oPRL.Add PARAM_SIDE_OF_PART, nSideOfPart
   oPRL.Add PARAM_OUTER_CORNER_RADIUS, dOuterCornerRadius
   
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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal vDefinitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oPR As IJDSymbolDefinition)
  On Error Resume Next
  
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oPR.IJDInputs.RemoveAllInput
   oPR.IJDRepresentations.RemoveAllRepresentation
  
   Dim oDFact As New DefinitionFactory
   Dim oIH As IJDInputsHelper
   
   oDFact.InitAbstractParameterRule oPR
   Set oIH = New InputHelper
   oIH.Definition = oPR
   oIH.InitAs m_sFamilyProgID
   ParameterRuleInputs oIH
   
   '----------- Define Output and Add it to the representation ------------------
   Dim oOH As IJDOutputsHelper
   
   Set oOH = New OutputHelper
   oOH.Representation = oPR.IJDRepresentations.Item(1)
   oOH.InitAs m_sFamilyProgID
   ParameterRuleOutputs oOH
   
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal sCB As String, _
                                                             ByVal vDP As Variant, _
                                                             ByVal oRM As Object) As Object
   Dim oDFact As New DefinitionFactory
  
   Set IJDUserSymbolServices_InstanciateDefinition = oDFact.InstanciateParameterRule( _
                                                     m_sRuleProgID, _
                                                     sCB, _
                                                     IJDUserSymbolServices_GetDefinitionName(vDP), _
                                                     oRM)
   Set oDFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, _
                                                       ByVal oRepName As String, _
                                                       ByVal oOutputColl As Object, _
                                                       vArrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, _
                                                     ByVal oTransactionMgr As Object) As Boolean
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

