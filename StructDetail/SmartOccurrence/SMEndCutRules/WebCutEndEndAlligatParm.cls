VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutEndEndAlligatParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of the smart occurrence.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
Option Explicit

Private Const m_sClassName As String = "WebCutEndEndAlligatParm"
'
Private Const m_sRuleProgID As String = PROJECT_NAME & "." & m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgid As String = ""
Private Const MODULE = PROJECT_PATH & m_sClassName & ".cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType

Private Sub ParameterRuleInputs(oIH As IJDInputsHelper)
   On Error GoTo ErrorHandler

   oIH.SetInput INPUT_BOUNDING
   oIH.SetInput INPUT_BOUNDED
  
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Private Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
   On Error GoTo ErrorHandler

   oOH.SetOutput "CutDepth"
   oOH.SetOutput "CutLengthU"
   oOH.SetOutput "CutLengthV"
   oOH.SetOutput "Slope"
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Private Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
   On Error GoTo ErrorHandler
     
   'Define cut depth as three times the Web Thickness
   Dim dCutDepth As Double
   Dim dLengthU As Double
   Dim dLengthV As Double
   Dim dSlope As Double
   
   Dim oBoundedPort As IJPort
   Dim oBoundingPort As IJPort
   
   Set oBoundedPort = oPRL.InputObject(INPUT_BOUNDED)
   Set oBoundingPort = oPRL.InputObject(INPUT_BOUNDING)
   
   dCutDepth = GetBoundedCutDepth(oBoundedPort.Connectable, False) * 3#
   
   EstimateAlligatorWebCutParameters oPRL.SmartOccurrence, _
                                     dLengthU, _
                                     dLengthV, _
                                     dSlope

   oPRL.Add "CutDepth", dCutDepth
   oPRL.Add "CutLengthU", dLengthU
   oPRL.Add "CutLengthV", dLengthV
   oPRL.Add "Slope", dSlope
   
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
   IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oPR As IJDSymbolDefinition)
   On Error Resume Next
  
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oPR.IJDInputs.RemoveAllInput
   oPR.IJDRepresentations.RemoveAllRepresentation

   Dim oDFact As New DefinitionFactory
   
   oDFact.InitAbstractParameterRule oPR
   
   Dim oIH As IJDInputsHelper
   
   Set oIH = New InputHelper
   oIH.Definition = oPR
   oIH.InitAs m_sFamilyProgid
   ParameterRuleInputs oIH
   
   Dim oOH As IJDOutputsHelper
   
   Set oOH = New OutputHelper
   oOH.Representation = oPR.IJDRepresentations.Item(1)
   oOH.InitAs m_sFamilyProgid
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
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, _
                                                       ByVal sRepName As String, _
                                                       ByVal oOutputColl As Object, _
                                                       arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, _
                                                     ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(oRep As IJDRepresentation)
   Dim oPRL As IJDParameterLogic
   
   Set oPRL = New ParameterLogic
   oPRL.Representation = oRep
   ParameterRuleLogic oPRL
End Sub

'*****************************************************************************************************
' Method Name   : IJEndCutDrawingType_UpdateDrawingType
' Inputs        : feature occurence object
' Output        : string, returns the end cut type.
' Description   : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    IJEndCutDrawingType_UpdateDrawingType = Sniped_No_Offset_WebCuts
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, IJEndCutDrawingType_UpdateDrawingType
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


