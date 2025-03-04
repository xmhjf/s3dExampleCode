VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CutWeb1_Parm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID & "EdgeFeatureRules.CutWeb1_Parm"
Const m_ParameterRuleName As String = CUSTOMERID & "EdgeFeatureRules.CutWeb1_Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "StructDetail\SmartOccurrence\" & CUSTOMERID & "EdgeFeatureRules\CutWeb1_Parm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  'Add port inputs
    pIH.SetInput INPUT_EDGE
    pIH.SetInput INPUT_POINT
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

    pOH.SetOutput "Height"
    pOH.SetOutput "Offset"
    pOH.SetOutput "Gap"
    pOH.SetOutput "Width"
    pOH.SetOutput "Radius"
    pOH.SetOutput "Angle"
    pOH.SetOutput "Extra"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
    
    Dim dWeb As Double
    Dim dAngle As Double
    Dim dWidth As Double
    Dim dFlange As Double
    Dim dOffset As Double
    
    Dim dBendAngle As Double
    Dim dBendWidth As Double
    Dim dCutWidth As Double
    Dim dMinCutWidth As Double
    Dim dCutAngle As Double
    
    Dim dOverlap As Double
    Dim dOverlapWidth As Double
    
    Dim oPort As IJPort
    Dim oPoint As Object
    Dim oEdgeFeature As IJStructFeature
    Dim oProfileKnuckle As IJProfileKnuckle
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oSDO_EdgeFeature = New StructDetailObjects.EdgeFeature
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    
    Set oPort = pPRL.InputObject(INPUT_EDGE)
    Set oPoint = pPRL.InputObject(INPUT_POINT)
    Set oEdgeFeature = pPRL.SmartOccurrence
    Set oSDO_EdgeFeature.object = oEdgeFeature
    
    Set oSDO_ProfilePart.object = oPort.Connectable
    dWeb = oSDO_ProfilePart.WebLength
    dFlange = oSDO_ProfilePart.flangeThickness

    If dWeb < 0.001 Then
       dWeb = 0.01
    End If
    
    If dFlange < 0.001 Then
       dFlange = 0.1
    End If

    ' Height is Profile's Flange Length
    pPRL.Add "Height", dWeb
    
    ' Offset is K1 value from table (based on Profile size)
    ' (expect to be .GT. then Web Thickness
    dOffset = dFlange + 0.01
    If dWeb <= 0.1 Then
        dOffset = 0.02
    ElseIf dWeb <= 0.125 Then
        dOffset = 0.02
        
    ElseIf dWeb <= 0.15 Then
        dOffset = 0.025
    ElseIf dWeb <= 0.2 Then
        dOffset = 0.03
        
    ElseIf dWeb <= 0.25 Then
        dOffset = 0.035
    ElseIf dWeb <= 0.3 Then
        dOffset = 0.035
    ElseIf dWeb < 0.35 Then
        dOffset = 0.04
    Else
        dOffset = 0.04
    End If
    pPRL.Add "Offset", dOffset
    
    ' Calculate Knuckle Split Parameters ...
    ' see "Help" Layer on \MarineLibrary\Edge Features\BendFeature_Plate_01.sym
    '   dBendAngle  = oProfileKnuckle.Angle / 2.0
    '   dBendWidth  = distance from center Knuckle Line to outside Knuckle Line at outer edge
    '   dCutWidth   = distance from cut center Line to outside Knuckle Line at outer edge
    '   dMinCutWidth= distance from cut origin to outside Knuckle line at cut origin
    '   dCutAngle   = angle from cut origin to outside Knuckle Line at outer edge
    GetKnuckleCutData oEdgeFeature, oPoint, dFlange, dOffset, _
                      dBendAngle, dBendWidth, dCutWidth, dMinCutWidth, dCutAngle

    dWidth = dWeb / 2#
    If dWidth < dMinCutWidth Then
        dWidth = dMinCutWidth + 0.015
    End If
    
    If False Then
        ' Width is C value from table (based on Profile size ?)
        dOverlapWidth = (dWidth - dMinCutWidth) * Sin((Atn(1) * 2#) - dBendAngle)
        
        pPRL.Add "Gap", dOverlap
        pPRL.Add "Width", dWidth
    Else
        ' Set Width based on amount of Overlap at outer edge
        dOverlap = 0.1
        dOverlapWidth = dMinCutWidth + (dOverlap / Sin((Atn(1) * 2#) - dBendAngle))
    
        pPRL.Add "Gap", dOverlap
        pPRL.Add "Width", dOverlapWidth
    End If

    ' Radius is size of fillets in upper corners (based on Profile size)
    pPRL.Add "Radius", (dWidth / 10#)

    ' Angle is amount of Bend at Feature (based on Profile size)
    ' ... dAngle = ((Atn(1) / 45#) * 10#) (default of 10 degrees)
    pPRL.Add "Angle", dCutAngle

    ' Extra is amount to insure to completely cut Profile at Bend (based on Profile size, Bend Angle)
    pPRL.Add "Extra", (dWeb / 3#)
    
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
    On Error GoTo ErrorHandler
  
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
  Exit Sub

ErrorHandler:
  Err.Raise LogError(Err, MODULE, "IJDUserSymbolServices_InitializeSymbolDefinition").Number
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

