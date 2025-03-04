VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bevel1Parm"
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

Const m_ParameterRuleProgid As String = "FETRules.Bevel1Parm"
Const m_ParameterRuleName As String = "FETRules.Bevel1Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\FreeEdgeTreatmentRules\Bevel1Parm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
'    'Add port inputs
    pIH.SetInput INPUT_TREATMENT_EDGE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "RefSideFirstBevelDepth"        'User code "A"
  pOH.SetOutput "RefSideFirstBevelAngle"        'User code "a"
  pOH.SetOutput "AntiRefSideFirstBevelDepth"    'User code "B"
  pOH.SetOutput "AntiRefSideFirstBevelAngle"    'User code "b"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR

  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
    Dim sERROR As String
    
    '
    'We need to get the thickness from the part to which this treatment was applied
    '
    
    'Part would be the connectable of the treatment edge
    Dim oPartFromPort As Object
    
    If TypeOf pPRL.SmartOccurrence Is IJShpStrEdgeTreatment Then
        Dim oFET As IJShpStrEdgeTreatment
        Set oFET = pPRL.SmartOccurrence
        
        Dim oFETPort As IJPort
        Set oFETPort = oFET.GetEdge
        
        Set oPartFromPort = oFETPort.Connectable
        
        Set oFETPort = Nothing
        Set oFET = Nothing
    Else
        ' not expected to be here. Exit with error message.
        sERROR = "Treatment object should support IJShpStrEdgeTreatment!!!"
        GoTo ErrorHandler
    End If
     
    'Now get the relevant thickness based on whether it is a plate or profile
    
    Dim dThickness As Double
    
    Dim sFlip As String
    sFlip = pPRL.SelectorAnswer("FETRules.BevelSel", "Flip")
      
    Dim sBevelType As String
    sBevelType = pPRL.SelectorAnswer("FETRules.BevelSel", "BevelType")
     
    Dim oHelper As New StructDetailObjects.Helper
    
    If oHelper.ObjectType(oPartFromPort) = SDOBJECT_PLATE Then
        Dim oPlate As New StructDetailObjects.PlatePart
        Set oPlate.object = oPartFromPort
        
        dThickness = oPlate.PlateThickness
    ElseIf oHelper.ObjectType(oPartFromPort) = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oPartFromPort
        
        dThickness = oProfile.WebThickness
    End If
    
    'Calculate and set the values for the parameters
    
    Dim dBevelAngle As Double
    dBevelAngle = 0.78539816 '45 degrees
    
    Select Case sBevelType
        Case "OneSided"
            If sFlip = "No" Then
                pPRL.Add "Nose", dThickness / 2
                pPRL.Add "NoseOrientationAngle", 1.570796
                pPRL.Add "RefSideFirstBevelDepth", dThickness / 2
                pPRL.Add "RefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "AntiRefSideFirstBevelDepth", 0
                pPRL.Add "AntiRefSideFirstBevelAngle", 0
                pPRL.Add "ReferenceSide", "molded"
            Else
                pPRL.Add "Nose", dThickness / 2
                pPRL.Add "NoseOrientationAngle", 1.570796
                pPRL.Add "RefSideFirstBevelDepth", 0
                pPRL.Add "RefSideFirstBevelAngle", 0
                pPRL.Add "AntiRefSideFirstBevelDepth", dThickness / 2
                pPRL.Add "AntiRefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "ReferenceSide", "antimolded"
            End If
        Case "TwoSided"
            If sFlip = "No" Then
                pPRL.Add "Nose", dThickness / 3
                pPRL.Add "NoseOrientationAngle", 1.570796
                pPRL.Add "RefSideFirstBevelDepth", dThickness / 3
                pPRL.Add "RefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "AntiRefSideFirstBevelDepth", dThickness / 3
                pPRL.Add "AntiRefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "ReferenceSide", "molded"
            Else
                pPRL.Add "Nose", dThickness / 3
                pPRL.Add "NoseOrientationAngle", 1.570796
                pPRL.Add "RefSideFirstBevelDepth", dThickness / 3
                pPRL.Add "RefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "AntiRefSideFirstBevelDepth", dThickness / 3
                pPRL.Add "AntiRefSideFirstBevelAngle", dBevelAngle
                pPRL.Add "ReferenceSide", "antimolded"
            End If
        End Select
    
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic", sERROR).Number
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




