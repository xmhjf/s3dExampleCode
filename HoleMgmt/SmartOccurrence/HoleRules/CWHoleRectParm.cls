VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWHoleRectParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurrence\HoleRules
'
' File: CWHoleRectParm.cls
'
' Author: Hole Mgmt Team
'
' Abstract: parameter rule for cableway hole rectangle trace smart class
'
' Description: This is the parameter rule for the cableway hole trace smart class.
'              It computes the item paramaters in the context of the smart occurrence.
'******************************************************************************

Option Explicit

Const m_ParameterRuleProgid As String = "HoleRules.CWHoleRectParm"
Const m_ParameterRuleName As String = "HoleRules.CWHoleRectParm"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'the hole trace is the only input
    pIH.SetInput INPUT_HOLETRACE
  
    Exit Sub
ErrorHandler:
    pIH.ReportError "ParameterRuleInputs"
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
''    pOH.SetOutput "Hwidth"
''    pOH.SetOutput "Hdepth"
    
    pOH.SetOutput "Horientation"
  
    Exit Sub
ErrorHandler:
    pOH.ReportError "ParameterRuleOutputs"
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    'get the HoleSmartOcc and HoleTraceAE
    Dim oHoleTrace As IJHoleTraceAE
    Dim oHoleSmartOcc As IJHoleSmartOcc
    Set oHoleSmartOcc = pPRL.SmartOccurrence
    Set oHoleTrace = oHoleSmartOcc.GetHoleTraceAE

    'get the cableway - there should be only one if in this method
    Dim oCableway As Object
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    For Each oCableway In oOutfitting
        If Not oCableway Is Nothing Then Exit For
    Next oCableway

    'has the orientation been set
    Dim oRtePathFeat As IJRtePathFeat
    Dim oRteFeatUtility As IJRtePathCrossSectUtility
    
    Set oRtePathFeat = oCableway
    Set oRteFeatUtility = oRtePathFeat
      
    Dim dOrientation As Double
    If Not oRteFeatUtility Is Nothing Then
        dOrientation = oRteFeatUtility.GetOrientationAngle
    End If
    pPRL.Add "Horientation", dOrientation
    
    Set oHoleTrace = Nothing
    Set oHoleSmartOcc = Nothing
    Set oCableway = Nothing
    Set oOutfitting = Nothing
    Set oRtePathFeat = Nothing
    Set oRteFeatUtility = Nothing
    
    '**********************************************
    ' start of old routine
    '**********************************************
''''''''''    Dim dWidth As Double, dDepth As Double
''''''''''
''''''''''    'get the HoleSmartOcc and HoleTraceAE
''''''''''    Dim oHoleTrace As IJHoleTraceAE
''''''''''    Dim oHoleSmartOcc As IJHoleSmartOcc
''''''''''    Set oHoleSmartOcc = pPRL.SmartOccurrence
''''''''''    Set oHoleTrace = oHoleSmartOcc.GetHoleTraceAE
''''''''''
''''''''''    'get the input Hwidth and Hdepth from the symbol
''''''''''    dWidth = pPRL.SmartItemAttribute("Hwidth")
''''''''''    dDepth = pPRL.SmartItemAttribute("Hdepth")
''''''''''
''''''''''    'get the cableway - there should be only one if in this method
''''''''''    Dim oCableway As Object
''''''''''    Dim oOutfitting As IJDObjectCollection
''''''''''    Set oOutfitting = oHoleTrace.GetParentOutfitting
''''''''''    For Each oCableway In oOutfitting
''''''''''        If Not oCableway Is Nothing Then Exit For
''''''''''    Next oCableway
''''''''''
''''''''''    'has the orientation been switched
''''''''''    Dim oRtePathCrossSect As IJRteCrossSectOccur
''''''''''    Dim oRteCrossSectOccur As IJRteCrossSectOccur
''''''''''    Dim bWidthInUDir As Boolean, bFlipAboutU As Boolean, bFlipAboutV As Boolean
''''''''''
''''''''''    Set oRtePathCrossSect = oCableway
''''''''''    Set oRteCrossSectOccur = oRtePathCrossSect.GetCrossSectOccur()
''''''''''    oRteCrossSectOccur.GetCSOrientFlags bWidthInUDir, bFlipAboutU, bFlipAboutV
''''''''''
''''''''''    'if bWidthInUDir is false then need to switch the width and depth
''''''''''    If bWidthInUDir Then
''''''''''        pPRL.Add "Hwidth", dWidth
''''''''''        pPRL.Add "Hdepth", dDepth
''''''''''    Else
''''''''''        pPRL.Add "Hwidth", dDepth
''''''''''        pPRL.Add "Hdepth", dWidth
''''''''''    End If
''''''''''
''''''''''    Set oHoleTrace = Nothing
''''''''''    Set oHoleSmartOcc = Nothing
''''''''''    Set oCableway = Nothing
''''''''''    Set oOutfitting = Nothing
''''''''''    Set oRtePathCrossSect = Nothing
''''''''''    Set oRteCrossSectOccur = Nothing

    Exit Sub
ErrorHandler:
    pPRL.ReportError
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

 
