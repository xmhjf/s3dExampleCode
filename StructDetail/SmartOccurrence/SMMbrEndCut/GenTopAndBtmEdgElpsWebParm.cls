VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenTopAndBtmEdgElpsWebParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenTopAndBtmEdgElpsWebParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "GenTopAndBtmEdgElpsWebParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    pOH.SetOutput "FaceOffset"
    pOH.SetOutput "TopFlangeThickness"
    pOH.SetOutput "BottomFlangeThickness"
    pOH.SetOutput "TopOffset"
    pOH.SetOutput "BottomOffset"
    pOH.SetOutput "TopEdgeOffset"
    pOH.SetOutput "BottomEdgeOffset"
    pOH.SetOutput "TopMinorAxis"
    pOH.SetOutput "TopMajorAxis"
    pOH.SetOutput "BottomMinorAxis"
    pOH.SetOutput "BottomMajorAxis"

    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim dCuttingDepth As Double
    Dim dMajorAxis As Double
    Dim dMinorAxis As Double
    Dim oPortBounded As IJPort

    sMsg = "Defining ParameterRule Outputs"
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Dim dBoundedFlgThickness As Double
    
    If TypeOf oPortBounded.Connectable Is ISPSMemberPartCommon Then
        Dim oMbrPart As New StructDetailObjects.MemberPart
        Set oMbrPart.object = oPortBounded.Connectable
        dBoundedFlgThickness = oMbrPart.flangeThickness
    ElseIf TypeOf oPortBounded.Connectable Is IJProfile Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = oPortBounded.Connectable
        dBoundedFlgThickness = oProfilePart.flangeThickness
    Else
    
    End If

    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    bTFL = False: bBFL = False: bTFR = False: bBFR = False

    CrossSection_Flanges oPortBounded.Connectable, bTFL, bBFL, bTFR, bBFR
    
    Dim dTopFlangeThickness As Double
    Dim dBottomFlangeThickness As Double
    dTopFlangeThickness = 0.00001 '0.01 mm
    dBottomFlangeThickness = 0.00001 '0.01 mm
     
    If (bTFL) Or (bTFR) Then dTopFlangeThickness = dBoundedFlgThickness
    If (bBFL) Or (bBFR) Then dBottomFlangeThickness = dBoundedFlgThickness

    pPRL.Add "FaceOffset", 0.00002
    pPRL.Add "TopFlangeThickness", dTopFlangeThickness
    pPRL.Add "BottomFlangeThickness", dBottomFlangeThickness
    pPRL.Add "TopOffset", 0.00002
    pPRL.Add "BottomOffset", 0.00002
    pPRL.Add "TopEdgeOffset", 0.00002
    pPRL.Add "BottomEdgeOffset", 0.00002
    
    'define the cutting depth to completely cut the Member cross Section
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
    pPRL.Add "CutDepth", dCuttingDepth

    dMajorAxis = GetMajorAxis(pPRL)
    dMinorAxis = GetMinorAxis(pPRL)
    
    If Abs(dMajorAxis - dMinorAxis) < LINEAR_TOLERANCE_Mbr Then
        'Major axis and minor axis are of equal size: for some reason sym file fails
        'to recompute. As a temporary fix, subtract a small value
        pPRL.Add "MinorAxis", (dMinorAxis - LINEAR_TOLERANCE_Mbr)
    Else
        pPRL.Add "MinorAxis", dMinorAxis
    End If
    
    pPRL.Add "TopMajorAxis", dMajorAxis
    pPRL.Add "BottomMajorAxis", dMajorAxis
    
    pPRL.Add "TopMinorAxis", dMinorAxis
    pPRL.Add "BottomMinorAxis", dMinorAxis

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pPR.IJDInputs.RemoveAllInput
    pPR.IJDRepresentations.RemoveAllRepresentation

    Dim pDFact As New DefinitionFactory
    Dim pIH As IJDInputsHelper
    Dim pOH As IJDOutputsHelper

    pDFact.InitAbstractParameterRule pPR

    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH

    Set pOH = New OutputHelper
    pOH.Representation = pPR.IJDRepresentations.Item(1)
    pOH.InitAs m_FamilyProgid
    ParameterRuleOutputs pOH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
    Dim pPRL As IJDParameterLogic
    Set pPRL = New ParameterLogic
    pPRL.Representation = pRep
    ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
