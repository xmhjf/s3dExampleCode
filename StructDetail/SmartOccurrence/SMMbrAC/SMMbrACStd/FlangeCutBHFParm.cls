VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeCutBHFParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
'
'  File        : FlangeCutBHFParm.cls
'
'  Description :
'
'  Author      : RPK - Date:11/June/15   (CR-271142)
'
'*********************************************************************************************

Const m_sClassName As String = "FlangeCutBHFParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

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
    
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "BottomFlange"
    pOH.SetOutput "BottomEdgeInsideOffset"
    pOH.SetOutput "BottomEdgeOffset"
    pOH.SetOutput "FlangeOffset"
    pOH.SetOutput "TopLength"
    pOH.SetOutput "Radius"
       
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
        
    ' -----------------------
    ' Set 'BottomFlange' flag
    ' -----------------------
    Dim oEndCutParmCM As Object
    Set oEndCutParmCM = SP3DCreateObject(CUSTOMERID & "MbrEndCut.EndCutParmCM")
    
    oEndCutParmCM.SetBottomFlangeParameter oPRL
        
    ' -----------------
    ' Set cutting depth
    ' -----------------
    Dim dCuttingDepth As Double
    
    dCuttingDepth = GetAppropriateCutDepth(oPRL, oEndCutParmCM.IsBottomFlange(oPRL))
    oPRL.Add "CutDepth", dCuttingDepth
    
    Dim dFlangeOffet As Double
    Dim dEdgeOffset As Double
    Dim dRadius As Double
    Dim dEdgeSetback As Double
    Dim dEdgeInsideOffset As Double
    
    oEndCutParmCM.ComputeHookFlangReliefParameters oPRL, True, False, dEdgeOffset, dRadius, _
                                                dEdgeSetback, dEdgeInsideOffset, dFlangeOffet, dCuttingDepth
       
    oPRL.Add "BottomEdgeInsideOffset", 0.00001  '0.01mm
    oPRL.Add "BottomEdgeOffset", 0.01 '10mm
    oPRL.Add "FlangeOffset", dFlangeOffet
    oPRL.Add "TopLength", 0.02 '20mm
    oPRL.Add "Radius", 0.02 '20mm
       
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
  
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
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid
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
    Dim oPRL As IJDParameterLogic
    Set oPRL = New ParameterLogic
    oPRL.Representation = pRep
    ParameterRuleLogic oPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************





