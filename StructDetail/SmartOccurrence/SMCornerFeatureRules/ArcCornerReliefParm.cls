VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArcCornerReliefParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : ArcCornerReliefParm.cls
'
'  Description :
'       It computes the item paramaters in the context of the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Outputs defined by name the collection of parameters
'
'  Author      : Alligators
'
'  History     :
'    18/APR/2011 - Creation
'
'*********************************************************************************************

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.ArcCornerReliefParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.ArcCornerReliefParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\ArcCornerReliefParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    ' Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
    
End Sub
    
'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
    
    pOH.SetOutput "Radius"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
    
    Const sMETHOD = "ParameterRuleLogic"

    ' --------------------------------------
    ' Get the position of the corner feature
    ' --------------------------------------
    Dim eLocation As eEndCutCornerLocation
    Dim bIsBottomFlange As Boolean
    
    GetCornerFeaturePositionOnEndCut oPRL.SmartOccurrence, eLocation, bIsBottomFlange

    Dim bIsBottomEdge As Boolean
    bIsBottomEdge = (eLocation = BottomEdgeInside Or eLocation = BottomEdgeOutside)
        
    ' --------------------------------------
    ' See how much the web overlaps the edge
    ' --------------------------------------
    Dim sGrandParentItemName As String
    Dim oGrandParentObj As Object
    sGrandParentItemName = GetGrandParentName(oPRL.SmartOccurrence, oGrandParentObj)
        
    Dim dInsideOverlap As Double
    Dim dOutsideOverlap As Double
    Dim dOutsideClearance As Double
    Dim dInsideClearance As Double
    Dim dEdgeLength As Double
    
    GetEdgeOverlapAndClearance oGrandParentObj, _
                               bIsBottomEdge, _
                               bIsBottomFlange, _
                               dInsideOverlap, _
                               dOutsideOverlap, _
                               dInsideClearance, _
                               dOutsideClearance, _
                               dEdgeLength
 
    Dim dOverlap As Double
    
    If eLocation = BottomEdgeInside Or eLocation = TopEdgeInside Then
        dOverlap = dInsideOverlap
    Else
        dOverlap = dOutsideOverlap
    End If
 
    ' ------------------------------------------------------------------
    ' Set to half the overlap, round to 5mm increment, 5mm min, 50mm max
    ' ------------------------------------------------------------------
    Dim lMult As Long
    Dim dRadius As Double
    
    lMult = Int(dOverlap / 0.01)
        
    If lMult < 10 Then
        dRadius = lMult * 0.005
    Else
        dRadius = 0.05
    End If
    
    ' Minimum of 5mm
    If dRadius <= 0# Then
        dRadius = 0.005
    End If
    
    oPRL.Add "Radius", dRadius
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
End Sub
  
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
  
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next
      
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
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
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

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

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************


