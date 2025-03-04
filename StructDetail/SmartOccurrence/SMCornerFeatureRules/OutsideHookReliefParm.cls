VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutsideHookReliefParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : OutsideHookReliefParm.cls
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.OutsideHookReliefParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.OutsideHookReliefParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\OutsideHookReliefParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
    
    pOH.SetOutput "Radius"
    pOH.SetOutput "Radius2"
    pOH.SetOutput "OffsetU"
    pOH.SetOutput "OffsetV"
    pOH.SetOutput "FlangeClearance"
    pOH.SetOutput "Vlength"
    
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
    
    ' ------------------------------------------
    ' Get the endcut creating the corner feature
    ' ------------------------------------------
    Dim oFeatureAsChild As IJDesignChild
    Dim oCornerParent As Object
    
    If TypeOf oPRL.SmartOccurrence Is IJDesignChild Then
        Set oFeatureAsChild = oPRL.SmartOccurrence
        Set oCornerParent = oFeatureAsChild.GetParent
    Else
        Exit Sub
    End If
    
    ' --------------------------------------
    ' Get the position of the corner feature
    ' --------------------------------------
    Dim eLocation As eEndCutCornerLocation
    Dim bIsBottomFlange As Boolean
    
    GetCornerFeaturePositionOnEndCut oPRL.SmartOccurrence, eLocation, bIsBottomFlange

    ' ---------------------------------------------------------------------------------
    ' Get the clearance between the bounding edge and the inside of the bounded flanges
    ' ---------------------------------------------------------------------------------
    Dim dInsideClearance As Double
    Dim dOutsideClearance As Double
    Dim dEdgeLength As Double
    
    Dim bIsBottomEdge As Boolean
    bIsBottomEdge = (eLocation = BottomEdgeInside Or eLocation = BottomEdgeOutside)
    GetEdgeOverlapAndClearance oCornerParent, _
                               bIsBottomEdge, _
                               bIsBottomFlange, _
                               , , _
                               dInsideClearance, _
                               dOutsideClearance, _
                               dEdgeLength
                               
    ' -------------------------------------------------------------------
    ' OffsetV is the bounding edge length measured in the sketching plane
    ' -------------------------------------------------------------------
    oPRL.Add "OffsetV", dEdgeLength
                               
    ' -----------------------------------------------------------------------------------------------
    ' Make the flange clearance equal to 1/4 the inside clearance, in standard increments, up to 50mm
    '     If it gets below 5mm, make it equal to the inside clearance minus 0.01mm
    ' Make the inside radius roughly half the flange clearance, in standard increments
    ' -----------------------------------------------------------------------------------------------
    Dim dFlangeClearance As Double
    Dim dRadius2 As Double
    
    If dInsideClearance > 0.2 Then
        dFlangeClearance = 0.05
        dRadius2 = 0.025
    ElseIf dInsideClearance > 0.14 Then
        dFlangeClearance = 0.035
        dRadius2 = 0.015
    ElseIf dInsideClearance > 0.1 Then
        dFlangeClearance = 0.025
        dRadius2 = 0.01
    ElseIf dInsideClearance > 0.06 Then
        dFlangeClearance = 0.015
        dRadius2 = 0.005
    ElseIf dInsideClearance > 0.04 Then
        dFlangeClearance = 0.01
        dRadius2 = 0.005
    ElseIf dInsideClearance > 0.02 Then
        dFlangeClearance = 0.005
        dRadius2 = 0.00002
    Else
        dFlangeClearance = dInsideClearance - 0.00005
        dRadius2 = dInsideClearance - 0.00005
    End If
    
    oPRL.Add "FlangeClearance", dFlangeClearance
    oPRL.Add "Radius2", dRadius2
    
    ' -----------------------------------------------------------------------------------------
    ' Make the outside radius 1/4 the outside clearance, in standard increments, up to 50mm
    ' Make the outside offset 15mm, reduced down in standard increments for smaller radii
    ' Make the offset from the edge equal to the radius, so the arc is centered over the corner
    ' -----------------------------------------------------------------------------------------
    Dim dRadius As Double
    Dim dVLength As Double
    Dim dOffsetU As Double
    
    If dOutsideClearance > 0.2 Then
        dRadius = 0.05
        dVLength = 0.025
        dOffsetU = dRadius
    ElseIf dOutsideClearance > 0.14 Then
        dRadius = 0.035
        dVLength = 0.015
        dOffsetU = dRadius
    ElseIf dOutsideClearance > 0.1 Then
        dRadius = 0.025
        dVLength = 0.01
        dOffsetU = dRadius
    ElseIf dOutsideClearance > 0.06 Then
        dRadius = 0.015
        dVLength = 0.005
        dOffsetU = dRadius
    ElseIf dOutsideClearance > 0.04 Then
        dRadius = 0.01
        dVLength = 0.005
        dOffsetU = dRadius
    ElseIf dOutsideClearance > 0.02 Then
        dRadius = 0.005
        dVLength = 0.00002
        dOffsetU = dRadius
    Else
        dRadius = 0.00001
        dVLength = 0.00002
        dOffsetU = dRadius
    End If
    
    oPRL.Add "Vlength", dVLength
    oPRL.Add "Radius", dRadius
    oPRL.Add "OffsetU", dOffsetU

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



