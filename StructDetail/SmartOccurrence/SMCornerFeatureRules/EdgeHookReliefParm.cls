VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdgeHookReliefParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : EdgeHookReliefParm.cls
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
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'
'    22/Sep/2011 - svsmylav TR-202526: 1. edge id of CF is corrected to JXSEC_BOTTOM_FLANGE_RIGHT/JXSEC_TOP_FLANGE_RIGHT in If condition
'                  in GetProjBoundingFlangeThickness method (earlier check was using bottom flange top/top flange bottom: this won't work for CF on outside corner).
'                  2. Added three AC item names to Select Case for 'gsFaceToEdge' answer,
'                  3. Codelist table name is corrected to 'ShapeAtEdgeOverlap'.
'
'*********************************************************************************************
Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.EdgeHookReliefParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.EdgeHookReliefParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\EdgeHookReliefParm.cls"

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
    
    pOH.SetOutput "OffsetU"
    pOH.SetOutput "OffsetV"
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
    
    ' -----------------------------------------------------------
    ' Get the assembly connection (these are not placed manually)
    ' -----------------------------------------------------------
    Dim sGrandParentItemName As String
    Dim oGrandParentObj As Object
    sGrandParentItemName = GetGrandParentName(oPRL.SmartOccurrence, oGrandParentObj)

    ' ------------------------------------
    ' Determine where the feature is going
    ' ------------------------------------
    Dim bIsOnBottomFlange As Boolean
    Dim eLocation As eEndCutCornerLocation
    
    GetCornerFeaturePositionOnEndCut oPRL.SmartOccurrence, eLocation, bIsOnBottomFlange
    
    Dim isBottomEdge As Boolean
    If eLocation = BottomEdgeInside Or eLocation = BottomEdgeOutside Then
        isBottomEdge = True
    End If

    ' ---------------------------
    ' Get the overlap at the edge
    ' ---------------------------
    Dim dOverlap As Double
    Dim dClearance As Double
    Dim dEdgeLength As Double
    
    GetEdgeOverlapAndClearance oGrandParentObj, isBottomEdge, bIsOnBottomFlange, dOverlap, , dClearance, , dEdgeLength

    ' ------------------------------------------------------------------
    ' If the overlap is negative, calculate the flange (not web) overlap
    ' ------------------------------------------------------------------
    Dim dBoundedFlangeThickness As Double
    GetFlangeThickness oPRL.SmartOccurrence, dBoundedFlangeThickness
    
    If dOverlap < 0# Then
        If dBoundedFlangeThickness - dOverlap > dEdgeLength Then
            dOverlap = dEdgeLength
        Else
            dOverlap = dBoundedFlangeThickness - dOverlap
        End If
    End If
    
    ' -----------------------
    ' Get the selector answer
    ' -----------------------
    Dim sBottomAnswerCol As String
    Dim sTopAnswerCol As String
    Dim sBottomShape As String
    Dim sTopShape As String
    
    GetMemberACTopAndBottomShape oGrandParentObj, sBottomAnswerCol, sBottomShape, sTopAnswerCol, sTopShape
    
    Dim sAnswer As String
    If isBottomEdge Then
        sAnswer = sBottomShape
    Else
        sAnswer = sTopShape
    End If
        
    ' ------------------------
    ' Default V-Offset is zero
    ' ------------------------
    Dim dVOffset As Double
    dVOffset = 0.00001
    
    Select Case LCase(sAnswer)
        ' ---------------------------------------------------
        ' If to the edge, set the VOffset to half the overlap
        ' ---------------------------------------------------
        Case LCase(gsFaceToEdge)

            dVOffset = dOverlap / 2#
            
        ' -------------------------------------------------
        ' If to the outside corner, set to the full overlap
        ' -------------------------------------------------
        Case LCase(gsFaceToOutsideCorner)

            dVOffset = dOverlap
            
    End Select
    
    If dVOffset < 0.00001 Then
        dVOffset = 0.00001
    End If
    
    oPRL.Add "OffsetV", dVOffset
    
    ' --------------
    ' Set the radius
    ' --------------
    ' We want to use standard radii
    ' We want to have a minimum flange clearance (future: based on the bounded member depth)
    ' We need to check if there is sufficient material for the desired flange clearance
    ' We need to cover the v offset
    ' Find the smallest radius that will cover the offset and provide the clearance
    
    Dim dStandardRadii(10) As Double
    dStandardRadii(1) = 0.015
    dStandardRadii(2) = 0.025
    dStandardRadii(3) = 0.035
    dStandardRadii(4) = 0.05
    dStandardRadii(5) = 0.075
    dStandardRadii(6) = 0.1
    dStandardRadii(7) = 0.15
    dStandardRadii(8) = 0.2
    dStandardRadii(9) = 0.25
    dStandardRadii(10) = 0.3
    
    Dim dRadius As Double
    dRadius = 0.05
    
    ' If the smallest standard radius is too large, don't use a standard radius
    ' Don't use more than 1/4 the available clearance
    If (2 * dStandardRadii(1) - dVOffset) > dClearance / 4# Then
        dRadius = dClearance / 8#
    Else
        Dim i As Long
        For i = 2 To 10
            ' If the next largest is too much for the available clearance, use the next smallest one
            ' If that is too small for the overlap, then don't use a standard radius
            If (2 * dStandardRadii(i) - dVOffset) > dClearance / 4# Then
                If dStandardRadii(i - 1) < dVOffset / 2# Then
                    dRadius = (dClearance + dVOffset) / 2#
                Else
                    dRadius = dStandardRadii(i - 1)
                End If
                
                Exit For
            ' If the next largest provides an opening that is at least 1/4 the clearance or 50mm, use it
            ElseIf (2 * dStandardRadii(i) - dVOffset) > 0.05 Or (2 * dStandardRadii(i) - dVOffset) > dClearance / 4# Then
                dRadius = dStandardRadii(i)
                Exit For
            End If
        Next i
    End If
 
    oPRL.Add "Radius", dRadius
    oPRL.Add "OffsetU", 0.00001
        
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

