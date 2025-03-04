VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShortEdgeRatholeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-2014, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : ShortEdgeRatholeParm.cls
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.ShortEdgeRatholeParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.ShortEdgeRatholeParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\ShortEdgeRatholeParm.cls"

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
  
    pOH.SetOutput "Offset"
    pOH.SetOutput "FlangeClearance"
    pOH.SetOutput "Radius"
    pOH.SetOutput "Angle"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
     
    ' If feature is meant to be from the corner, use the distance between the two surfaces
    ' If feature is meant to be from the edge, use half the distance between the two surfaces
    Dim dOffset As Double
    Dim dInside As Double
    Dim dRemaining As Double
    
    dInside = Abs(GetSurfaceOffsetForEdgeCornerFeature(pPRL.SmartOccurrence))
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    AssemblyConnection_SmartItemName pPRL.SmartOccurrence, sParentItemName, oParentObj
    
    Dim sAnswer As String
    GetSelectorAnswer oParentObj, "ShapeAtEdge", sAnswer
    
    If LCase(sAnswer) = LCase("EdgeToFlange") Then
        dOffset = dInside / 2#
    Else
        dOffset = 0.00001
    End If

    dRemaining = dInside - dOffset

    ' Set radius equal to the distance after the offset, rounded down to nearest 5mm increment, max 50mm
    Dim lMult As Long
    lMult = Int(dRemaining / 0.005)
    
    Dim dRadius As Double
    dRadius = 0.05
    
    If lMult < 10 Then
        dRadius = lMult * 0.005
    End If
    
    ' If less than 5, use the remaining distance
    If dRadius = 0# Then
        If dRemaining >= 0.00003 Then
            dRadius = dRemaining - 0.00002
        Else
            dRadius = 0.00001
        End If
    End If

    pPRL.Add "Offset", dOffset
    pPRL.Add "FlangeClearance", dInside - 0.00001
    pPRL.Add "Radius", dRadius
    pPRL.Add "Angle", Atn(1) * 2

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
