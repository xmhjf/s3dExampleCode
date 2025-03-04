VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOnCFParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*********************************************************************************************
'  Copyright (C) 2013, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCollarRules
'  File        : CollarOnLScallopCFParm.cls
'
'  Description :
'   It computes the item paramaters in the context of the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
'  Author      : Alligators
'
'  History     :
'   ---------------------------------------------------------------------
'    DATE        |    AUTHOR    |          COMMENTS
'   ---------------------------------------------------------------------
'   01/APR/2013  -  mchandak    -    Created
'
'*********************************************************************************************
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "CollarRules.CollarOnCFParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CollarRules.CollarOnCFParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "CollarRules\CollarOnCFParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    'Add port inputs
    pIH.SetInput INPUT_CORNERFEATURE
    pIH.SetInput INPUT_CORNERCUTOUTPORT
   
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
    
   pOH.SetOutput "LapDistance"
   pOH.SetOutput "SideOfPart"
   pOH.SetOutput "LongEdges"
   
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
  
    Dim lLongEdges As Long
    Dim lSideOfPart As Long
    
    lSideOfPart = 1
    lLongEdges = 0
    
    Dim oSDOCollar As StructDetailObjectsex.SDOCollar
    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    Set oSDOCollar.object = pPRL.SmartOccurrence
    
    pPRL.Add "LapDistance", 0.05
    pPRL.Add "SideOfPart", lSideOfPart
    pPRL.Add "LongEdges", lLongEdges
    
  
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
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
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************





