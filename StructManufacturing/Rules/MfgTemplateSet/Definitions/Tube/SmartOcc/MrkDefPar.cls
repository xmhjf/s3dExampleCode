VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultMarkPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project      : StrMfgTemplateSelTube
'  File         : DefaultMarkPar.cls
'
'  Description  : Default Tube Template Marking Parameter rule
'
'  History       :
'      Siva        2nd September 2011    created
'
'
'**********************************************************************************************************
Option Explicit

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Const MODULE = "DefaultMarkPar"

Const m_ParameterRuleProgid As String = "StrMfgTemplateSelTube.DefaultMarkPar"
Const m_ParameterRuleName As String = "StrMfgTemplateSelTube.DefaultMarkPar"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = "ParameterRuleLogic"
On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_TUBE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  Const METHOD = "ParameterRuleLogic"
  On Error GoTo ErrorHandler
    
    Dim oPlate As Object
    Set oPlate = pPLH.InputObject(INPUT_TUBE)
    
    pPLH.Add "BaseCtrlLineMark", 7200   'Apply
    pPLH.Add "FittingMark", 7200        'Apply
    pPLH.Add "FrameMarks", 7200         'Apply
    pPLH.Add "QuarterLineMarks", 7200   'Apply
    pPLH.Add "SeamMarks", 7200           'Apply
    pPLH.Add "ShipDirectionMark", 7200  'Apply
    pPLH.Add "CustomMark", 7201         'Ignore
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractParameterRule pSelector
 
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)
       
  pOH.SetOutput "BaseCtrlLineMark"
  pOH.SetOutput "FittingMark"
  pOH.SetOutput "FrameMarks"
  pOH.SetOutput "QuarterLineMarks"
  pOH.SetOutput "SeamMarks"
  pOH.SetOutput "ShipDirectionMark"
  pOH.SetOutput "CustomMark"

  DefineViewPrefix pSelector, "INGR"
  Set pOH = Nothing

End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPLH As IJDParameterLogic
  Set pPLH = New ParameterLogic
  pPLH.Representation = pRep
  ParameterRuleLogic pPLH
End Sub
 
