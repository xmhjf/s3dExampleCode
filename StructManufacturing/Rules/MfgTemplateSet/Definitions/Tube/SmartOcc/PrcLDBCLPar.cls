VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongDstBCLProcPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project      : StrMfgTemplateSelTube
'  File         : LongDstBCLProcPar.cls
'
'  Description  : Long Dist BCL Pos Template process Parameter rule
'
'  History      :
'      Siva        2nd September 2011    created
'
'**********************************************************************************************************
Option Explicit

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"

Const MODULE = "LongDstBCLProcPar"

Const m_ParameterRuleProgid As String = "StrMfgTemplateSelTube.LongDstBCLProcPar"
Const m_ParameterRuleName As String = "StrMfgTemplateSelTube.LongDstBCLProcPar"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = "ParameterRuleLogic"
On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_TUBE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  Const METHOD = "ParameterRuleLogic"
  On Error GoTo ErrorHandler
    
    pPLH.Add "MinHeight", 7100          ' Min Height
    pPLH.Add "MaxHeight", 7100          ' Max Height
    pPLH.Add "Extension", 7100          ' ExtnRad90
    pPLH.Add "Side", 7100               ' Outer
    pPLH.Add "SideEnd", 7100            ' Base
    pPLH.Add "Type", 7101               ' Long Distance BCL
    pPLH.Add "BasePlane", 7100          ' By System
    pPLH.Add "TemplateService", 7100            ' Default
    pPLH.Add "UserDefinedValues", 7100  ' Default
    pPLH.Add "TemplateNaming", 7100     ' Default
    pPLH.Add "Bevel", 7101              ' Ignore
    
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

  pOH.SetOutput "MinHeight"
  pOH.SetOutput "MaxHeight"
  pOH.SetOutput "Extension"
  pOH.SetOutput "Side"
  pOH.SetOutput "SideEnd"
  pOH.SetOutput "Type"
  pOH.SetOutput "BasePlane"
  pOH.SetOutput "TemplateService"
  pOH.SetOutput "UserDefinedValues"
  pOH.SetOutput "TemplateNaming"
  pOH.SetOutput "Bevel"
  
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

 
