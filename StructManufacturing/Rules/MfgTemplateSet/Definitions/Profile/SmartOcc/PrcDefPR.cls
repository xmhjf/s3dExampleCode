VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrcDefPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Process Default Parametric Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Private m_SupportPlatePart As IJMfgWrapPlatePart

Const MODULE = "DefProPR"

Const m_ParameterRuleProgid As String = "StrMfgTemplateSelectorProfile.PrcDefPR"
Const m_ParameterRuleName As String = "StrMfgTemplateSelectorProfile.PrcDefPR"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = "ParameterRuleLogic"
On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PLATE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  Const METHOD = "ParameterRuleLogic"
  On Error GoTo ErrorHandler
    
    
    pPLH.Add "MinHeight", 5300 'Fixed
    pPLH.Add "MaxHeight", 5310 'Fixed
    pPLH.Add "Extension", 5320  'Fixed
    pPLH.Add "Side", 5330 'side
    pPLH.Add "Type", 5340 'Frame
    pPLH.Add "Orientation", 5350    'AlongFrame
    pPLH.Add "Position", 5360  'FramesAndEdges
    pPLH.Add "BasePlane", 5372        'By System
    pPLH.Add "Direction", 5380   'Transversal
    pPLH.Add "TemplateService", 5390
    pPLH.Add "UserDefinedValues", 53100 'DefaultValues
    pPLH.Add "TemplateNaming", 53110
    
    
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
  pOH.SetOutput "Type"
  pOH.SetOutput "Orientation"
  pOH.SetOutput "Position"
  pOH.SetOutput "BasePlane"
  pOH.SetOutput "Direction"
  pOH.SetOutput "TemplateService"
  pOH.SetOutput "UserDefinedValues"
  pOH.SetOutput "TemplateNaming"
  
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


 
