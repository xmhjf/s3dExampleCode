VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MrkBoxPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Marking Parameter rule for Default_TemplateMarkingPlate smart item
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Private m_SupportPlatePart As IJMfgWrapPlatePart

Const MODULE = "MrkBoxPar"

Const m_ParameterRuleProgid As String = "StrMfgTemplateSelectorPlate.MrkBoxPar"
Const m_ParameterRuleName As String = "StrMfgTemplateSelectorPlate.MrkBoxPar"
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
    
    Dim oPlate As Object
    Set oPlate = pPLH.InputObject(INPUT_PLATE)
      
    pPLH.Add "SideMark", 5200 'Apply
    pPLH.Add "SeamMarks", 5210 'Apply
    pPLH.Add "BaseCtlLineMark", 5220 'Apply
    pPLH.Add "SightLineMark", 5230 'Apply
    pPLH.Add "ShipDirectionMark", 5240 'Apply
    pPLH.Add "LabelMark", 5250 'Apply
    pPLH.Add "FrameMarks", 5260 'Apply
    pPLH.Add "KnuckleMarks", 5270 'Apply
    pPLH.Add "ReferenceMarks", 5280 'Apply
    pPLH.Add "TemplateMarks", 5292 'SlotAsCut
    pPLH.Add "CustomMarks", 52100 'Apply
    
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
   
    
    pOH.SetOutput "SideMark"
    pOH.SetOutput "SeamMarks"
    pOH.SetOutput "BaseCtlLineMark"
    pOH.SetOutput "SightLineMark"
    pOH.SetOutput "ShipDirectionMark"
    pOH.SetOutput "LabelMark"
    pOH.SetOutput "FrameMarks"
    pOH.SetOutput "KnuckleMarks"
    pOH.SetOutput "ReferenceMarks"
    pOH.SetOutput "TemplateMarks"
    pOH.SetOutput "CustomMarks"

  
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


 
