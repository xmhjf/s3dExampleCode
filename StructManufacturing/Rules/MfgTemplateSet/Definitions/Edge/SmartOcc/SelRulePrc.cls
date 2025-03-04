VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelRuleProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StrMfgTemplateSelEdge
'  File         : SelRuleProc.cls
'
'  Description  : Edge Template Process Selector rule
'
'   Selector:
'   ---------
'   It selects the list of avaible Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parametrize the selection
'
'  History      :
'      Siva        20th April 2012    created
'
'**********************************************************************************************************

Option Explicit

Const MODULE = "SelRuleProc"

Const m_SelectorProgid As String = "StrMfgTemplateSelEdge.SelRuleProc"
Const m_SelectorName As String = "StrMfgTemplateSelEdge.SelRuleProc"
Const m_FamilyProgid As String = ""
'ISPSMemberPartPrismatic (for Common Struct Member Part ) ----------------------
Const ISPSMemberPartPrismatic = "{53B6E606-78C6-4BCA-A640-43A2258EDED1}"
'- ProfilePart ------------------------------------------------------'
Const IJProfilePart = "{69F3E7BF-40A0-11D2-B324-080036024603}"

Implements IJDUserSymbolServices
Implements IJMfgTemplatePartInputs

Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = "SelectorInputs"
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_EDGE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = "SelectorQuestions"
On Error GoTo ErrorHandler
  
  'No questions defined
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
Const METHOD = "SelectorLogic"
On Error GoTo ErrorHandler

    Dim oCatalogConnection As IJDPOM
    Set oCatalogConnection = GetCatalogResourceMgr
    
    If IsSmartItemExist(oCatalogConnection, "Default_TemplateProcessEdge") = True Then
        pSL.Add "Default_TemplateProcessEdge"
    End If
    
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
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector

  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH

  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  DefineViewPrefix pSelector, "INGR"
  SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSL As IJDSelectorLogic
  Set pSL = New SelectorLogic
  pSL.Representation = pRep
  SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Private Function IJMfgTemplatePartInputs_GetPartFilterCriteria() As String
    Const METHOD = "IJMfgTemplatePartInputs_GetPartFilterCriteria"
    On Error GoTo ErrorHandler
    
    IJMfgTemplatePartInputs_GetPartFilterCriteria = IJProfilePart _
                        & " AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart] " _
                        & " AND NOT [GSCADSTRMFGTemplateCmd.CTemplateFilter,IsAStraightProfilePart] "
                        
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
