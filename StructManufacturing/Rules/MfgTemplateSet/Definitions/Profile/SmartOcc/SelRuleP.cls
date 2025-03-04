VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelRuleP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Process Selection Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************


Option Explicit

Const MODULE = "SelRuleP"

Const m_SelectorProgid As String = "StrMfgTemplateSelectorProfile.SelRuleP"
Const m_SelectorName As String = "StrMfgTemplateSelectorProfile.SelRuleP"
Const m_FamilyProgid As String = ""
Private m_oCatalogConnection As IJDPOM

'- ProfilePart ------------------------------------------------------'
Private Const IJProfilePart = "{69F3E7BF-40A0-11D2-B324-080036024603}"

Implements IJDUserSymbolServices
Implements IJMfgTemplatePartInputs

Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = "SelectorInputs"
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PROFILE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = "SelectorQuestions"
On Error GoTo ErrorHandler
  
  'No questions defined
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
Const METHOD = "SelectorLogic"
On Error GoTo ErrorHandler

'    Dim sERROR As String
'    sERROR = ""
    
    ' If the input PlatePart is in 'Light' Representation, No Item Should be Selected.
    ' This will delete the Smart Occurrence.
    Dim oPartGeometryState As IJPartGeometryState
    Dim partGeomState As PartGeometryStateType
    
    Set m_oCatalogConnection = GetCatalogResourceMgr
    Set oPartGeometryState = pSL.InputObject(INPUT_PROFILE)
    
    If oPartGeometryState.PartGeometryState = LightPart Then
        Set oPartGeometryState = Nothing
        Exit Sub
    End If
    Set oPartGeometryState = Nothing

    If IsSmartItemExist(m_oCatalogConnection, "Default_TemplateProcessProfile") = True Then
        pSL.Add "Default_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Frame_TemplateProcessProfile") = True Then
        pSL.Add "Frame_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Perpendicular_TemplateProcessProfile") = True Then
        pSL.Add "Perpendicular_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "UserDefined_TemplateProcessProfile") = True Then
        pSL.Add "UserDefined_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Even_TemplateProcessProfile") = True Then
        pSL.Add "Even_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "CenterLine_TemplateProcessProfile") = True Then
        pSL.Add "CenterLine_TemplateProcessProfile"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "AftForward_TemplateProcessProfile") = True Then
        pSL.Add "AftForward_TemplateProcessProfile"
    End If
    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pSL.ReportError sERROR
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
    
    ' Below filter criteria selects Profile parts that are detailed and NOT stright.
    IJMfgTemplatePartInputs_GetPartFilterCriteria = IJProfilePart _
                        & " AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart] " _
                        & " AND [GSCADSTRMFGTemplateCmd.CTemplateFilter,IsAMfgParent] " _
                        & " AND NOT [GSCADSTRMFGTemplateCmd.CTemplateFilter,IsAStraightProfilePart] "
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
