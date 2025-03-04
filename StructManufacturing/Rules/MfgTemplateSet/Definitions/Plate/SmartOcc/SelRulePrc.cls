VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelRulePrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Marking Parameter rule for Default_TemplateProcessPlate smart item
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Const MODULE = "SelRulePrc"

Const m_SelectorProgid As String = "StrMfgTemplateSelectorPlate.SelRulePrc"
Const m_SelectorName As String = "StrMfgTemplateSelectorPlate.SelRulePrc"
Const m_FamilyProgid As String = ""
Private m_oCatalogConnection As IJDPOM

'- PlatePart --------------------------------------------------------'
Private Const IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}"

Implements IJDUserSymbolServices
Implements IJMfgTemplatePartInputs

Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = "SelectorInputs"
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PLATE
    
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
    Set oPartGeometryState = pSL.InputObject(INPUT_PLATE)
    
    If oPartGeometryState.PartGeometryState = LightPart Then
        Set oPartGeometryState = Nothing
        Exit Sub
    End If
    Set oPartGeometryState = Nothing
    
    Dim oPlatePart As StructDetailObjects.PlatePart
    Set oPlatePart = New StructDetailObjects.PlatePart
    Set oPlatePart.object = pSL.InputObject(INPUT_PLATE)

    If IsSmartItemExist(m_oCatalogConnection, "Default_TemplateProcessPlate") = True Then
        pSL.Add "Default_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "FrameEqualHeight_TemplateProcessPlate") = True Then
        pSL.Add "FrameEqualHeight_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Frame_TemplateProcessPlate") = True Then
        pSL.Add "Frame_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Perpendicular_TemplateProcessPlate") = True Then
        pSL.Add "Perpendicular_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "UserDefined_TemplateProcessPlate") = True Then
        pSL.Add "UserDefined_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Even_TemplateProcessPlate") = True Then
        pSL.Add "Even_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "CenterLine_TemplateProcessPlate") = True Then
        pSL.Add "CenterLine_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "StemStern_TemplateProcessPlate") = True Then
        pSL.Add "StemStern_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "AftForward_TemplateProcessPlate") = True Then
        pSL.Add "AftForward_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "PerpendicularXY_TemplateProcessPlate") = True Then
        pSL.Add "PerpendicularXY_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "Box_TemplateProcessPlate") = True Then
        pSL.Add "Box_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "UserDefinedBox_TemplateProcessPlate") = True Then
        pSL.Add "UserDefinedBox_TemplateProcessPlate"
    End If
    If IsSmartItemExist(m_oCatalogConnection, "UserDefBoxEdges_TemplateProcessPlate") = True Then
        pSL.Add "UserDefBoxEdges_TemplateProcessPlate"
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
    
    ' Below filter criteria selects Plate parts that are detailed and NOT Planar/Flat.
    IJMfgTemplatePartInputs_GetPartFilterCriteria = IJPlatePart _
                                                    & " AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart] " _
                                                    & " AND [GSCADSTRMFGTemplateCmd.CTemplateFilter,IsAMfgParent] " _
                                                    & " AND NOT [GSCADSTRMFGTemplateCmd.CTemplateFilter,IsAFlatPlatePart] "
                                
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
