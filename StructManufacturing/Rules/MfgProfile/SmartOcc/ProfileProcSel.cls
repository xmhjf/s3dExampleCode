VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileProcSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : ProfileProcSel.cls
'
'  Description  : Profile Process Selector rule
'
'   Selector:
'   ---------
'   It selects the list of avaible Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parametrize the selection
'
'  Author       :
'
'  History      : 2004.04.21 MJV Included the correct errorhandling procedures
'
'
'**********************************************************************************************************

Option Explicit

Const MODULE = "ProfileProcSel"

Const m_SelectorProgid As String = "StructMfgSelectors.ProfileProcSel"
Const m_SelectorName As String = "StructMfgSelectors.ProfileProcSel"
Const m_FamilyProgid As String = ""
Private m_oCatalogConnection As IJDPOM

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = "SelectorInputs"
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PROFILE
  
  Exit Sub
ErrorHandler:
      Err.Raise LogError(Err, MODULE, METHOD).Number
  '   pIH.ReportError
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
    
    ' If the input ProfilePart is in 'Light' Representation, No Item Should be Selected.
    ' This will delete the Smart Occurrence.
    
    Dim oObject As Object
    Set oObject = pSL.InputObject(INPUT_PROFILE)
    
    Set m_oCatalogConnection = GetCatalogResourceMgr
    
    Dim lNestingStage As Long
    lNestingStage = GetProjectRootNestingStage()
    
    If lNestingStage = PRODUCTION_STAGE Then
        If TypeOf oObject Is IJPartGeometryState Then
            Dim oPartGeometryState As IJPartGeometryState
            Set oPartGeometryState = oObject
    
            If oPartGeometryState.PartGeometryState = LightPart Then
                Set oPartGeometryState = Nothing
                Exit Sub
            End If
            Set oPartGeometryState = Nothing
        End If
    End If

    Dim Profile_Curvature As ProfileCurvature
     
  ' * Get the Curvature of a Profile Part
    Profile_Curvature = EvaluateProfileCurvature(oObject)
    
    Select Case Profile_Curvature
        Case PROFILE_CURVATURE_Straight
            If IsSmartItemExist(m_oCatalogConnection, "LinearProfileProcess") = True Then
                pSL.Add "LinearProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultProfileProcess") = True Then
                pSL.Add "DefaultProfileProcess"
            End If
        Case PROFILE_CURVATURE_CurvedAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongFlange
            If IsSmartItemExist(m_oCatalogConnection, "SingleCurveProfileProcess") = True Then
                pSL.Add "SingleCurveProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultProfileProcess") = True Then
                pSL.Add "DefaultProfileProcess"
            End If
        Case (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) 'Both Curved
            If IsSmartItemExist(m_oCatalogConnection, "DoubleCurveProfileProcess") = True Then
                pSL.Add "DoubleCurveProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultProfileProcess") = True Then
                pSL.Add "DefaultProfileProcess"
            End If
        Case PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongWeb Or (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_CurvedAlongFlange Or (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or _
                            (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_BendKnuckleAlongFlange, _
             PROFILE_CURVATURE_BendKnuckleAlongWeb
            If IsSmartItemExist(m_oCatalogConnection, "KnuckledProfileProcess") = True Then
                pSL.Add "KnuckledProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultProfileProcess") = True Then
                pSL.Add "DefaultProfileProcess"
            End If
        Case Else
            If IsSmartItemExist(m_oCatalogConnection, "DefaultProfileProcess") = True Then
                pSL.Add "DefaultProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "LinearProfileProcess") = True Then
                pSL.Add "LinearProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "SingleCurveProfileProcess") = True Then
                pSL.Add "SingleCurveProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DoubleCurveProfileProcess") = True Then
                pSL.Add "DoubleCurveProfileProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "KnuckledProfileProcess") = True Then
                pSL.Add "KnuckledProfileProcess"
            End If
    End Select
  
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
' pSL.ReportError sERROR
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
  pIH.Definition = pSelector
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


 
