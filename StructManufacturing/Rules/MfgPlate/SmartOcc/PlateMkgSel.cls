VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateMkgSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : PlateMkgSel.cls
'
'  Description  : Plate Marking Selector rule
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

Const MODULE = "PlateMkgSel"

Const m_SelectorProgid As String = "StructMfgSelectors.PlateMkgSel"
Const m_SelectorName As String = "StructMfgSelectors.PlateMkgSel"
Const m_FamilyProgid As String = ""
Private m_oCatalogConnection As IJDPOM

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  Const method = "SelectorInput"
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PLATE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
  'pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  Const method = "SelectorQuestions"
  On Error GoTo ErrorHandler
  
  'No questions defined
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
  'pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
  Const method = "SelectorLogic"
  On Error GoTo ErrorHandler
  
  Dim sERROR As String
  sERROR = ""
  
  Dim ePlateType As StructPlateType
  
  Set m_oCatalogConnection = GetCatalogResourceMgr
  
  ' * Get the general Plate Part information
  Dim oDetailedPlatePart As StructDetailObjects.PlatePart
  Set oDetailedPlatePart = CreateObject("StructDetailObjects.PlatePart")
  Dim oPlate As Object
  Set oPlate = pSL.InputObject(INPUT_PLATE)
  Set oDetailedPlatePart.object = oPlate
  ePlateType = oDetailedPlatePart.PlateType
  
  ' * Get the Curvature of a Plate Part
  Dim ePlateCurvature As PlateCurvature
  Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
  Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
  Set oMfgPlateWrapper.object = oPlate

  'Determine if we deal with a flanged bracket
  Dim bIsFlangedBracket As Boolean
  bIsFlangedBracket = False

  bIsFlangedBracket = IsFlangedBracket(oPlate)
 
 ' ********************* Selection code - TODO  ****************
  
  If bIsFlangedBracket = True Then
        If IsSmartItemExist(m_oCatalogConnection, "FlangedPlateMarking") = True Then
            pSL.Add "FlangedPlateMarking"
        End If
        If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateMarking") = True Then
            pSL.Add "DefaultPlateMarking"
        End If
  ElseIf ePlateType = Hull Then
        ' Calling CurvatureType function takes long time so do not call it if not really necessary.
        ePlateCurvature = oMfgPlateWrapper.CurvatureType
  
        If ePlateCurvature = PLATE_CURVATURE_Flat Then
            If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                pSL.Add "FlatPlateMarking"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                pSL.Add "InternalPlateMarking"
            End If
        Else
            If IsSmartItemExist(m_oCatalogConnection, "ShellPlateMarking") = True Then
                pSL.Add "ShellPlateMarking"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                pSL.Add "PlateMarkingDefault"
            End If
        End If
   Else
       ' Calling CurvatureType function takes long time so do not call it if not really necessary.
        ePlateCurvature = oMfgPlateWrapper.CurvatureType

        ' If the input PlatePart is in 'Light' Representation, No Item Should be Selected.
        ' This will delete the Smart Occurrence.
                
        Dim lNestingStage As Long
        lNestingStage = GetProjectRootNestingStage()
        
        If lNestingStage = PRODUCTION_STAGE Then
            If TypeOf oPlate Is IJPartGeometryState Then
                Dim oPartGeometryState As IJPartGeometryState
                Set oPartGeometryState = oPlate
        
                If oPartGeometryState.PartGeometryState = LightPart Then
                    Set oPartGeometryState = Nothing
                    Set oDetailedPlatePart = Nothing
                    Set oPlate = Nothing
                    Exit Sub
                End If
                Set oPartGeometryState = Nothing
            End If
        End If

        Select Case ePlateType
         Case DeckPlate
            Select Case ePlateCurvature
             Case PLATE_CURVATURE_Flat
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                    pSL.Add "FlatPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                  PLATE_CURVATURE_DoubleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "ShellPlateMarking") = True Then
                    pSL.Add "ShellPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_SingleCurvature_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_SingleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case Else
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
            End Select
            
         Case LBulkheadPlate
            Select Case ePlateCurvature
             Case PLATE_CURVATURE_Flat
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                    pSL.Add "FlatPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                  PLATE_CURVATURE_DoubleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "ShellPlateMarking") = True Then
                    pSL.Add "ShellPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_SingleCurvature_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_SingleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case Else
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
            End Select
         Case TBulkheadPlate
            Select Case ePlateCurvature
             Case PLATE_CURVATURE_Flat
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                    pSL.Add "FlatPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                  PLATE_CURVATURE_DoubleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "ShellPlateMarking") = True Then
                    pSL.Add "ShellPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case PLATE_CURVATURE_SingleCurvature_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_SingleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
             Case Else
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                    pSL.Add "InternalPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
            End Select
         Case CollarPlate
            If IsSmartItemExist(m_oCatalogConnection, "FlatPlateLeastMarking") = True Then
                 pSL.Add "FlatPlateLeastMarking"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                pSL.Add "FlatPlateMarking"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "InternalPlateMarking") = True Then
                pSL.Add "InternalPlateMarking"
            End If
         Case Else
            Select Case ePlateCurvature
             Case PLATE_CURVATURE_Flat
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateMarking") = True Then
                    pSL.Add "FlatPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                  PLATE_CURVATURE_DoubleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "ShellPlateMarking") = True Then
                    pSL.Add "ShellPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_SingleCurvature_Knuckled
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case PLATE_CURVATURE_SingleCurvature
                If IsSmartItemExist(m_oCatalogConnection, "KnuckledPlateMarking") = True Then
                    pSL.Add "KnuckledPlateMarking"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
             Case Else
                If IsSmartItemExist(m_oCatalogConnection, "PlateMarkingDefault") = True Then
                    pSL.Add "PlateMarkingDefault"
                End If
            End Select
           
        End Select
        Set oMfgPlateWrapper = Nothing
    End If
  ' *********************************************************
  
  Set oDetailedPlatePart = Nothing
  Set oPlate = Nothing
  
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
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


 
