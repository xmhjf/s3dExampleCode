VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateProcSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : PlateProcSel.cls
'
'  Description  : Plate process Selector rule
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

Const MODULE = "PlateProcSel"

Const m_SelectorProgid As String = "StructMfgSelectors.PlateProcSel"
Const m_SelectorName As String = "StructMfgSelectors.PlateProcSel"
Const m_FamilyProgid As String = ""

Private m_oCatalogConnection As IJDPOM

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  Const method = "SelectorsInputs"
  
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
  
'  Dim sERROR As String
'  sERROR = ""
'
  'Dim PlateType As String
  Dim ePlateType As StructPlateType
  
  ' * Get the general Plate Part information
  Dim oDetailedPlatePart As StructDetailObjects.PlatePart
  Set oDetailedPlatePart = CreateObject("StructDetailObjects.PlatePart")
  Dim oPlate As Object
  Set oPlate = pSL.InputObject(INPUT_PLATE)
  Set oDetailedPlatePart.object = oPlate
  ePlateType = oDetailedPlatePart.PlateType

  Set m_oCatalogConnection = GetCatalogResourceMgr
  
  ' * Get and set the Curvature of a Plate Part
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
        If IsSmartItemExist(m_oCatalogConnection, "FlangedPlateProcess") = True Then
            pSL.Add "FlangedPlateProcess"
        End If
        If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
            pSL.Add "DefaultPlateProcess"
        End If
   ElseIf ePlateType = Hull Then
        ' oMfgPlateWrapper will set the Curvature on the Mfg Plate
        ePlateCurvature = oMfgPlateWrapper.CurvatureType
        
        If ePlateCurvature = PLATE_CURVATURE_Flat Then
            If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                pSL.Add "FlatPlateProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                pSL.Add "DefaultPlateProcess"
            End If
        Else
            If IsSmartItemExist(m_oCatalogConnection, "ShellPlateProcess") = True Then
                pSL.Add "ShellPlateProcess"
            End If
            If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                pSL.Add "DefaultPlateProcess"
            End If
        End If
   Else
        ' oMfgPlateWrapper will set the Curvature on the Mfg Plate
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
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                      PLATE_CURVATURE_DoubleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "ShellPlateProcess") = True Then
                        pSL.Add "ShellPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_Knuckled
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_SingleCurvature_Knuckled, _
                      PLATE_CURVATURE_SingleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case Else
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "InternalPlateProcess") = True Then
                        pSL.Add "InternalPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
             End Select
          Case LBulkheadPlate
             Select Case ePlateCurvature
                 Case PLATE_CURVATURE_Flat
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                      PLATE_CURVATURE_DoubleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "ShellPlateProcess") = True Then
                        pSL.Add "ShellPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_Knuckled
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_SingleCurvature_Knuckled, _
                      PLATE_CURVATURE_SingleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case Else
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "InternalPlateProcess") = True Then
                        pSL.Add "InternalPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
             End Select
          Case TBulkheadPlate
             Select Case ePlateCurvature
                 Case PLATE_CURVATURE_Flat
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_DoubleCurvature_Knuckled, _
                      PLATE_CURVATURE_DoubleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "ShellPlateProcess") = True Then
                        pSL.Add "ShellPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_Knuckled
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case PLATE_CURVATURE_SingleCurvature_Knuckled, _
                      PLATE_CURVATURE_SingleCurvature
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                 Case Else
                    If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                        pSL.Add "DefaultPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "InternalPlateProcess") = True Then
                        pSL.Add "InternalPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                        pSL.Add "FlatPlateProcess"
                    End If
                    If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                        pSL.Add "DevelopablePlateProcess"
                    End If
             End Select
          Case Else
            If ePlateCurvature = PLATE_CURVATURE_Flat Then
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                    pSL.Add "FlatPlateProcess"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                    pSL.Add "DefaultPlateProcess"
                End If
            Else
                If IsSmartItemExist(m_oCatalogConnection, "DefaultPlateProcess") = True Then
                    pSL.Add "DefaultPlateProcess"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "InternalPlateProcess") = True Then
                    pSL.Add "InternalPlateProcess"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "ShellPlateProcess") = True Then
                    pSL.Add "ShellPlateProcess"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "FlatPlateProcess") = True Then
                    pSL.Add "FlatPlateProcess"
                End If
                If IsSmartItemExist(m_oCatalogConnection, "DevelopablePlateProcess") = True Then
                    pSL.Add "DevelopablePlateProcess"
                End If
            End If
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


 
