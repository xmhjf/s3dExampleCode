VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateProcParm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : PlateProcParm2.cls
'
'  Description  : Plate process Parameter rule
'
'  Author       :
'
'  History      :
'
'
'**********************************************************************************************************
Option Explicit

Const MODULE = "PlateProcParm3"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.PlateProcParm3"
Const m_ParameterRuleName As String = "StructMfgSelectors.PlateProcParm3"
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
    
  Dim ePlateType As StructPlateType
  
  'Get  Plate initialized for detailed wrappers
  Dim oDetailedPlatePart As StructDetailObjects.PlatePart
  Set oDetailedPlatePart = CreateObject("StructDetailObjects.PlatePart")
  Dim oPlate As Object
  Set oPlate = pPLH.InputObject(INPUT_PLATE)
  Set oDetailedPlatePart.object = oPlate
  ePlateType = oDetailedPlatePart.plateType
  
  'Get  Plate initialized for Manufacturing wrappers
  Dim ePlateCurvature As PlateCurvature
  Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
  Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
  Set oMfgPlateWrapper.object = oPlate
  ' Calling CurvatureType function takes long time so do not call it if not really necessary.
  ePlateCurvature = oMfgPlateWrapper.CurvatureType
  
  'Determine if we deal with a flanged bracket
  Dim bIsFlangedBracket As Boolean
  bIsFlangedBracket = False

  bIsFlangedBracket = IsFlangedBracket(oPlate)
  
' ******************************************************************************
' * Using the SNU as main unfold option                                        *
' ******************************************************************************
  If bIsFlangedBracket = True Then
        pPLH.Add "PlateUnwrapAlgorithm", 1130 ' "Developable"
        pPLH.Add "PlateMargin", 1160 '"Calculated"
  ElseIf ePlateType = Hull Then
        If ePlateCurvature = PLATE_CURVATURE_Flat Then
            pPLH.Add "PlateUnwrapAlgorithm", 11331 '"Flat"
            pPLH.Add "PlateRollLines", 11302 '"Ignore"
        'CR-CP-248064  Enhance the logic to select appropriate unfold alogorithm to a given plate shape
        'if Curvature is single curvature, select Developable Algorithm
        ElseIf ePlateCurvature = PLATE_CURVATURE_Knuckled Or ePlateCurvature = PLATE_CURVATURE_SingleCurvature_Knuckled _
                      Or ePlateCurvature = PLATE_CURVATURE_SingleCurvature Then
            pPLH.Add "PlateUnwrapAlgorithm", 1130 '"Developable"
        Else
            pPLH.Add "PlateUnwrapAlgorithm", 1131 '"SNU"
        End If
        pPLH.Add "PlateMargin", 1162 '"Ignore"
   Else

        Select Case ePlateCurvature
            Case PLATE_CURVATURE_Flat
                    pPLH.Add "PlateUnwrapAlgorithm", 11331 '"Flat"
                    pPLH.Add "PlateRollLines", 11302 '"Ignore"    'Only required for developable plates
            
            Case PLATE_CURVATURE_SingleCurvature_Knuckled, _
                 PLATE_CURVATURE_SingleCurvature
                    pPLH.Add "PlateUnwrapAlgorithm", 1130 ' "Developable"
                    pPLH.Add "PlateRollLines", 11300 '"Fixed"
             
            Case PLATE_CURVATURE_Knuckled
                    pPLH.Add "PlateUnwrapAlgorithm", 1130 ' "Developable"
                    pPLH.Add "PlateRollLines", 11300 '"Fixed"
                
            Case Else
                    pPLH.Add "PlateUnwrapAlgorithm", 11341 '"TwoDirUnfold"
                    pPLH.Add "PlateRollLines", 11302 '"Ignore"    'Only required for developable plates
        End Select
        pPLH.Add "PlateMargin", 1162 '"Ignore"
        Set oMfgPlateWrapper = Nothing
    End If

' * Set default values to the remaining process items
    pPLH.Add "PlateNeutralAxis", 1120 '"Calculated"
    pPLH.Add "PlateUpside", 1111 '"MoldedSide"
    pPLH.Add "PlateWeldTab", 1152 '"Ignore"
    pPLH.Add "PlateFeatureTab", 1162 '"Ignore"
    pPLH.Add "PlateKnuckleTab", 1192
    pPLH.Add "PlateBevel", 1180 '"Fixed"
    pPLH.Add "PlateShrinkage", 113211 '"Ignore"

    Set oDetailedPlatePart = Nothing
    Set oPlate = Nothing
'  Set oPlateClass = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pPLH.ReportError
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
  pOH.SetOutput "PlateUpside"
  pOH.SetOutput "PlateNeutralAxis"
  pOH.SetOutput "PlateUnwrapAlgorithm"
  pOH.SetOutput "PlateWeldTab"
  pOH.SetOutput "PlateFeatureTab"
  pOH.SetOutput "PlateKnuckleTab"
  pOH.SetOutput "PlateBevel"
  pOH.SetOutput "PlateMargin"
  pOH.SetOutput "PlateShrinkage"
       
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

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


