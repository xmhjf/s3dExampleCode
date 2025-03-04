VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateProcParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_ParameterRuleProgid As String = "StructMfgSelectors.PlateProcParm"
Const m_ParameterRuleName As String = "StructMfgSelectors.PlateProcParm"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PLATE
    
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler
    
  Dim PlateType As StructPlateType
  Dim PlateCurvature As StructMfgPlateCurvatureTypes
  
  'Get  Plate initialized for detailed wrappers
  Dim oDetailedPlatePart As StructDetailObjects.PlatePart
  Set oDetailedPlatePart = CreateObject("StructDetailObjects.PlatePart")
  Dim oPlate As Object
  Set oPlate = pPLH.InputObject(INPUT_PLATE)
  Set oDetailedPlatePart.object = oPlate
  PlateType = oDetailedPlatePart.PlateType
  
  'Get  Plate initialized for Manufacturing wrappers
  Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
  Set oMfgPlateWrapper = CreateObject("MfgRuleHelpers.PlatePartHlpr")
  Set oMfgPlateWrapper.object = oPlate
  PlateCurvature = oMfgPlateWrapper.CurvatureType

' ******************************************************************************
' * Using the SNU as main unfold option                                        *
' ******************************************************************************

    If PlateType = Hull Then
        pPLH.Add "PlateUnwrapAlgorithm", 1131 '"SNU"
    Else
        If PlateCurvature = StructMfgPlateCurvatureFlat Then
             pPLH.Add "PlateUnwrapAlgorithm", 11331 '"Flat"
             pPLH.Add "PlateRollLines", 11302 '"Ignore"    'Only required for developable plates
        Else
             If PlateCurvature = StructMfgPlateCurvatureDevelopable Then
                pPLH.Add "PlateUnwrapAlgorithm", 1030 ' "Developable"
                pPLH.Add "PlateRollLines", 11300 '"Fixed"
             Else
                If PlateCurvature = StructMfgPlateCurvatureKnuckled Then
                   pPLH.Add "PlateUnwrapAlgorithm", 1030 ' "Developable"
                   pPLH.Add "PlateRollLines", 11300 '"Fixed"
                Else
                   pPLH.Add "PlateUnwrapAlgorithm", 1131 '"SNU"
                   pPLH.Add "PlateRollLines", 11302 '"Ignore"    'Only required for developable plates
                End If
             End If
        End If
    End If

' * Set default values to the remaining process items
    pPLH.Add "PlateNeutralAxis", 1120 '"Calculated"
    pPLH.Add "PlateUpside", 1111 '"MoldedSide"
    pPLH.Add "PlateWeldTab", 1151 '"Ignore"
    pPLH.Add "PlateFeatureTab", 1161 '"Ignore"
    pPLH.Add "PlateBevel", 1180 '"Fixed"
    pPLH.Add "PlateBridge", 1171 '"Ignore"       'Future consideration


' ** These are specific for the Hicadec unfolding algorithms
' ** and should only have the Ignore state since they do not apply to SNU
    pPLH.Add "PlateBaseLine", 1138 '"Ignore"
    pPLH.Add "PlateSurface", 11393 '"Ignore"
    pPLH.Add "PlateWLUnfold", 11383 '"Ignore"
    pPLH.Add "PlateFSAFT", 113103 '"Ignore"
    pPLH.Add "PlateFSFORE", 113113 '"Ignore"
    pPLH.Add "PlateMX", 113123 '"Ignore"

    Set oDetailedPlatePart = Nothing
    Set oPlate = Nothing
'  Set oPlateClass = Nothing
  
  Exit Sub
ErrorHandler:
  pPLH.ReportError
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
  pIH.Definition = pSelector
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
  pOH.SetOutput "PlateBevel"
  pOH.SetOutput "PlateBaseLine"
  pOH.SetOutput "PlateWLUnfold"
  pOH.SetOutput "PlateSurface"
  pOH.SetOutput "PlateFSAFT"
  pOH.SetOutput "PlateFSFORE"
  pOH.SetOutput "PlateMX"
       
       
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

 