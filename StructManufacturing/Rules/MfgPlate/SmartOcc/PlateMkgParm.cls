VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateMkgParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : PlateMkgParm.cls
'
'  Description  : Plate Marking Parameter rule
'
'  Author       :
'
'  History      : 2004.04.21 MJV Included the correct errorhandling procedures
'
'
'**********************************************************************************************************

Option Explicit

Const MODULE = "PlateMkgParm"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.PlateMkgParm"
Const m_ParameterRuleName As String = "StructMfgSelectors.PlateMkgParm"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  Const method = "ParameterRuleInputs"
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PLATE
    
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, method).Number
    'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
   Const method = "ParameterRuleLogic"
   On Error GoTo ErrorHandler
    
' ************************************************************
' * Currently no Parameter rule logic identified for marking *
' ************************************************************

  ' * Get the general Plate Part information
  Dim oDetailedPlatePart As StructDetailObjects.PlatePart
  Set oDetailedPlatePart = CreateObject("StructDetailObjects.PlatePart")
  Dim oPlate As Object
  Set oPlate = pPLH.InputObject(INPUT_PLATE)
  Set oDetailedPlatePart.object = oPlate
  
  pPLH.Add "PlPlateLocation", pPLH.SmartItemAttribute("PlPlateLocation")
  pPLH.Add "PlProfileLocation", pPLH.SmartItemAttribute("PlProfileLocation")
  pPLH.Add "PlCollarLocation", pPLH.SmartItemAttribute("PlCollarLocation")
       
  Set oDetailedPlatePart = Nothing
  Set oPlate = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
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
  pIH.Definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)
  pOH.SetOutput "PlPlateLocation", imsARGUMENT_IS_BSTR
  pOH.SetOutput "PlProfileLocation", imsARGUMENT_IS_BSTR
  pOH.SetOutput "PlCollarLocation", imsARGUMENT_IS_BSTR
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


 
