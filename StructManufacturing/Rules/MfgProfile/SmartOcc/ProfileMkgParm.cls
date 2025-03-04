VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileMkgParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : ProfileMkgParm.cls
'
'  Description  : Profile Marking Parameter rule
'
'  Author       :
'
'  History      : 2004.04.21 MJV Included the correct errorhandling procedures
'
'
'**********************************************************************************************************
Option Explicit

Const MODULE = "ProfileMkgParm"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.ProfileMkgParm"
Const m_ParameterRuleName As String = "StructMfgSelectors.ProfileMkgParm"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
   Const method = "ParameterRuleInputs"
   On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PROFILE
    
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
  
    Dim oProfileClass As Object

  If TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJStiffenerPart Then
      Set oProfileClass = New StructDetailObjects.ProfilePart
  ElseIf TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJBeamPart Then
      Set oProfileClass = New StructDetailObjects.BeamPart
  End If
  Set oProfileClass.object = pPLH.InputObject(INPUT_PROFILE)
      
'  If oProfileClass.StiffenerSide = "Base" Then
      pPLH.Add "PrUserMark", 22120 ' "Apply"
      pPLH.Add "PrLocation", 22010  ' "Apply"
      pPLH.Add "PrLapConnection", 22030 ' "Apply"
      pPLH.Add "PrEndConnection", 22020 ' "Apply"
'  ElseIf oProfileClass.StiffenerSide = "Offset" Then
'      pPLH.Add "PrUserMark", 22120 ' "Apply"
'      pPLH.Add "PrLocation", 22010  ' "Apply"
'      pPLH.Add "PrLapConnection", 22030 ' "Apply"
'      pPLH.Add "PrEndConnection", 22020 '"Apply"
'  End If
  
  Set oProfileClass = Nothing
  
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
  pOH.SetOutput "PrUserMark"
  pOH.SetOutput "PrLocation"
  pOH.SetOutput "PrLapConnection"
  pOH.SetOutput "PrEndConnection"
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


 
