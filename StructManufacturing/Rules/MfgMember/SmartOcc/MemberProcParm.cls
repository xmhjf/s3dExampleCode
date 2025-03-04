VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberProcParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : MemberProcParm.cls
'
'  Description  : Member Marking Parameter rule
'
'  Author       :
'
'  History      : Created   Feb 8, 2008
'
'
'**********************************************************************************************************
Option Explicit

Const MODULE = "MemberProcParm"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.MemberProcParm"
Const m_ParameterRuleName As String = "StructMfgSelectors.MemberProcParm"
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
    
    Dim oProfilePartSupport As IJProfilePartSupport
    Set oProfilePartSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = pPLH.InputObject(INPUT_PROFILE)
    
    Dim oEntityHelper As New MfgEntityHelper
    Dim sProfileSectionTypeName As String
    sProfileSectionTypeName = oEntityHelper.GetProfileSectionTypeName(oProfilePartSupport)
    If sProfileSectionTypeName = "RectangleTube" Or sProfileSectionTypeName = "RoundTube" Or _
        sProfileSectionTypeName = "RoundBar" Or sProfileSectionTypeName = "TwoL" Or _
        sProfileSectionTypeName = "LBar" Then
        If oEntityHelper.CheckIfPartExistsInSectionXML(pPLH.InputObject(INPUT_PROFILE)) Then
            If oEntityHelper.CheckIfCodeListEntryExists(pPLH.InputObject(INPUT_PROFILE), "SMMemberProcessUpside", 30) Then
                pPLH.Add "ProfileUpside", 30 'WebLeftFromXML
            End If
        End If
    End If
    
    pPLH.Add "ProfileUpside", 10   ' WebLeft
    pPLH.Add "ProfileNeutralAxis", 20    ' Fixed
    pPLH.Add "ProfileIBL", 10           '  NumberIs1
    pPLH.Add "ProfileWeldTab", 1      ' Ignore
    pPLH.Add "ProfileCoamingUnwrap", 5  ' Fixed
    pPLH.Add "ProfileClamp", 5   '  Fixed
    pPLH.Add "ProfileBevel", 10 ' "Ignore"
    
    Set oEntityHelper = Nothing
    Set oPartSupport = Nothing
    Set oProfilePartSupport = Nothing
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
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)
  pOH.SetOutput "ProfileUpside"
  pOH.SetOutput "ProfileNeutralAxis"
  pOH.SetOutput "ProfileIBL"
  pOH.SetOutput "ProfileWeldTab"
  pOH.SetOutput "ProfileCoamingUnwrap"
  pOH.SetOutput "ProfileClamp"
  pOH.SetOutput "ProfileBevel"
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


 
