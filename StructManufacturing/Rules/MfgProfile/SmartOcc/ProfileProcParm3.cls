VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileProcParm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : ProfileProcParm3.cls
'
'  Description  : Profile Marking Parameter rule
'   Profile Upside
'       Web Left for angle and bulb profiles
'       Molded Side for flat bar and built-up
'   Neutral Axis            - Calculate
'   Coaming                 - Ignore
'   Clamp                   - Apply
'   Weld Compensation       - Ignore
'   Weld Tab                - Ignore
'   Mfg Built-up as Plate   - Ignore
'   Unfold Built-up         - Apply

'  Author       : StructMfg HSV Team
'
'  History      : 07/07/2010
'
'
'**********************************************************************************************************
Option Explicit

Const MODULE = "ProfileProcParm3"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.ProfileProcParm3"
Const m_ParameterRuleName As String = "StructMfgSelectors.ProfileProcParm3"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = "ParameterRuleInputs"
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PROFILE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
    Const METHOD = "ParameterRuleLogic"
    On Error GoTo ErrorHandler
    
    Dim oProfileClass As StructDetailObjects.ProfilePart
    Dim oProfilePartSupport As IJProfilePartSupport
    Set oProfilePartSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = pPLH.InputObject(INPUT_PROFILE)
    
    Set oProfileClass = New StructDetailObjects.ProfilePart
    Set oProfileClass.object = pPLH.InputObject(INPUT_PROFILE)

    ' check if the cross-section is a built-up or a flatbar
    Dim bIsBuiltUp As Boolean
    bIsBuiltUp = oProfileClass.IsCrossSectionABuiltUp
    If bIsBuiltUp = True Then
        pPLH.Add "ProfileNeutralAxis", 2123  ' "Calculated_BuiltUp"
    Else
        pPLH.Add "ProfileNeutralAxis", 2122   ' "Calculated_Standard"
    End If
     
    pPLH.Add "ProfileMfgAsPlate", 2159  ' "Ignore"
    
    Dim bstrSectionType As String
    bstrSectionType = oProfileClass.SectionType
    
    If bstrSectionType = "UA" Or bstrSectionType = "B" Then
        pPLH.Add "ProfileUpside", 2110   ' WebLeft
    ElseIf ((bIsBuiltUp = True) Or (bstrSectionType = "FB")) Then
        Dim eThicknessSide As ThicknessSide
        eThicknessSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
        If eThicknessSide = SideA Then
            pPLH.Add "ProfileUpside", 2110   ' WebLeft
        Else
            pPLH.Add "ProfileUpside", 2111   ' WebRight
        End If
    Else
        pPLH.Add "ProfileUpside", 2110   ' WebLeft
    End If
    
    ' always unfold
    pPLH.Add "ProfileUnfold", 2148              ' Apply
    
    pPLH.Add "ProfileWeldTab", 1141             ' Ignore
    pPLH.Add "ProfileWeldCompensation", 2182    ' Ignore
    pPLH.Add "ProfileCoamingUnwrap", 2152       ' Ignore
    pPLH.Add "ProfileClamp", 2160               ' Calculated
    pPLH.Add "ProfileBevel", 2171               ' Ignore
    pPLH.Add "ProfileIBL", 2132                 ' IBLInterval

    Set oProfileClass = Nothing
    Set oProfilePartSupport = Nothing
    Set oPartSupport = Nothing
  
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
  pOH.SetOutput "ProfileUpside"
  pOH.SetOutput "ProfileNeutralAxis"
  pOH.SetOutput "ProfileIBL"
  pOH.SetOutput "ProfileWeldTab"
  pOH.SetOutput "ProfileCoamingUnwrap"
  pOH.SetOutput "ProfileClamp"
  pOH.SetOutput "ProfileBevel"
  pOH.SetOutput "ProfileMfgAsPlate"
  pOH.SetOutput "ProfileUnfold"
  pOH.SetOutput "ProfileWeldCompensation"
  
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


 


