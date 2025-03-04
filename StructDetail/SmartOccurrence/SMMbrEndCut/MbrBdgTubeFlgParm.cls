VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrBdgTubeFlgParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : MbrBdgTubeFlgParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    27/Jul/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "MbrBdgTubeFlgParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' Parameters for end Cuts
    pOH.SetOutput "Radius"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"

    'If this is bottom flange, sets this parameter to True; Otherwise sets this to False
    pOH.SetOutput "BottomFlange"
  
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dCuttingDepth As Double
    sMsg = "Defining ParameterRule Outputs"
    
    Dim oPortBounding As IJPort
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    
    Dim oPortBounded As IJPort
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    
    Dim oBounding As Object
    Set oBounding = oPortBounding.Connectable
    
    Dim bBuiltup As Boolean, oBUMember As ISPSDesignedMember
    IsFromBuiltUpMember oBounding, bBuiltup, oBUMember
    
    If bBuiltup Then
        Set oBounding = oBUMember
    End If
    
    Dim dBoundedHT As Double
    
    If TypeOf oBounding Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBounding
        dBoundedHT = oBoundedMemberPart.Height
    ElseIf TypeOf oBounding Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New ProfilePart
        Set oBoundedProfilePart.object = oBounding
        dBoundedHT = oBoundedProfilePart.Height
    Else
        'Specail need to be handled
    End If
    
    Dim dSetbackFromBoundingMbr As Double
    
    dSetbackFromBoundingMbr = 0.00006 '0.06 mm
    
    pPRL.Add "Radius", dBoundedHT / 2 + dSetbackFromBoundingMbr
    
    'define the cutting depth to completely cut the Member cross Section
    Dim sBottomFlange As String
    GetSelectorAnswer pPRL, "BottomFlange", sBottomFlange

    Dim bIsBottomFlange As Boolean
    If sBottomFlange = "No" Then
      bIsBottomFlange = False
        pPRL.Add "BottomFlange", 0
    Else
      bIsBottomFlange = True
        pPRL.Add "BottomFlange", 1
    End If

    'define the cutting depth to completely cut the Member cross Section
    dCuttingDepth = GetAppropriateCutDepth(pPRL, bIsBottomFlange)
    pPRL.Add "CutDepth", dCuttingDepth

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation

  Dim pDFact As New DefinitionFactory
  Dim pIH As IJDInputsHelper
  Dim pOH As IJDOutputsHelper
  
  pDFact.InitAbstractParameterRule pPR
  
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
