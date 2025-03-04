VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlgNonPenGussetParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFlgNonPenGussetParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    01/November/2012 - Created to support EndCut when a standard member i8s bounded to plate's lateral edge
'    20/November/2012 - Modified the parameter Names As per the Requirements
'*********************************************************************************************

Const m_sClassName As String = "GenFlgNonPenGussetParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
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
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "WeldLength"
    pOH.SetOutput "FlangeClearance"
    pOH.SetOutput "BottomFlange"
    pOH.SetOutput "Nose"
    pOH.SetOutput "FlgRightToLeft"
    pOH.SetOutput "WebLeftToFlgLeft"
    pOH.SetOutput "LeftSnipeAngle"
    pOH.SetOutput "RightSnipeAngle"
      
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dCuttingDepth As Double
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Dim bIsBottomFlange As Boolean
    Dim dWeldLength As Double
    Dim dFlangeClearance As Double
    Dim dWebLeftToFlgLeft As Double
    Dim dFlgRightToLeft As Double
    Dim WebThk As Double
    Dim dLeftSnipeAngle As Double
    Dim dRightSnipeAngle As Double
   
    Dim dNose As Double
    Dim oBoundedObject As Object
    Dim oSectionAttrbs As IJDAttributes
    Dim oBoundedMemberPart As ISPSMemberPartPrismatic

    Dim sBottomFlange As String
    sMsg = "Defining ParameterRule Outputs"
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    Set oBoundedObject = oPortBounded.Connectable
    Set oBoundedMemberPart = oBoundedObject
    
    'getting the depth of the section for weld length
    Set oSectionAttrbs = oBoundedMemberPart.CrossSection.definition
    Dim dDepth As Double
    dDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    '**************************
    dCuttingDepth = GetAppropriateCutDepth(pPRL, bIsBottomFlange)
    pPRL.Add "CutDepth", dCuttingDepth
    '*********************
    GetSelectorAnswer pPRL, "BottomFlange", sBottomFlange

    If sBottomFlange = "No" Then
        bIsBottomFlange = False
        pPRL.Add "BottomFlange", 0
    Else
        bIsBottomFlange = True
        pPRL.Add "BottomFlange", 1
    End If
    
    'Nose and WeldLength parameter values are given as per design Requirements.
    dNose = 0.015  '15 mm
    dWeldLength = 1.5 * dDepth
    dFlangeClearance = 0.0002 '0.2mm is the default value
    Dim dPI As Double
    dPI = Atn(1) * 4#
    dLeftSnipeAngle = dPI / 4#
    dRightSnipeAngle = dPI / 4#
    
    pPRL.Add "WeldLength", dWeldLength
    pPRL.Add "FlangeClearance", dFlangeClearance
    pPRL.Add "Nose", dNose
    pPRL.Add "LeftSnipeAngle", dLeftSnipeAngle
    pPRL.Add "RightSnipeAngle", dRightSnipeAngle
    'getting the values of distances between flange left to right and webLeft to flange left
    Dim oBndMember As StructDetailObjects.MemberPart
    Set oBndMember = New MemberPart
    Set oBndMember.object = oBoundedObject
    WebThk = oBndMember.WebThickness
    dFlgRightToLeft = dDepth
    dWebLeftToFlgLeft = (dFlgRightToLeft - WebThk) / 2
    pPRL.Add "FlgRightToLeft", dFlgRightToLeft
    pPRL.Add "WebLeftToFlgLeft", dWebLeftToFlgLeft
    
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
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
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
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
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




