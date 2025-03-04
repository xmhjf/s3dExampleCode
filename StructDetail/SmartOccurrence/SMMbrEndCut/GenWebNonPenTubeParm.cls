VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebNonPenTubeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenWebNonPenTubeParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    09/OCT/2012 - Created
'*********************************************************************************************

Const m_sClassName As String = "GenWebNonPenTubeParm"
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
    pOH.SetOutput "TopSnipeAngle"
    pOH.SetOutput "BottomSnipeAngle"
    pOH.SetOutput "ExtrusionLength"
    pOH.SetOutput "Nose"
    pOH.SetOutput "Clearance"
      
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
    
    sMsg = "Defining ParameterRule Outputs"
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    
    
    Dim oBoundedObject As Object
    Dim oSectionAttrbs As IJDAttributes
    Dim oBoundedMemberPart As ISPSMemberPartPrismatic
    Set oBoundedObject = oPortBounded.Connectable
    Set oBoundedMemberPart = oBoundedObject
    Set oSectionAttrbs = oBoundedMemberPart.CrossSection.definition
    Dim dDiameter As Double
    dDiameter = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value

    If IsTubularMember(oBoundedObject) Then
        'For all tubular members the Cutting Depth should be set complete Depth
        'since it doesn't have any Flanges irrespective of Web/Flange Penetrated cases
        'Only for Flanged Cross Section we give optimum CutDepth such that
        'Flange CutDepth will do remaining Cut for Flange Penetrated cases but this check
        'is not required for Tubular Cross Section(e.g PIPE, HSSC, CS, P, R etc...)
        'Hence setting the full Cut Depth
        dCuttingDepth = EndCut_GetCutDepth(pPRL)
        
    End If
    'As per the Design ,Snipe Angles for top and Bottom are given as 45 degrees
    'Nose as 0.1mm
    'ExtrusionLength as 1.5 * dia of Tube
    Dim dTopSnipeAngle As Double
    Dim dBottomSnipeAngle As Double
    Dim dExtrusionLength As Double
    Dim dNose As Double
    Dim dPI As Double
    Dim dClearance As Double
    dPI = Atn(1) * 4#
    dTopSnipeAngle = dPI / 4#
    dBottomSnipeAngle = dPI / 4#
    dExtrusionLength = 1.5 * dDiameter
    dNose = 0.0001 '0.1mm
    dClearance = 0.00001 '0.01mm
    
        
    pPRL.Add "CutDepth", dCuttingDepth
    pPRL.Add "TopSnipeAngle", dTopSnipeAngle
    pPRL.Add "BottomSnipeAngle", dBottomSnipeAngle
    pPRL.Add "ExtrusionLength", dExtrusionLength
    pPRL.Add "Nose", dNose
    pPRL.Add "Clearance", dClearance
    
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




