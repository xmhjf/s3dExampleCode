VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrByTubeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007-12 Intergraph Corporation. All rights reserved.
'
'File : MbrTubeParm.cls
'
'Author : Alligators
'
'Description :
'
'History:
'    11/Jul/2012 - svsmylav
'            TR-206472(DM-216590): For two cases, eigher (i) bounding member is a Tube/BUTube or
'            (ii) for Flange-penetrated case, parameter "FlangeMidHeightFromCentre" value is set to 0.1 mm.
'********************************************************************
Option Explicit

Const m_sClassName As String = "MbrByTubeParm"
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

    ' Parameters for Top or Bottom Parametric Snipe Cuts
    pOH.SetOutput "WebClearance"
    pOH.SetOutput "Radius"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "FlangeMidHeightFromCentre"
  
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
    
    Dim dAngle As Double
    
    Dim dCuttingDepth As Double
    sMsg = "Defining ParameterRule Outputs"
    
    Dim oPortBounding As IJPort, oPortBounded As IJPort
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    
    Dim oBounding As Object, oBounded As Object
    Set oBounding = oPortBounding.Connectable
    Set oBounded = oPortBounded.Connectable
    
    Dim bBuiltup As Boolean, oBUMember As ISPSDesignedMember
    IsFromBuiltUpMember oBounding, bBuiltup, oBUMember
    
    If bBuiltup Then
        Set oBounding = oBUMember
    End If
    
    Dim dBoundingHT, dBoundedHT, dBoundedFT  As Double
    
    If TypeOf oBounding Is ISPSMemberPartCommon Then
        Dim oBoundingMemberPart As StructDetailObjects.MemberPart
        Set oBoundingMemberPart = New MemberPart
        Set oBoundingMemberPart.object = oBounding
        dBoundingHT = oBoundingMemberPart.Height
    ElseIf TypeOf oBounding Is IJProfile Then
        Dim oBoundingProfilePart As StructDetailObjects.ProfilePart
        Set oBoundingProfilePart = New ProfilePart
        Set oBoundingProfilePart.object = oBounding
        dBoundingHT = oBoundingProfilePart.Height
    Else
        'Specail need to be handled
    End If
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBounded
        dBoundedHT = oBoundedMemberPart.Height
        dBoundedFT = oBoundedMemberPart.flangeThickness
    ElseIf TypeOf oBounded Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New ProfilePart
        Set oBoundedProfilePart.object = oBounded
        dBoundedHT = oBoundedProfilePart.Height
        dBoundedFT = oBoundedProfilePart.flangeThickness
    Else
        'Specail need to be handled
    End If
    
    pPRL.Add "WebClearance", 0.005
    
    Dim dSetbackFromBoundingMbr As Double
    dSetbackFromBoundingMbr = 0.00003 '0.03 mm
    
    pPRL.Add "Radius", dBoundingHT / 2 + dSetbackFromBoundingMbr
    
    Dim bWebPenetrates As Boolean
    bWebPenetrates = IsWebPenetrated(oPortBounding, oPortBounded)
    
    If (Not IsTubularMember(oBoundedMemberPart.object)) And bWebPenetrates Then
        pPRL.Add "FlangeMidHeightFromCentre", ((dBoundedHT / 2) - (dBoundedFT / 2))
    Else
        pPRL.Add "FlangeMidHeightFromCentre", 0.0001 '0.1 mm
    End If
    
    'define the cutting depth to completely cut the Member cross Section
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
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

