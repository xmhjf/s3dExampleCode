VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAngleWebParm_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrAngleWebParm_1.cls
'
'Author : D.A. Trent
'
'Description :
'   Control Parameters of the Member Items for the MbrAngleWebSel Selector Class
'
'
'History:
'
'********************************************************************
Option Explicit

Const m_sClassName As String = "MbrAngleWebParm_1"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"


Implements IJDUserSymbolServices

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

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    ' Parameters for SeamAngle Cuts
    pOH.SetOutput "WebSeamAngle"
    pOH.SetOutput "WebSeamPointOffset"
    pOH.SetOutput "FlangeSeamAngle"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
  
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sWeldPart As String
    
    Dim dPI As Double
    Dim dCuttingDepth As Double
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sMsg = "Defining ParameterRule Outputs"
    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    Set oSmartOccurrence = pPRL.SmartOccurrence
    Set oSmartItem = oSmartOccurrence.SmartItemObject
    Set oSmartClass = oSmartItem.Parent
    Set oSymbolDefinition = oSmartClass.SelectionRuleDef

    sWeldPart = oCommonHelper.GetAnswer(oSmartOccurrence, _
                                        oSymbolDefinition, _
                                        "WeldPart")
     
    pPRL.Add "WebSeamPointOffset", 0.00001
    
    'Second Symbol contains angles measured opposite of First Symbol
    'Therefore Input values are same for both First, Second Symbols
    dPI = Atn(1#) * 4

    If Trim(LCase(sWeldPart)) = LCase("First") Then
        pPRL.Add "WebSeamAngle", dPI / 3#
        pPRL.Add "FlangeSeamAngle", dPI / 2#
    Else
        pPRL.Add "WebSeamAngle", dPI / 3#
        pPRL.Add "FlangeSeamAngle", dPI / 2#
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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

' ********************************************************************************************
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
  pIH.Definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
