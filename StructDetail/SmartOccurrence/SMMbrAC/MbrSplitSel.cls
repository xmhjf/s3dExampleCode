VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrSplitSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrSplitSel.cls
'
'  Description :
'       Selector for Selection of Member bounded by Member:
'           End to End Colinear cases
'       Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'
'
'  Author      : Alligators
'
'  History     :
'    26/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "MbrSplitSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    ' Add your question below
    '===============================
    pQH.SetQuestion "SplitEndToEndCase", "NoAngle", "SplitEndCutTypes", "CMSetSplitEndToEndCaseAnswer"
    pQH.SetQuestion "FlipEndToEndAngle", "NoFlip", "FlipSplitEndCutAngle", "CMSetFlipEndToEndAngleAnswer"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member Assembly Connection here
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
       
    ' Case: End To End / Colinear Axis / Identical CrossSection / same Configuration
    '   Straight Web/Flange Cut or Special Offset Cuts
    '   Web and Flange have Physical Connections
    pSelectorLogic.Add "Split_MbrAC_01"

    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
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
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

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
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
End Sub
'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

'*********************************************************************************************
' Method      : CMSetSplitEndToEndCaseAnswer
' Description : Sets the default answer for the question "SplitEndToEndCase"
'
'*********************************************************************************************
Public Sub CMSetSplitEndToEndCaseAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetSplitEndToEndCaseAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("SplitEndToEndCase") = "NoAngle"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Question's Answer...", METHOD
    
End Sub

'*********************************************************************************************
' Method      : CMSetFlipEndToEndAngleAnswer
' Description : Sets the default answer for the question "FlipEndToEndAngle"
'
'*********************************************************************************************
Public Sub CMSetFlipEndToEndAngleAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetFlipEndToEndAngleAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("FlipEndToEndAngle") = "NoFlip"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Question's Answer...", METHOD
    
End Sub

