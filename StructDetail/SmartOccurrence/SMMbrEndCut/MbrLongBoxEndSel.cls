VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrLongBoxEndSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrLongBoxEndSel.cls
'
'Author : D.A. Trent
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Web Weld case
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   available Selector Items:
'
'
'History:
'
'********************************************************************

Option Explicit

Const m_sClassName As String = "MbrLongBoxEndSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices
'


' List any graphic Inputs that the Selector has here
Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
End Sub

' List any Questions/Answers that the Selector uses here
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"

    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
End Sub

' Select the Basic Type of Member Assembly Connection here
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Set oSDOWebCut.object = pSelectorLogic.SmartOccurrence
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    Dim bIsBoundingTubular As Boolean
    bIsBoundingTubular = IsTubularMember(oSDOWebCut.Bounding)
    
    If bIsBoundingTubular Then
        ' There is only one Long Box item (symbol)
        pSelectorLogic.Add "M_Web_LongBox_Tube"
    Else
        pSelectorLogic.Add "M_Web_LongBox_01"
    End If

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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

' ********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




