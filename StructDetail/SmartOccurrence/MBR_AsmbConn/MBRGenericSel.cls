VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MBRGenericSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrGenericSel.cls
'
'Author : D.A. Trent
'
'Description :
'   Selector for Selection of Muiltple Generic Member End to non-MemberPart
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'   with both Ports being: SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   The Supporting/Bounding objects are retreive from the Reference Collection
'
'   available EndCut Items:
'
'History:
'
'********************************************************************

Option Explicit

Const m_FamilyProgid As String = ""
Const m_SelectorName As String = "MBR_AsmbConn.MbrGenericSel"
Const m_SelectorProgid As String = "MBR_AsmbConn.MbrGenericSel"
Const MODULE = "\StructDetail\Data\SmartOccurrence\MBR_AsmbConn\MbrGenericSel.cls"

Implements IJDUserSymbolServices
'


' List any graphic Inputs that the Selector has here
Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
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
    
    pQH.SetQuestion "EndCutType", "W", "EndCutTypeCodeList"
    pQH.SetQuestion "WeldPart", "First"
    
    Exit Sub
    
ErrorHandler:
    pQH.ReportError sMsg, METHOD
End Sub

' Select the Basic Type of Member Assembly Connection here
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lStatus As Long
    Dim nRefArguments As Long
    
    Dim sConfig As String
    Dim sSelector As String
    
    Dim oArgObject As Object
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oBoundingEndPort As ISPSSplitAxisPort
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    Dim oHelper As GSCADSmartOccurrence.IJSmartOccurrenceHelper
    
    On Error GoTo ErrorHandler
    sMsg = "Unknown Error"
    
    Set oAppConnection = pSelectorLogic.SmartOccurrence
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        Exit Sub
    End If
    
    ' Verify both Member Parts from the AppConnection are the same
    If Not oBoundedData.MemberPart Is oBoundingData.MemberPart Then
        Exit Sub
    End If
    
    ' Verify that the Reference Collection contains valid Supporting/Bounding objects
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    nRefArguments = oEditJDArgument.GetCount
    If nRefArguments < 1 Then
        Exit Sub
    End If
    
    
    For iIndex = 1 To nRefArguments
        Set oArgObject = oEditJDArgument.GetEntityByIndex(iIndex)
        If TypeOf oArgObject Is IJPort Then
        ElseIf TypeOf oArgObject Is IJPlane Then
        ElseIf TypeOf oArgObject Is IJPoint Then
        ElseIf TypeOf oArgObject Is IJSurfaceBody Then
        Else
        End If
    Next iIndex
    
    pSelectorLogic.Add "Generic_MbrEndCut_01"
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
    pIH.Definition = pSelector
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


