VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAngleWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrAngleWebSel.cls
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

Const m_FamilyProgid As String = ""
Const m_SelectorName As String = "Mbr_EndCuts.MbrAngleWebSel"
Const m_SelectorProgid As String = "Mbr_EndCuts.MbrAngleWebSel"
Const MODULE = "\StructDetail\Data\SmartOccurrence\Mbr_EndCuts\MbrAngleWebSel.cls"

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
    Dim lStatus As Long
    
    Dim sWeldPart As String
    Dim sWebTypeCase As String
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    sMsg = "Unknown Error"
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    On Error GoTo ErrorHandler
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    On Error Resume Next
    Set oBoundedObject = oSDO_WebCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
        
    Set oBoundingObject = oSDO_WebCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    InitEndCutConnectionData oBoundedObject, oBoundingObject, _
                             oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        sMsg = "EndCut Ports are Not valid"
        GoTo ErrorHandler
    End If
    
    ' Retreive the WebType from the Owning AssemblyConnection
    Parent_WebTypeCase oEndCutObject, sWebTypeCase
    
    ' Retreive WeldPart answer to determine "Frist" or "Second"
    sWeldPart = pSelectorLogic.Answer("WeldPart")
    
    ' Select item based on WebType case
    If InStr(LCase(sWebTypeCase), LCase("AngleWebSquareFlange")) > 0 Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            pSelectorLogic.Add "M_Web_SeamAngle_Case1"
        Else
            pSelectorLogic.Add "M_Web_SeamAngle_Case1_2"
        End If
        
    ElseIf InStr(LCase(sWebTypeCase), LCase("AngleWebBevelFlange")) > 0 Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            pSelectorLogic.Add "M_Web_SeamAngle_Case2"
        Else
            pSelectorLogic.Add "M_Web_SeamAngle_Case2_2"
        End If
    
    ElseIf InStr(LCase(sWebTypeCase), LCase("AngleWebAngleFlange")) > 0 Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            pSelectorLogic.Add "M_Web_SeamAngle_Case3"
        Else
            pSelectorLogic.Add "M_Web_SeamAngle_Case3_2"
        End If
    
    ElseIf InStr(LCase(sWebTypeCase), LCase("DistanceWebDistanceFlange")) > 0 Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            pSelectorLogic.Add "M_Web_SeamAngle_Case4"
        Else
            pSelectorLogic.Add "M_Web_SeamAngle_Case4_2"
        End If
    
    ElseIf InStr(LCase(sWebTypeCase), LCase("OffsetWebOffsetFlange")) > 0 Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            pSelectorLogic.Add "M_Web_SeamAngle_Case5"
        Else
            pSelectorLogic.Add "M_Web_SeamAngle_Case5_2"
        End If
    
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
