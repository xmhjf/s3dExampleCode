VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrEndFlangeSel.cls
'
'Author : D.A. Trent
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Flange case
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
Const m_SelectorName As String = "Mbr_EndCuts.MbrEndFlangeSel"
Const m_SelectorProgid As String = "Mbr_EndCuts.MbrEndFlangeSel"
Const MODULE = "\StructDetail\Data\SmartOccurrence\Mbr_EndCuts\MbrEndFlangeSel.cls"

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
    pQH.SetQuestion "BottomFlange", "No", "BooleanCol"
    
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
    
    Dim sWebCutItem As String
    Dim sFlangeTypeCase As String
    
    Dim oEndCutObject As Object
    Dim oWebCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    sMsg = "Unknown Error"
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    On Error GoTo ErrorHandler
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oEndCutObject
    
    On Error Resume Next
    Set oBoundedObject = oSDO_FlangeCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
        
    Set oBoundingObject = oSDO_FlangeCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    Set oWebCutObject = oSDO_FlangeCut.WebCut
    If (oWebCutObject Is Nothing) Then
        sMsg = "Web Cut Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    On Error GoTo ErrorHandler
    InitEndCutConnectionData oBoundedObject, oBoundingObject, _
                             oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        sMsg = "EndCut Ports are Not valid"
        GoTo ErrorHandler
    End If
    
    ' Special Case
    ' If Web Cut is Tight Weld
    ' No Flange Cut is applied
    sMsg = "Checking Web Cut SmartItem.Name"
    If TypeOf oWebCutObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oWebCutObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        sWebCutItem = oSmartItem.Name
        
        If Trim(LCase(sWebCutItem)) = LCase("M_Web_Weld_FL") Then
            Exit Sub
        ElseIf Trim(LCase(sWebCutItem)) = LCase("M_Web_Weld_FLT") Then
            Exit Sub
        ElseIf Trim(LCase(sWebCutItem)) = LCase("M_Web_Weld_FLB") Then
            Exit Sub
        End If
    
    End If
    
    ' Retreive the FlangeType from the Owning AssemblyConnection
    Parent_FlangeTypeCase oEndCutObject, sFlangeTypeCase
    
    ' Flange Left/Right is Straight
    If InStr(LCase(sFlangeTypeCase), LCase("W1W1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    
    ' Flange Left/Right is Sniped
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1S1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1W1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("W1S1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    
    ' Flange Left/Right is Coped
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1C1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("W1C1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1W1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    
    ' Flange Left/Right is Coped/Sniped
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("C1S1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
    ElseIf InStr(LCase(sFlangeTypeCase), LCase("S1C1")) > 0 Then
        FlangeType_Select pSelectorLogic, oBoundedData, oBoundingData, _
                          sFlangeTypeCase, False
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
