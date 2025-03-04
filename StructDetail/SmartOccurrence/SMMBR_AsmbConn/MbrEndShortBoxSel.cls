VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndShortBoxSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrEndShortBoxSel.cls
'
'Author : D.A. Trent
'
'Description :
'   Selector for Selection of Member bounded by Member:
'       End to End Non-Colinear cases
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   available Selector Items:
'       ShortBox_W1W1_W1W1 :
'           Weld_Web_Top/Weld_Web_Bottom/Weld_Top_Flange/Weld_Bottom_Flange
'       ShortBox_W1S1_W1S1 :
'           Weld_Web_Top/Snipe_Web_Bottom/Weld_Top_Flange/Snipe_Bottom_Flange
'       ShortBox_S1W1_S1W1 :
'           Snipe_Web_Top/Weld_Web_Bottom/Snipe_Top_Flange/Weld_Bottom_Flange
'       ShortBox_S1S1_S1S1 :
'           Snipe_Web_Top/Snip_Web_Bottom/Snipe_Top_Flange/Snipe_Bottom_Flange
'       ShortBox_W1W1_X1X1 :
'           Weld_Web_Top/Weld_Web_Bottom/None_Top_Flange/None_Bottom_Flange
'
'       ShortBox_W1C1_W1S1 :
'           Weld_Web_Top/Cope_Web_Bottom/Weld_Top_Flange/Snipe_Bottom_Flange
'       ShortBox_C1W1_S1W1 :
'           Cope_Web_Top/Weld_Web_Bottom/Snipe_Top_Flange/Weld_Bottom_Flange
'       ShortBox_C1C1_S1S1 :
'           Cope_Web_Top/Cope_Web_Bottom/Snipe_Top_Flange/Snipe_Bottom_Flange
'
'       ShortBox_C1S1_S1S1 :
'           Cope_Web_Top/Snipe_Web_Bottom/Snipe_Top_Flange/Snipe_Bottom_Flange
'       ShortBox_S1C1_S1S1 :
'           Snipe_Web_Top/Cope_Web_Bottom/Snipe_Top_Flange/Snipe_Bottom_Flange
'
'       ShortBox_W1W1_C1C1 :
'           Weld_Web_Top/Weld_Web_Bottom/Cope_Top_Flange/Cope_Bottom_Flange
'       ShortBox_W1W1_C1W1 :
'           Weld_Web_Top/Weld_Web_Bottom/Cope_Top_Flange/Weld_Bottom_Flange
'       ShortBox_W1W1_W1C1 :
'           Weld_Web_Top/Weld_Web_Bottom/Wled_Top_Flange/Cope_Bottom_Flange
'
'
'History:
'
'********************************************************************

Option Explicit

Const m_sClassName As String = "MbrEndShortBoxSel"
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
    
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList"
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
    
    Dim sConfig As String
    Dim sSelector As String
    Dim sIdealizedBoundary As String
    
    Dim bColinear As Boolean
    Dim bEndToEnd As Boolean
    Dim bIdentical As Boolean
    Dim bRightAngle As Boolean
    
    Dim oAppConnection As IJAppConnection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    On Error GoTo ErrorHandler
    sMsg = "Unknown Error"
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    Set oAppConnection = pSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        Exit Sub
    End If
    
    ' Assembly Connection is End To End Type LongBox case:
    '   Member Axis are Not Collinear (ideally should be perpendicular)
    '
    ' Verify this is a valid Case for this Selector
    ' Check if Assembly Connection is End To End Type
    CheckEndToEndConnection oBoundedData, oBoundingData, bEndToEnd, bColinear, bRightAngle
    If Not bEndToEnd Then
        Exit Sub
    ElseIf bColinear Then
        Exit Sub
    End If
    
    ' Check if either Bounding or Bounded Cross Section Type is Tubular
    sIdealizedBoundary = eIdealized_Unk
    CheckIdealizedBoundary oBoundedData, oBoundingData, sIdealizedBoundary
    If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_BoundingTube) Then
        pSelectorLogic.Add "ShortBox_W1W1_W1W1"
        Exit Sub
    Else
        CheckIdealizedBoundary oBoundingData, oBoundedData, sIdealizedBoundary
        If Trim(LCase(sIdealizedBoundary)) = LCase(eIdealized_BoundingTube) Then
            pSelectorLogic.Add "ShortBox_W1W1_W1W1"
            Exit Sub
        End If
    End If
    
    ' Determine Configuration between Bounded and Bounding Member
    '   sConfig = "Top_Top"     : Bounded/Bounding Top Flanges are in same general direction
    '   sConfig = "Bottom_Top"  : Bounded/Bounding Top Flanges are in the opposite general direction
    '   sConfig = "Left_Top"    : Bounded Web_Left/Bounding Top Flange are in same general direction
    '   sConfig = "Right_Top"   : Bounded Web_Right/Bounding Top Flange are in same general direction
    EndCut_WebFlangeConfig oBoundedData, oBoundingData, sConfig
    
    ' Simple Weld (Straight) cut on both Web and Flange
    pSelectorLogic.Add "ShortBox_W1W1_W1W1"
    
    If Trim(LCase(sConfig)) = LCase("Top_Top") Or _
       Trim(LCase(sConfig)) = LCase("Bottom_Top") Then
        ' Simple Snipe (Angled) cut on Web and Flange
        pSelectorLogic.Add "ShortBox_W1S1_W1S1"
        pSelectorLogic.Add "ShortBox_S1W1_S1W1"
        pSelectorLogic.Add "ShortBox_S1S1_S1S1"
        
        ' Simple Cope (Curved) cut on Web and Sniped (Angled) cut on Flange
        pSelectorLogic.Add "ShortBox_W1C1_W1S1"
        pSelectorLogic.Add "ShortBox_C1W1_S1W1"
        pSelectorLogic.Add "ShortBox_C1C1_S1S1"
        
        ' Cope (Curved) and Snipe (Angled) cut on Web and Sniped (Angled) cut on Flange
        pSelectorLogic.Add "ShortBox_C1S1_S1S1"
        pSelectorLogic.Add "ShortBox_S1C1_S1S1"

    ElseIf Trim(LCase(sConfig)) = LCase("Left_Top") Or _
           Trim(LCase(sConfig)) = LCase("Right_Top") Then
        ' Weld (Straight) cut on Web and Coped (Curved) cut on Flange
        pSelectorLogic.Add "ShortBox_W1W1_C1C1"
        pSelectorLogic.Add "ShortBox_W1W1_C1W1"
        pSelectorLogic.Add "ShortBox_W1W1_W1C1"
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



