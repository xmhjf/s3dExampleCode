VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrEndCutSel.cls
'
'Author : D.A. Trent
'
'Description :
'   Root Default Selector for selecting EndCut Assembly Connection with SPS Members as inputs
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   available EndCut Selectors:
'       MBR_AlongAxisSel
'       MBR_EndSplitSel
'           MBR_EndSeamAngleSel
'       MBR_EndMiterSel
'           MBR_EndLongBoxSel
'           MBR_EndShortBoxSel
'
'History:
'
'********************************************************************

Option Explicit

Const m_sClassName As String = "MbrEndCutSel"
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
    Dim sCurrentChoice As String
    Dim sIdealizedBoundary As String
    
    Dim bCopeCut As Boolean
    Dim bColinear As Boolean
    Dim bEndToEnd As Boolean
    Dim bIdentical As Boolean
    Dim bRightAngle As Boolean
    Dim bAllowBearing As Boolean
    
    Dim bTopFlange As Boolean
    Dim bBottomFlange As Boolean
    
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    Dim oHelper As GSCADSmartOccurrence.IJSmartOccurrenceHelper
    
    On Error GoTo ErrorHandler
    sMsg = "Unknown Error"
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    Set oAppConnection = pSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        Exit Sub
        'If not a supported Member By Member Case
        'Do Not select a current Item
        'GoTo UseSPSACMacros
    End If
    
    ' if Supported/Bounded Member Part and Supporting/Bounding MemberPart are same
    ' then this is Muiltple General/Generic Supporting/Bounding objects
    '   Plate Part Edge/Face Ports, Ref. Planes, Points, IJSurfaceBody
    If oBoundedData.MemberPart Is oBoundingData.MemberPart Then
        ' Case: Muiltple Generic Member End to non-MemberPart
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        If oEditJDArgument.GetCount > 0 Then
            sSelector = "MBR_GenericSel"
            pSelectorLogic.Add sSelector
        Else
            ' Case: Free Member End Cut
            sSelector = "MBR_FreeEndSel"
            pSelectorLogic.Add sSelector
        End If
        
        Exit Sub
    End If
    
    ' Check if Assembly Connection is End To End Type
    sSelector = ""
    CheckEndToEndConnection oBoundedData, oBoundingData, bEndToEnd, bColinear, bRightAngle
    If bEndToEnd Then
        ' Assembly Connection is End To End Type
        ' Check if Axis are Colinear
        If Not oBoundingData.MemberPart.IsPrismatic Then
            ' Bounding Member is NOT a MemberPartPrismatic
            ' Initial implementation will not support:
            ' Member Part Prismatic bounded by Designed Member Part
            Exit Sub
            
        ElseIf bColinear Then
            ' Case: End To End / Colinear Axis
            sSelector = "MBR_SplitSel"
            pSelectorLogic.Add "MBR_SplitSel"
        Else
            ' Case: End To End / Non-Colinear Axis
            sSelector = "MBR_MiterSel"
            pSelectorLogic.Add sSelector
        End If
    Else
        ' Default Assembly Connection is Along Axis Type
        sSelector = "MBR_AxisSel"
        pSelectorLogic.Add sSelector
    End If
    
    ' Add the Item to the current Selection
    If Len(Trim(sSelector)) < 1 Then
        sMsg = "Select: SPS MemberByMemberSel"
        pSelectorLogic.Add "SPS_MemberByMember"
    End If
    
    Exit Sub
    
UseSPSACMacros:
    ' Current Configuration is not valid for Member By Member End Cuts
    ' Default to the Common Struct Assembly Connection
    sMsg = "Defaulting to SPS MemberByMemberSel"
    pSelectorLogic.Add "SPS_MemberByMember"
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
