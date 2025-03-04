VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrSplitWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : MbrSplitWebSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Web Weld case
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port will be SPSMemberAxisAlong (Bounding Member)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'  Author      : Alligators
'
'  History     :
'    28/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "MbrSplitWebSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices
'
'*********************************************************************************************
' Method      : SelectorInputs
' Description :
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    pQH.SetQuestion "WeldPart", "First"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description :
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lStatus As Long
    
    Dim sWebTypeCase As String
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    Dim sIdealizedBounded As String
    Dim sIdealizedBoundary As String
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    sMsg = "Unknown Error"
    Dim oAppConnection As IJAppConnection
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
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
        
    CheckIdealizedBoundary oBoundedData, oBoundingData, sIdealizedBoundary
    CheckIdealizedBoundary oBoundingData, oBoundedData, sIdealizedBounded
    
    ' Retreive the WebType from the Owning AssemblyConnection
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    Dim sWeldPart As String
    GetSelectorAnswer oParentObj, "SplitEndToEndCase", sWebTypeCase
    sWeldPart = pSelectorLogic.Answer("WeldPart")
    
    ' Case: End To End / Colinear Axis / Identical CrossSection / same Configuration
    ' Check if Assembly Connection Members Cross Sections are Identical
    Dim bIdentical As Boolean
    AreCrossSectionsIdentical oBoundedData, oBoundingData, bIdentical

    If Trim(LCase(sIdealizedBounded)) = LCase(eIdealized_BoundingTube) Then
        pSelectorLogic.Add "MbrSplitWeb_Tube"
    ElseIf InStr(LCase(sWebTypeCase), LCase("NoAngle")) > 0 Then
        pSelectorLogic.Add "MbrSplitWeb_NoAngle"
    ElseIf bIdentical Then
        If InStr(LCase(sWebTypeCase), LCase("AngleWebSquareFlange")) > 0 Then
            If Trim(LCase(sWeldPart)) = LCase("First") Then
                pSelectorLogic.Add "MbrSplitWeb_AngleWebSqFl"
            Else
                pSelectorLogic.Add "MbrSplitWeb_AngleWebSqFl_2"
            End If
        ElseIf InStr(LCase(sWebTypeCase), LCase("AngleWebBevelFlange")) > 0 Then
            If Trim(LCase(sWeldPart)) = LCase("First") Then
                pSelectorLogic.Add "MbrSplitWeb_AngleWebBevelFl"
            Else
                pSelectorLogic.Add "MbrSplitWeb_AngleWebBevelFl_2"
            End If
        ElseIf InStr(LCase(sWebTypeCase), LCase("AngleWebAngleFlange")) > 0 Then
            If Trim(LCase(sWeldPart)) = LCase("First") Then
                pSelectorLogic.Add "MbrSplitWeb_AngleWebAngleFl"
            Else
                pSelectorLogic.Add "MbrSplitWeb_AngleWebAngleFl_2"
            End If
        ElseIf InStr(LCase(sWebTypeCase), LCase("DistanceWebDistanceFlange")) > 0 Then
            If Trim(LCase(sWeldPart)) = LCase("First") Then
                pSelectorLogic.Add "MbrSplitWeb_DistWebDistFl"
            Else
                pSelectorLogic.Add "MbrSplitWeb_DistWebDistFl_2"
            End If
        ElseIf InStr(LCase(sWebTypeCase), LCase("OffsetWebOffsetFlange")) > 0 Then
            If Trim(LCase(sWeldPart)) = LCase("First") Then
                pSelectorLogic.Add "MbrSplitWeb_OffsetWebOffsetFl"
            Else
                pSelectorLogic.Add "MbrSplitWeb_OffsetWebOffsetFl_2"
            End If
        Else    'No Angle case given as default
            pSelectorLogic.Add "MbrSplitWeb_NoAngle"
        End If
    Else 'No Angle case given as default
        pSelectorLogic.Add "MbrSplitWeb_NoAngle"
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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
' Method      : CMSelector
' Description :
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
