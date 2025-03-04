VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlangePenetratedSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFlangePenetratedSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Generic Port/Object End Cuts
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port/Object will be: (Bounding Object)
'           Plate Base/Offset/Lateral Face Port
'           Profile Base/Offset/ Lateral SubPort
'           MemberPart Base/Offset/ Lateral SubPort
'           Reference(Grid) Plane (IJPlane)
'           Point (IJPoint)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
''    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "GenFlangePenetratedSel"
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

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

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

    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"

    'Question: Create Top Flange Cut ?
    pQH.SetQuestion "CreateTopFlangeCut", "No", "BooleanCol", "CMSetTopFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenFlangePenetratedSel"
    'Question: Create Bottom Flange Cut ?
    pQH.SetQuestion "CreateBottomFlangeCut", "No", "BooleanCol", "CMSetBottomFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenFlangePenetratedSel"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate standard/normal cut
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler

    Dim sMsg As String

    Dim oEndCutObject As Object
    Dim oBoundedPort As Object
    Dim oBoundingPort As Object

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    sMsg = "Unknown Error"

    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    sMsg = "Set the Bounded input"
    Set oBoundedPort = oSDO_WebCut.BoundedPort
    If (oBoundedPort Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding input"
    Set oBoundingPort = oSDO_WebCut.BoundingPort
    If (oBoundingPort Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    '===========================================================
    'Get the Edge Mapping and Section Alias
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As New Collection

    Set oEdgeMapColl = GetEdgeMap(pSelectorLogic.SmartOccurrence, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    Select Case lSectionAlias
        Case 0, 8, 9, 10, 11, 12, 17
            'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
            'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
            pSelectorLogic.Add "GenFaceFlangeSel"
        Case 1, 6
            'WebTopFlangeRight, WebTopFlange
            pSelectorLogic.Add "GenTopEdgeFlangeSel"
        Case 2
            'WebBuiltUpTopFlangeRight
            pSelectorLogic.Add "GenBUTopEdgeFlangeSel"
        Case 3, 7
            'WebBottomFlangeRight, WebBottomFlange
            pSelectorLogic.Add "GenBtmEdgeFlangeSel"
        Case 4
            'WebBuiltUpBottomFlangeRight
            pSelectorLogic.Add "GenBUBtmEdgeFlangeSel"
        Case 5, 13, 19
            'WebTopAndBottomRightFlanges, WebTopAndBottomFlanges, TwoWebsBetweenFlanges
            'TR-CP-252628  Smart Classes Names not Updated in Detailing Rules to support Oracle
            pSelectorLogic.Add "GenTopAndBtmEdgeFlgSel"
        Case 14
            'FlangeLeftAndRightBottomWebs
            pSelectorLogic.Add "GenFaceFlangeSel"
        Case 15
            'FlangeLeftAndRightTopWebs
            pSelectorLogic.Add "GenFaceFlangeSel"
        Case 16, 18
            'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
            pSelectorLogic.Add "GenFaceFlangeSel"
        Case 20
            'Tube/Circular Cross Section
        Case Else
            'Unknown Section Alias
    End Select
        
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : CMSetTopFlangeCutAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetTopFlangeCutAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetTopFlangeCutAnswer"
    On Error GoTo ErrorHandler
    
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSelCM As New EndCutSelCM
    
    pSL.answer("CreateTopFlangeCut") = oSelCM.ComputeTopFlangeCutAnswer(pSL.SmartOccurrence)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Sub

'*********************************************************************************************
' Method      : CMSetBottomFlangeCutAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetBottomFlangeCutAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                      ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetBottomFlangeCutAnswer"
    On Error GoTo ErrorHandler
    
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSelCM As New EndCutSelCM
    
    pSL.answer("CreateBottomFlangeCut") = oSelCM.ComputeBottomFlangeCutAnswer(pSL.SmartOccurrence)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Sub


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
    pIH.definition = pSelector
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
