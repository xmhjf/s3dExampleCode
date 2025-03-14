VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFaceWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFaceWebSel.cls
'
'  Description :
'       Selector for Selection of Member bounded by Member End Cuts: Flange case
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/SEP/2011 - CM - TR-CP-202226,Enabled Genric AC's for Tube(B'ded) Vs Std Mbr(B'ding)
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "GenFaceWebSel"
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
    
    'Questions for Connecting to the Bounding Edges
    pQH.SetQuestion "ConnectTop", "No", "BooleanCol", "CMSetConnectTopAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectFace", "No", "BooleanCol", "CMSetConnectFaceAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectBottom", "No", "BooleanCol", "CMSetConnectBottomAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"

    'Questions for Creating Corner Features
    pQH.SetQuestion "FaceTopOutsideCorner", "No", "BooleanCol", "CMSetFaceTopOutsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "FaceBottomOutsideCorner", "No", "BooleanCol", "CMSetFaceBottomOutsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate centre cut
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sSelector As String

    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object

    Dim oSDO_WebCut As StructDetailObjects.WebCut

    sMsg = "Unknown Error"

    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject

    sMsg = "Set the Bounded object"
    Set oBoundedObject = oSDO_WebCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding object"
    Set oBoundingObject = oSDO_WebCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    ' Select Generic Web Cut based on type of Bounding Object
    sSelector = ""
    If TypeOf oBoundingObject Is IJPort Then
        ' --------------------------------------
        ' Pick item based on cross section alias
        ' --------------------------------------
        Dim lSectionAlias As Long
        Dim bPenetratesWeb As Boolean
        Dim oEdgeMapColl As JCmnShp_CollectionAlias
        
        Set oEdgeMapColl = New Collection
        Set oEdgeMapColl = GetEdgeMap(pSelectorLogic.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
        Select Case lSectionAlias
            Case 0, 8, 9, 10, 11, 12, 17
                'Web, WebTopFlangeLeft, WebBuiltUpTopFlangeLeft, WebBottomFlangeLeft
                'WebBuiltUpBottomFlangeLeft, WebTopAndBottomLeftFlanges, TwoWebsTwoFlanges,
                sSelector = "Web_F"
            Case 14
                'FlangeLeftAndRightBottomWebs
                sSelector = "WebB_F"
            Case 15
                'FlangeLeftAndRightTopWebs
                sSelector = "WebT_F"
            Case 16, 18
                'FlangeLeftAndRightWebs, TwoFlangesBetweenWebs
                sSelector = "WebTB_F"
            Case 20 ' Tube
                sSelector = ""
            Case Else
                'Unknown Section Alias
        End Select
    End If

    pSelectorLogic.Add sSelector

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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
