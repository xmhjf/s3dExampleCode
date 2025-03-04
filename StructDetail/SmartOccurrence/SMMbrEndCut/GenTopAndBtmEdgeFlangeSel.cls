VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenTopAndBtmEdgeFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenTopAndBtmEdgeFlangeSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Member End Cuts: Flange case
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "GenTopAndBtmEdgeFlangeSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
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
    pQH.SetQuestion "ConnectTopEdge", "No", "BooleanCol", "CMSetConnectTopEdgeAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectTopInside", "No", "BooleanCol", "CMSetConnectTopInsideAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectFace", "No", "BooleanCol", "CMSetConnectFaceAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectBottomInside", "No", "BooleanCol", "CMSetConnectBottomInsideAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectBottomEdge", "No", "BooleanCol", "CMSetConnectBottomEdgeAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "ConnectBottom", "No", "BooleanCol", "CMSetConnectBottomAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    
    'Questions for Creating Corner Features
    pQH.SetQuestion "TopEdgeOutsideCorner", "No", "BooleanCol", "CMSetTopEdgeOutsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "TopEdgeInsideCorner", "No", "BooleanCol", "CMSetTopEdgeInsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "FaceTopInsideCorner", "No", "BooleanCol", "CMSetFaceTopInsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "FaceBottomInsideCorner", "No", "BooleanCol", "CMSetFaceBottomInsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "BottomEdgeInsideCorner", "No", "BooleanCol", "CMSetBottomEdgeInsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "BottomEdgeOutsideCorner", "No", "BooleanCol", "CMSetBottomEdgeOutsideCornerAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"

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
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Unknown Error"

    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oEndCutObject
    
    sMsg = "Set the Bounded input"
    Set oBoundedObject = oSDO_FlangeCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding input"
    Set oBoundingObject = oSDO_FlangeCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    ' Select Generic Flange Cut based on type of Bounding Object
    If TypeOf oBoundingObject Is IJPort Then
        If IsTubularMember(oBoundedObject) Then
            '
        Else
            pSelectorLogic.Add "Flange_TE_F_BE"
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
