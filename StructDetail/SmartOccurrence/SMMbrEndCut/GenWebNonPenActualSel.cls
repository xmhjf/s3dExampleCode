VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebNonPenActualSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : GenWebNonPenActualSel.cls
'
'Author : Alligators
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
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'   available Selector Items:
'
'
'History:
'
'    12/SEP/2011 - CM - TR-CP-202226,Enabled Genric AC's for Tube(B'ded) Vs Std Mbr(B'ding)
'********************************************************************

Option Explicit

Const m_sClassName As String = "GenWebNonPenActualSel"
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

    ' Parent has question for PC toward bounding object
    
    'Question: Create Top PC ?
    pQH.SetQuestion "ConnectTop", "No", "BooleanCol", "CMSetTopPCAnswer", m_SelectorProgid
    'Question: Create Bottom PC ?
    pQH.SetQuestion "ConnectBottom", "No", "BooleanCol", "CMSetBottomPCAnswer", m_SelectorProgid
    'Question: Create Top Corner Feature ?
    pQH.SetQuestion "TopCorner", "No", "BooleanCol", "CMSetTopCornerAnswer", m_SelectorProgid
    'Question: Create Bottom Corner Feature ?
    pQH.SetQuestion "BottomCorner", "No", "BooleanCol", "CMSetBottomCornerAnswer", m_SelectorProgid

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
End Sub

' Select the Basic Type of Member Assembly Connection here
Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Unknown Error"

    On Error GoTo ErrorHandler
    
    ' ----------------------
    ' Get the web cut object
    ' ----------------------
    Dim oEndCutObject As Object
    Set oEndCutObject = oSL.SmartOccurrence
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    ' ------------------------------------
    ' Get the bounded and bounding objects
    ' ------------------------------------
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
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

    ' --------------------------------
    ' Check the validity of the inputs
    ' --------------------------------
    Dim lStatus As Long ' reports validity
    Dim oBoundedData As MemberConnectionData ' not used
    Dim oBoundingData As MemberConnectionData ' not used
    
    InitEndCutConnectionData oBoundedObject, oBoundingObject, _
                             oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        sMsg = "EndCut Ports are Not valid"
        GoTo ErrorHandler
    End If

    ' -------------------------------------
    ' If the bounding object is connectable
    ' -------------------------------------
    If TypeOf oBoundingObject Is IJPort Then

        ' -----------------------------
        ' pick the basic cut
        ' -----------------------------
        oSL.Add "FlushActual"

    ' ----------------------------------------------
    ' If not a connectable, select the standard type
    ' ----------------------------------------------
    ' FullDepth should probably always be "Yes" when not a connectable
    Else
        oSL.Add "FlushActual"
    End If

    Exit Sub

ErrorHandler:
    oSL.ReportError sMsg, METHOD
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

Public Sub CMSetTopCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                ByRef ppArgument As Object)

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSelCM As New EndCutSelCM
    
    pSL.answer("TopCorner") = "No"
    
    Exit Sub
End Sub

Public Sub CMSetBottomCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSelCM As New EndCutSelCM
    pSL.answer("BottomCorner") = "No"
    
    Exit Sub
End Sub

Public Sub CMSetTopPCAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                            ByRef ppArgument As Object)

    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    oSL.answer("ConnectTop") = "No"

    Exit Sub
    
End Sub

Public Sub CMSetBottomPCAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                      ByRef ppArgument As Object)

    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    oSL.answer("ConnectBottom") = "No"

    Exit Sub

End Sub

