VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootGenWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : RootGenWebSel.cls
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
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    06/Sep/2011 - pnalugol
'           for supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    03/August/2015 - MDT TR-274273 Change SM orientation from vertical to horizontal- deletes PCs under GenMemACs
'*********************************************************************************************

Const m_sClassName As String = "RootGenWebSel"
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
    
    pQH.SetQuestion "SelectByRule", "Yes", "BooleanCol"

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

    Dim oEndCutObject As Object
    Dim oBoundedPort As Object
    Dim oBoundingPort As Object

    Dim oSDO_WebCut As StructDetailObjects.WebCut

    Dim oConnectable As IJConnectable
    Dim oPort As IJPort

    sMsg = "Unknown Error"

    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    sMsg = "Set the Bounded object"
    Set oBoundedPort = oSDO_WebCut.BoundedPort
    If (oBoundedPort Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding object"
    Set oBoundingPort = oSDO_WebCut.BoundingPort
    If (oBoundingPort Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    Set oPort = oBoundingPort
    Set oConnectable = oPort.Connectable
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    
    bIsBoundedTube = IsTubularMember(oBoundedPort.Connectable)
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    
    '202055
    If bIsBoundingTube And bIsBoundedTube Then
         pSelectorLogic.Add "GenTubeByTubeSel"
        Exit Sub
    ElseIf bIsBoundingTube Then
        pSelectorLogic.Add "GenByTubeSel"
        Exit Sub
    End If
   
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = True
    
    'Get the Edge Mapping data here to cache for the cases mentioned in TR-274273
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartPrismatic Or TypeOf oBoundingPort.Connectable Is IJProfile Then
        bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
        Set oEdgeMapColl = GetEdgeMap(oSDO_WebCut.object, oBoundingPort, oBoundedPort)
    End If

    'If is Member or profile then..
    If TypeOf oConnectable Is ISPSMemberPartPrismatic Or TypeOf oConnectable Is IJProfile Then
        'If is Penetrated then use WebPenetratedSel
        If bPenetratesWeb Then
            pSelectorLogic.Add "GenWebPenetratedSel"
            'If NOT penetrated then use give BOTH options
        Else
            pSelectorLogic.Add "GenWebNonPenetratedSel"
            pSelectorLogic.Add "GenWebPenetratedSel"
        End If

        'If is NOT Member then use WebNonPenetratedSel
    Else
        pSelectorLogic.Add "GenWebNonPenetratedSel"
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




