VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlgByTubeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFlgByTubeSel.cls
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
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    11/Jul/2012 - svsmylav
'           TR-206472(DM-216590): Additional argument 'bPenetratesWeb' is passed to 'GetSelForMbrBoundedToTube'.
'*********************************************************************************************

Const m_sClassName As String = "GenFlgByTubeSel"
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
    pQH.SetQuestion "ConnectTube", "No", "BooleanCol", "CMSetConnectTubeAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "CreateTopFlangeCut", "No", "BooleanCol", "CMSetTopFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenFlgByTubeSel"
    pQH.SetQuestion "CreateBottomFlangeCut", "No", "BooleanCol", "CMSetBottomFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenFlgByTubeSel"

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

    sMsg = "Selection logic"
    
    Dim oBoundedObject As Object, oBoundingObject As Object
    If TypeOf pSelectorLogic.InputObject(INPUT_BOUNDED) Is IJPort Then
        Dim oBoundedPort As IJPort
        Set oBoundedPort = pSelectorLogic.InputObject(INPUT_BOUNDED)
        Set oBoundedObject = oBoundedPort.Connectable
    Else
        Set oBoundedObject = pSelectorLogic.InputObject(INPUT_BOUNDED)
    End If
    
    Dim oDesignParent As IJDesignParent
    Dim ppChildren As IJDObjectCollection
    
    If TypeOf pSelectorLogic.InputObject(INPUT_BOUNDING) Is IJPort Then
        Dim oBoundingPort As IJPort
        Set oBoundingPort = pSelectorLogic.InputObject(INPUT_BOUNDING)
        Set oDesignParent = oBoundingPort.Connectable
        Set oBoundingObject = oDesignParent
    Else
        Set oDesignParent = pSelectorLogic.InputObject(INPUT_BOUNDING)
    End If
        
    Dim bIsBoundingTube As Boolean
    Dim selString As String
    
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    
    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    
    CrossSection_Flanges oBoundedPort.Connectable, bTopFL, bBtmFL, bTopFR, bBtmFR
    
    If bIsBoundingTube Then
        Dim bPenetratesWeb As Boolean
    
        bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
        If bPenetratesWeb Then
            If (bTopFL And bTopFR) Then
                pSelectorLogic.Add "Gen_Flange_SnipedSetback_Tube_LR"
                pSelectorLogic.Add "Gen_Flange_StraightSetback_Tube_LR"
                pSelectorLogic.Add "Gen_Flange_Straight_LR"
            Else
                pSelectorLogic.Add "Gen_Flange_SnipedSetback_Tube_R"
                pSelectorLogic.Add "Gen_Flange_StraightSetback_Tube_R"
                pSelectorLogic.Add "Gen_Flange_Straight_R"
            End If
            
        Else
            pSelectorLogic.Add "Gen_Flange_Tube"   'center cut
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
'*********************************************************************************************
' Method      : CMSetTopFlangeCutAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetTopFlangeCutAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetTopFlangeCutAnswer"
    On Error GoTo ErrorHandler
    
    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oSL.SmartOccurrence
    
    Dim oSelCM As New EndCutSelCM
    
    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim selString As String
    
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.Bounded
    Set oBounding = oSDO_WebCut.Bounding
    
    
    CrossSection_Flanges oBounded, bTopFL, bBtmFL, bTopFR, bBtmFR
    
    If bTopFL Or bTopFR Then
        
        Dim bPenetratesWeb As Boolean
        bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
        
        If Not bPenetratesWeb Then
            Dim eCase As eMemberBoundingCase
            eCase = GetBoundingCaseForTube(oBounded, oBounding)
            
             Select Case eCase
                Case Center, FCAndOSBottom, FCAndOSTop
                    oSL.answer("CreateTopFlangeCut") = "No"
                Case Else
                    oSL.answer("CreateTopFlangeCut") = "Yes"
            End Select
        Else
            oSL.answer("CreateTopFlangeCut") = "No"
        End If
    Else
        oSL.answer("CreateTopFlangeCut") = "No"
    End If
    
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
    
    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSelCM As New EndCutSelCM
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oSL.SmartOccurrence
      
    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim selString As String
    
    CrossSection_Flanges oSDO_WebCut.Bounded, bTopFL, bBtmFL, bTopFR, bBtmFR
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.Bounded
    Set oBounding = oSDO_WebCut.Bounding
    
    If bBtmFL Or bBtmFR Then
        Dim bPenetratesWeb As Boolean
        bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
        
        If Not bPenetratesWeb Then
            Dim eCase As eMemberBoundingCase
            eCase = GetBoundingCaseForTube(oBounded, oBounding)
            
            Select Case eCase
                 Case Center, FCAndOSBottom, FCAndOSTop
                    oSL.answer("CreateBottomFlangeCut") = "No"
                Case Else
                    oSL.answer("CreateBottomFlangeCut") = "Yes"
            End Select
        Else
            oSL.answer("CreateBottomFlangeCut") = "No"
        End If
    Else
        oSL.answer("CreateBottomFlangeCut") = "No"
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Sub

