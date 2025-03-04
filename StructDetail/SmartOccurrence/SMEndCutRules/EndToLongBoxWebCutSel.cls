VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToLongBoxWebCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMEndCutRules
'  File        : EndToLongBoxWebCutSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    12/Jan/2012 - svsmylav
'           CR-174918: Updated code to select items for profile pseudo-knuckle, Convex case.
'    15/Feb/2012 - svsmylav
'           CR-174918: Updated code to to select item for 'Extend' Mfg. option.
'*********************************************************************************************

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToLongBoxWebCutSel.cls"

'------------- Give the Selector Name and ProgId
Const m_SelectorProgid As String = CUSTOMERID + "EndCutRules.EndToLongBoxWebCutSel"
Const m_SelectorName As String = CUSTOMERID + "EndCutRules.EndToLongBoxWebCutSel"
Const m_FamilyProgid As String = ""
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    sError = "Setting questions."
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList"
    pQH.SetQuestion "KnuckleEndCut", "LongBoxed", "ProfileKnuckleEndCutTypes"
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", sError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    Dim sError As String
    
    Dim sEndCutType As String
    Dim sSectionType As String
    Dim sKnuckleEndCut As String
    Dim sApplyTreatment As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundedPart As Object
    Dim oBoundedInput As Object
    
    'verify Bounded Object implements IJPort interface
    Set oBoundedInput = pSLH.InputObject(INPUT_BOUNDED)
    If oBoundedInput Is Nothing Then
        sError = "pSLH.InputObject(INPUT_BOUNDED) is not valid object (Nothing)"
        GoTo ErrorHandler
    ElseIf Not TypeOf oBoundedInput Is IJPort Then
        sError = "pSLH.InputObject(INPUT_BOUNDED) is not valid IJPort object"
        GoTo ErrorHandler
    End If
    
    Set oBoundedPort = oBoundedInput
    Set oBoundedPart = oBoundedPort.Connectable
    sSectionType = GetCrossSectionType(oBoundedPart)
    
    sEndCutType = pSLH.Answer(QUES_ENDCUTTYPE)
    sKnuckleEndCut = pSLH.Answer("KnuckleEndCut")
    sApplyTreatment = pSLH.Answer(gsApplyTreatment)
    
    Dim oProfileKnuckle As Object
    Dim iKnuckleMfgMethod As Integer
    Dim bIsConvex As Boolean
    Dim bCnvxSplitAndExtend As Boolean

    bIsConvex = False
    bCnvxSplitAndExtend = False
    
    GetProfileKnuckleType pSLH.SmartOccurrence, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex
    If iKnuckleMfgMethod = pkmmSplitAndExtend And bIsConvex Then bCnvxSplitAndExtend = True
    
    Dim bFreeEndCutExtend_Convex As Boolean
    If iKnuckleMfgMethod <> pkmmSplitAndExtend Then
        'Free EndCut this cut is Not because of 'SplitAndExtend', confirm that it is from 'Extend' also a convex knuckle
        bFreeEndCutExtend_Convex = IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedPort)
    End If
    
    'The Profile Knuckle "LogBoxed" is basically a Free End type
    Select Case sEndCutType
    
        Case "Snip"
            'If the profile is a Flat Bar
            If sSectionType = "FB" Then
                'Snipe profile endcut
                pSLH.Add "WebCut_LongBoxF1B"
            Else
                'Snipe profile endcut
                pSLH.Add "WebCut_LongBoxF1"
            End If
            
        Case "Welded"
            'Cut web normal with no physical connection
            If bCnvxSplitAndExtend Or bFreeEndCutExtend_Convex Then
                'We need to consider smart item where bounding is bent-plate
                pSLH.Add "WebCut_BendPlate"
            Else
                pSLH.Add "WebCut_LongBoxC1"
            End If
        
        Case Else
            ' Unknown EndCutType:: No Web Cut applied

    End Select
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", sError).Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

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

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    pSLH.Representation = pRep
    SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



