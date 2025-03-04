VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : MbrEndFlangeSel.cls
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

Const m_sClassName As String = "MbrEndFlangeSel"
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
' Description :
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    pQH.SetQuestion "WeldPart", "First"
    pQH.SetQuestion "BottomFlange", "No", "BooleanCol"
    
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
    
    Dim oEndCutObject As Object
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
       
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Dim oWebCutObject As Object
    sMsg = "Unknown Error"
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oEndCutObject
    
    Set oBoundedObject = oSDO_FlangeCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
        
    Set oBoundingObject = oSDO_FlangeCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    Set oWebCutObject = oSDO_FlangeCut.WebCut
    If (oWebCutObject Is Nothing) Then
        sMsg = "Web Cut Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    Dim sBottomFlange As String
    sBottomFlange = pSelectorLogic.answer("BottomFlange")
    
    Dim sWebSelection As String
    sWebSelection = ""
    
    If TypeOf oEndCutObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oWebCutObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        sWebSelection = oSmartItem.Name
    End If
    
    Dim oAppConn As IJAppConnection
    Dim oParentObj As Object
    Dim sParentItemName As String
    Dim sFlangeTypeCase As String
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    Dim bFreeEndCut As Boolean
    
    CrossSection_Flanges oSDO_FlangeCut.Bounded, bTFL, bBFL, bTFR, bBFR
    
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    'in case of long box
    If StrComp(sParentItemName, "Generic_LongBox", vbTextCompare) = 0 Or _
        StrComp(sParentItemName, "Generic_LongBoxCuts", vbTextCompare) = 0 Then
        If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
            pSelectorLogic.Add "MbrEndFlg_LB_Straight_LR"
        Else
            pSelectorLogic.Add "MbrEndFlg_LB_Straight_R"
        End If
        Exit Sub
    End If
    
    bFreeEndCut = oSDO_FlangeCut.IsFreeEndCut
    
    If bFreeEndCut Then
        'Handle cases for Free End Cuts on Flange
        If StrComp(sWebSelection, "MbrWeb_FreeAngledCut", vbTextCompare) = 0 Then
           'For this Web Cut currently we are not supporting any Flange Cut(as of now)
           'since this Web Cut has "CutVang" parameter we shall need to have a special
           'symbol to handle. For Now we are exiting
           
           Exit Sub
           
        End If
        
        If ((bTFL) And (bTFR) Or _
            (bBFL) And (bBFR)) Then
            pSelectorLogic.Add "MbrFree_Flange_Sniped_LR"
            pSelectorLogic.Add "MbrFree_Flange_Straight_LR"
        Else
            pSelectorLogic.Add "MbrFree_Flange_Sniped_R"
            pSelectorLogic.Add "MbrFree_Flange_Straight_R"
        End If
    
        Exit Sub
    End If
    
    Set oAppConn = oParentObj

    
    If GetMbrAssemblyConnectionType(oAppConn) = eACType.ACType_Split Then
        ' Retreive the FlangeType from the Owning AssemblyConnection
        GetSelectorAnswer oParentObj, "SplitEndToEndCase", sFlangeTypeCase
        
        'No flange cut created for SeamAngle cases; so these cases not included in sel logic
        If InStr(LCase(sFlangeTypeCase), LCase("NoAngle")) > 0 Then
            If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
                    pSelectorLogic.Add "MbrEndFlg_Straight_LR"   'straight cut
            Else
                    pSelectorLogic.Add "MbrEndFlg_Straight_R"
            End If
        Else
            If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
                    pSelectorLogic.Add "MbrEndFlg_Straight_LR"   'straight cut
            Else
                    pSelectorLogic.Add "MbrEndFlg_Straight_R"
            End If
        End If
    ElseIf GetMbrAssemblyConnectionType(oAppConn) = eACType.ACType_Miter Then
        Select Case sWebSelection
            Case "MbrMiterWeb"
                If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
                    pSelectorLogic.Add "MbrEndFlg_Straight_LR"   'straight cut
                Else
                    pSelectorLogic.Add "MbrEndFlg_Straight_R"
                End If
            Case "MbrMiterWeb_SS"
                If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
                    pSelectorLogic.Add "MbrEndFlg_StraightSetback_LR"
                Else
                    pSelectorLogic.Add "MbrEndFlg_StraightSetback_R"
                End If
            Case "MbrMiterWeb_SB"
                If sBottomFlange = "Yes" Then
                    If bBFL Then
                        pSelectorLogic.Add "MbrEndFlg_StraightSetback_LR"
                    Else
                        pSelectorLogic.Add "MbrEndFlg_StraightSetback_R"
                    End If
                Else
                    If bTFL Then
                        pSelectorLogic.Add "MbrEndFlg_Straight_LR"
                    Else
                        pSelectorLogic.Add "MbrEndFlg_Straight_R"
                    End If
                End If
            Case "MbrMiterWeb_ST"
                If sBottomFlange = "Yes" Then
                    If bBFL Then
                        pSelectorLogic.Add "MbrEndFlg_Straight_LR"
                    Else
                        pSelectorLogic.Add "MbrEndFlg_Straight_R"
                    End If
                Else
                    If bTFL Then
                        pSelectorLogic.Add "MbrEndFlg_StraightSetback_LR"
                    Else
                        pSelectorLogic.Add "MbrEndFlg_StraightSetback_R"
                    End If
                End If
            Case Else
                If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
                    pSelectorLogic.Add "MbrEndFlg_Straight_LR"
                Else
                    pSelectorLogic.Add "MbrEndFlg_Straight_R"
                End If
        End Select
    Else
        If ((sBottomFlange = "Yes") And (bBFL = True)) Or ((sBottomFlange = "No") And bTFL = True) Then
            pSelectorLogic.Add "MbrEndFlg_Straight_LR"   'straight cut
        Else
            pSelectorLogic.Add "MbrEndFlg_Straight_R"
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

