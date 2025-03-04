VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StifEndToFCAndOS1EdgeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : StifEndToFCAndOS1EdgeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/NOV/2011 - Created
'
'*********************************************************************************************


Const m_sClassName As String = "StifEndToFCAndOS1EdgeDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
ErrorHandler:
    pIH.ReportError strError, METHOD
    
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "ItemAggregator").Number
    
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'     Axis EndCuts can consist of (3) Web Cut and/or (3) Flange Cuts, but all the cuts
'     cannot be applied together. Depending on the orientation either the 3 Web Cuts
'     and 1 Flange Cut will get applied or vice versa(1 Web Cut and 3 Flange Cuts)
'
'     Index(1) - Web Cut for Bounded Member(Web Center Cut ---> if Web Penetrated OR
'                                           Web Normal Cut ---> if Web Non Penetrated)
'     Index(2) - Top Web Cut ---> Applied if (Web Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'     Index(3) - Bottom Web Cut ---> Applied if (Web Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'     Index(4) - Flange Cut for Top Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                      Flange Normal Cut ---> if Web Penetrated)
'                 *****This Index(4) item will be created only if the TopWebCut doesnt exist (i.e if Index(2) Item Member conditional is False)
'                 *****if TopWebCut exist the next item Member(i.e Index(5)) will create the Flange Cut
'
'     Index(5) - Flange Cut for Top Flange(if Exist) of Bounded Member
'                This item will be created only if the TopWebCut exist (i.e if Index(2) Item Member conditional is True)
'                index(5) item member is Vice versa of Index(4)Item Member
'     Index(6) - Top Flange Cut for Top Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'     Index(7) - Bottom Flange Cut for Top Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'     Index(8) - Flange Cut for Bottom Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                         Flange Normal Cut ---> if Web Penetrated)
'                 *****This Index(8) item will be created only if the BottomWebCut doesnt exist (i.e if Index(3) Item Member conditional is False)
'                 *****if BottomWebCut exist the next item Member(i.e Index(9)) will create the Flange Cut
'
'     Index(9) - Flange Cut for Bottom Flange(if Exist) of Bounded Member
'                This item will be created only if the BottomWebCut exist (i.e if Index(3) Item Member conditional is True)
'                index(9) item member is Vice versa of Index(8)Item Member
'     Index(10) - Top Flange Cut for Bottom Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'     Index(11) - Bottom Flange Cut for Bottom Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sDispId As String
    
    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    
    '******************************************************************************
    ' if any more item members are needed to be added please add only at the bottom
    '******************************************************************************
    
    ' should be responsible for creation of Web Cut(Center Cut or Normal Cut)
    lDispId = 1
    sDispId = "StiffEndToMbrFace_WebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Top Web Cut
    lDispId = 2
    sDispId = "StiffEndToMbrFace_TopWebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopWebCutNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute2", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Bottom Web Cut
    lDispId = 3
    sDispId = "StiffEndToMbrFace_BtmWebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmWebCutNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute3", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' The next Item Member should be responsible for creation of Flange Cut on Top
    '    --->Center Cut or
    '    --->Normal Cut(when TopWebCut doesn't exist)
    ' The item Member will create the Cut only when Top Web Cut doesnt exist
    lDispId = 4
    sDispId = "StiffEndToMbrFace_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    Set oMemDesc = Nothing
    
    ' The next Item Member should be responsible for creation of Flange Cut on Top
    '     --->Normal Cut(Only when TopWebCut exist)
    ' The item Member will create the Cut only when Top Web Cut exist(because TopWebCut will act as Input to Flange Cut)
    lDispId = 5
    sDispId = "StiffEndToMbrFace_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Top Flange Cut for Top Flange
    lDispId = 6
    sDispId = "StiffEndToMbrFace_TopFlangeCutForTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutForTopFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Bottom Flange Cut for Top Flange
    lDispId = 7
    sDispId = "StiffEndToMbrFace_BtmFlangeCutForTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutForTopFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Flange Cut on Bottom(Ceneter Cut or Normal Cut)
    lDispId = 8
    sDispId = "StiffEndToMbrFace_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Flange Cut on Bottom(Ceneter Cut or Normal Cut)
    lDispId = 9
    sDispId = "StiffEndToMbrFace_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Top Flange Cut for Bottom Flange
    lDispId = 10
    sDispId = "StiffEndToMbrFace_TopFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutForBtmFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    Set oMemDesc = Nothing
    
    ' should be responsible for creation of Bottom Flange Cut for Bottom Flange
    lDispId = 11
    sDispId = "StiffEndToMbrFace_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutForBtmFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

' ** Start CM **
' *******************************************************************************************
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
' *******************************************************************************************
Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
    Const METHOD = m_DefinitionProgid & "::CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    bIsNeeded = True
 
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
'
End Sub

'*********************************************************************************************
' Method      : CM_IsTopWebCutNeeded
' Description : Conditional method to determine if a Top web cut is needed
'
'********************************************************************************************
Public Sub CM_IsTopWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
    Const METHOD = m_DefinitionProgid & "::CM_IsTopWebCutNeeded"
    On Error GoTo ErrorHandler

    Dim sMsg As String

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    ' Set the Bounded and Bounding Port
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim cTop As ConnectedEdgeInfo
    Dim cBottom As ConnectedEdgeInfo
    Dim cInsideTF As ConnectedEdgeInfo
    Dim cInsideBF As ConnectedEdgeInfo

    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, cTop, cBottom, cInsideTF, cInsideBF

    Dim sShapeAtFaceAnswer As String
    Dim sShapeOutsideAnswer As String, sTopBraceTypeAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases
    If bPenetratesWeb Then
        'We have two cases: 1. Bounded-top is above/touching top of bounding flange
        ' or 2) Bounded-top is intersecting web of bounding member
        If cTop.IntersectingEdge = eBounding_Edge.Above Or cTop.IntersectingEdge = eBounding_Edge.Bottom Or _
           cTop.IntersectingEdge = eBounding_Edge.Top Or cTop.IntersectingEdge = eBounding_Edge.Below Then
            'Case 1
            GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
            Select Case LCase(sShapeOutsideAnswer)
                Case LCase(gsStraight), LCase(gsSniped)
                    bIsNeeded = True
                Case Else
                    '
            End Select
        ElseIf cTop.IntersectingEdge = eBounding_Edge.Web_Right Then
            'Case 2
            GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
            Select Case LCase(sShapeAtFaceAnswer)
                Case LCase(gsCope), LCase(gsSnipe)
                    bIsNeeded = True
                Case Else
                    '
            End Select
        End If
    Else
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceType", sTopBraceTypeAnswer
        If sTopBraceTypeAnswer = "InsetMember" Then
            bIsNeeded = True
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsBtmWebCutNeeded
' Description : Conditional method to determine if a Bottom web cut is needed
'
'********************************************************************************************
Public Sub CM_IsBtmWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmWebCutNeeded"
    On Error GoTo ErrorHandler

    Dim sMsg As String

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    ' Set the Bounded and Bounding Port
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim cTop As ConnectedEdgeInfo
    Dim cBottom As ConnectedEdgeInfo
    Dim cInsideTF As ConnectedEdgeInfo
    Dim cInsideBF As ConnectedEdgeInfo

    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, cTop, cBottom, cInsideTF, cInsideBF


    Dim sShapeAtFaceAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases
    If bPenetratesWeb Then
        'We have two cases: 1. Bounded-bottom is below/touching bottom of bounding flange
        ' or 2) Bounded-top is intersecting web of bounding member
        If cBottom.IntersectingEdge = eBounding_Edge.Below Or _
          cBottom.IntersectingEdge = eBounding_Edge.Bottom Then
            'Case 1
            Dim sShapeOutsideAnswer As String
            GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
            Select Case LCase(sShapeOutsideAnswer)
                Case LCase(gsStraight), LCase(gsSniped)
                    bIsNeeded = True
                Case Else
                    '
            End Select
        ElseIf cBottom.IntersectingEdge = eBounding_Edge.Web_Right Then
            'Case 2
            GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
            Select Case LCase(sShapeAtFaceAnswer)
                Case LCase(gsCope), LCase(gsSnipe)
                    bIsNeeded = True
                Case Else
                    'Unknown Case
            End Select
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsTopFlangeCutNeeded
' Description : Conditional method to determine if a Top Flange cut is needed
' Only when the Top flange exists, this cut is applied. It could be a centre cut, if a Flange Penetarted case; a standard/normal
' cut, if web penetarted case.
'********************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    bIsNeeded = False

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Set oAppConnection = pMemberDescription.CAO
    Dim lStatus As Long

    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = pMemberDescription.CAO

    If pMemberDescription.dispid = 4 Then
       If Not oMemberObjects.ItemByDispid(2) Is Nothing Then
         'Top Web Cut exist
         bIsNeeded = False
         Exit Sub
       Else
         'TopWebCut doesnt exist
         'Continue
       End If
    ElseIf pMemberDescription.dispid = 5 Then
       If Not oMemberObjects.ItemByDispid(2) Is Nothing Then
         'Top Web Cut exist
         'Continue
       Else
         'TopWebCut doesnt exist
         bIsNeeded = False
         Exit Sub
       End If
    End If

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR

    If (bTopFL) Or (bTopFR) Then
       bIsNeeded = True
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsTopFlangeCutForTopFlangeNeeded
' Description : Conditional method to determine if a Top cut for top Flange is needed
' Only when the Top flange exists and if a Flange Penetrated case, this cut is applied.
'********************************************************************************************
Public Sub CM_IsTopFlangeCutForTopFlangeNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutForTopFlangeNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    Set oAppConnection = pMemberDescription.CAO

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim WL As ConnectedEdgeInfo
    Dim WR As ConnectedEdgeInfo
    Dim TFL As ConnectedEdgeInfo
    Dim TFR As ConnectedEdgeInfo


    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR


    Dim sShapeAtFaceAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases
    If (bTopFL) Or (bTopFR) Then

        bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
        If Not bPenetratesWeb Then
             GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, WL, WR, TFL, TFR
            'We have two cases: 1. Bounded intersecting edge is above the bounding flange edge, or
            ' or 2) intersecting with web of bounding member
            Dim iCaseFlag As Integer
            iCaseFlag = 2 'Default is considered as 2; set to 1 if it is otherwise
            If bTopFL Then
                 If TFL.IntersectingEdge = eBounding_Edge.Above Or _
                    TFL.IntersectingEdge = eBounding_Edge.Top Then
                    iCaseFlag = 1 'Case 1
                 End If
            Else
                 If WL.IntersectingEdge = eBounding_Edge.Above Or _
                    WL.IntersectingEdge = eBounding_Edge.Top Then
                    iCaseFlag = 1 'Case 1
                 End If
            End If

            If iCaseFlag = 1 Then 'Case 1
                'Check the answer to the question defined at selector level
                Dim sShapeOutsideAnswer As String
                GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
                Select Case LCase(sShapeOutsideAnswer)
                    Case LCase(gsStraight), LCase(gsSniped)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            ElseIf iCaseFlag = 2 Then 'Case 2
                'Check the answer to the question defined at selector level
                GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
                Select Case LCase(sShapeAtFaceAnswer)
                    Case LCase(gsCope), LCase(gsSnipe)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsBtmFlangeCutForTopFlangeNeeded
' Description : Conditional method to determine if a Bottom cut for top Flange is needed
' Only when the Top flange exists and if a Flange Penetrated case, this cut is applied.
'********************************************************************************************
Public Sub CM_IsBtmFlangeCutForTopFlangeNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutForTopFlangeNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Set oAppConnection = pMemberDescription.CAO

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim WL As ConnectedEdgeInfo
    Dim WR As ConnectedEdgeInfo
    Dim TFL As ConnectedEdgeInfo
    Dim TFR As ConnectedEdgeInfo

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR


    Dim sShapeAtFaceAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases

    If (bTopFL) Or (bTopFR) Then
        bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
        If Not bPenetratesWeb Then
             GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, WL, WR, TFL, TFR
            'We have two cases: 1. Bounded intersecting edge is below the bounding flange edge, or
            ' or 2) intersecting with web of bounding member
            Dim iCaseFlag As Integer
            iCaseFlag = 2 'Default is considered as 2; set to 1 if it is otherwise
            If bTopFR Then
                 If TFR.IntersectingEdge = eBounding_Edge.Below Or _
                    TFR.IntersectingEdge = eBounding_Edge.Bottom Then
                        iCaseFlag = 1 'Case 1
                 End If
            Else
                 If WR.IntersectingEdge = eBounding_Edge.Below Or _
                    WR.IntersectingEdge = eBounding_Edge.Bottom Then
                        iCaseFlag = 1 'Case 1
                 End If
            End If

            If iCaseFlag = 1 Then 'Case 1
                'Check the answer to the question defined at selector level
                Dim sShapeOutsideAnswer As String
                GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
                Select Case LCase(sShapeOutsideAnswer)
                    Case LCase(gsStraight), LCase(gsSniped)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            ElseIf iCaseFlag = 2 Then 'Case 2
                'Check the answer to the question defined at selector level
                GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
                Select Case LCase(sShapeAtFaceAnswer)
                    Case LCase(gsCope), LCase(gsSnipe)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsBtmFlangeCutNeeded
' Description : Conditional method to determine if a Bottom cut for top Flange is needed
' Only when the Bottom flange exists, this cut is applied.
'********************************************************************************************
Public Sub CM_IsBtmFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    bIsNeeded = False

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim lStatus As Long
    Set oAppConnection = pMemberDescription.CAO

    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = pMemberDescription.CAO

    If pMemberDescription.dispid = 8 Then
       If Not oMemberObjects.ItemByDispid(3) Is Nothing Then
         'Bottom Web Cut exist
         bIsNeeded = False
         Exit Sub
       Else
         'Bottom Web Cut doesnt exist
         'Continue
       End If
    ElseIf pMemberDescription.dispid = 9 Then
       If Not oMemberObjects.ItemByDispid(3) Is Nothing Then
         'Bottom Web Cut exist
         'Continue
       Else
         'Bottom Web Cut doesnt exist
         bIsNeeded = False
         Exit Sub
       End If
    End If

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg


    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR

    If (bBtmFL) Or (bBtmFR) Then
       bIsNeeded = True
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsTopFlangeCutForBtmFlangeNeeded
' Description : Conditional method to determine if a Top cut for Bottom Flange is needed
' Only when the Bottom flange exists and if a Flange Penetrated case, this cut is applied.
'********************************************************************************************
Public Sub CM_IsTopFlangeCutForBtmFlangeNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
'
Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutForBtmFlangeNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    bIsNeeded = False

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    Set oAppConnection = pMemberDescription.CAO

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim WL As ConnectedEdgeInfo
    Dim WR As ConnectedEdgeInfo
    Dim BFL As ConnectedEdgeInfo
    Dim BFR As ConnectedEdgeInfo

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR


    Dim sShapeAtFaceAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases
    If (bBtmFL) Or (bBtmFR) Then

        bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
        If Not bPenetratesWeb Then
             GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, WL, WR, BFL, BFR
            'We have two cases: 1. Bounded intersecting edge is above the bounding flange edge, or
            ' or 2) intersecting with web of bounding member
            Dim iCaseFlag As Integer
            iCaseFlag = 2 'Default is considered as 2; set to 1 if it is otherwise
            If bBtmFL Then
                 If BFL.IntersectingEdge = eBounding_Edge.Above Or _
                    BFL.IntersectingEdge = eBounding_Edge.Top Then
                        iCaseFlag = 1 'Case 1
                 End If
            Else
                 If WL.IntersectingEdge = eBounding_Edge.Above Or _
                    WL.IntersectingEdge = eBounding_Edge.Top Then
                        iCaseFlag = 1 'Case 1
                 End If
            End If

            If iCaseFlag = 1 Then 'Case 1
                'Check the answer to the question defined at selector level
                Dim sShapeOutsideAnswer As String
                GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
                Select Case LCase(sShapeOutsideAnswer)
                    Case LCase(gsStraight), LCase(gsSniped)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            ElseIf iCaseFlag = 2 Then 'Case 2
                'Check the answer to the question defined at selector level
                GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
                Select Case LCase(sShapeAtFaceAnswer)
                    Case LCase(gsCope), LCase(gsSnipe)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsBtmFlangeCutForBtmFlangeNeeded
' Description : Conditional method to determine if a Bottom cut for Bottom Flange is needed
' Only when the Bottom flange exists and if a Flange Penetrated case, this cut is applied.
'********************************************************************************************
Public Sub CM_IsBtmFlangeCutForBtmFlangeNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutForBtmFlangeNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim bPenetratesWeb As Boolean
    Dim lStatus As Long

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Set oAppConnection = pMemberDescription.CAO

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    Dim TopOrWL As ConnectedEdgeInfo
    Dim WR As ConnectedEdgeInfo
    Dim BFL As ConnectedEdgeInfo
    Dim BFR As ConnectedEdgeInfo

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR


    Dim sShapeAtFaceAnswer As String

    bIsNeeded = False  'Initial value, make it to True only for valid cases
    If (bBtmFL) Or (bBtmFR) Then

        bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
        If Not bPenetratesWeb Then
             GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, TopOrWL, WR, BFL, BFR
            'We have two cases: 1. Bounded intersecting edge is below the bounding flange edge, or
            ' or 2) intersecting with web of bounding member
            Dim iCaseFlag As Integer
            iCaseFlag = 2 'Default is considered as 2; set to 1 if it is otherwise
            If bBtmFR Then
                 If BFR.IntersectingEdge = eBounding_Edge.Below Or _
                    BFR.IntersectingEdge = eBounding_Edge.Bottom Then
                        iCaseFlag = 1 'Case 1
                 End If
            Else
                 If WR.IntersectingEdge = eBounding_Edge.Below Or _
                    WR.IntersectingEdge = eBounding_Edge.Bottom Then
                        iCaseFlag = 1 'Case 1
                 End If
            End If

            If iCaseFlag = 1 Then 'Case 1
                'Check the answer to the question defined at selector level
                Dim sShapeOutsideAnswer As String
                GetSelectorAnswer pMemberDescription.CAO, "ShapeOutside", sShapeOutsideAnswer
                Select Case LCase(sShapeOutsideAnswer)
                    Case LCase(gsStraight), LCase(gsSniped)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            ElseIf iCaseFlag = 2 Then 'Case 2
                'Check the answer to the question defined at selector level
                GetSelectorAnswer pMemberDescription.CAO, "ShapeAtFace", sShapeAtFaceAnswer
                Select Case LCase(sShapeAtFaceAnswer)
                    Case LCase(gsCope), LCase(gsSnipe)
                        bIsNeeded = True
                    Case Else
                        '
                End Select
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique

    IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'********************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Dim pAPDs As IJDPropertyDescriptions
    Set pAD = pDefinition
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Removes all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Removes all the previous Member descriptions
    
    ItemMembers pMDs
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description
'********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                            ByVal defParams As Variant, _
                                                            ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_DefinitionProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description
'********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function
