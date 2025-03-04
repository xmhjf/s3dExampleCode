VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisToFCAndOSNoEdgeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAxisToFCAndOSNoEdgeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    22/aug/2011 - svsmylav
'           B-25298(CR-179926) : Conditionals of web cuts and flange cuts are modified to
'           provide seperate Flangecuts for standard member bounded to tube case.
'    25/Aug/2011 - pnalugol
'           Added new methods to support Inset brace (B-52232)
'    15/Sep/2011 - pnalugol
'           Modified/Added methods to support Braces on Flange penetrated cases
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'
'*********************************************************************************************

Const m_sClassName As String = "MbrAxisToFCAndOSNoEdgeDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Private Const TOLERANCE_VALUE = 0.0001   'Tolerence for Double comparison

'*Require by S:\StructuralModeling\Client\AsmConnCommands
Implements ISPSFACInputHelper
Implements IJUserAttributeMgmt

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
    
ErrorHandler:
    pIH.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"

    'CStructAssemblyConnection
    pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CM_FinalConstructAggregator"
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateAggregator"

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
' Axis EndCuts can consist of (3) Web Cut and/or (3) Flange Cuts, but all the cuts
' cannot be applied together. Depending on the orientation either the 3 Web Cuts
' and 1 Flange Cut will get applied or vice versa(1 Web Cut and 3 Flange Cuts)

' Index(1) - Web Cut for Bounded Member(Web Center Cut ---> if Web Penetrated OR
'                                       Web Normal Cut ---> if Web Non Penetrated)
' Index(2) - Top Web Cut ---> Applied if (Web Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(3) - Bottom Web Cut ---> Applied if (Web Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(4) - Flange Cut for Top Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                  Flange Normal Cut ---> if Web Penetrated)
'             *****This Index(4) item will be created only if the TopWebCut doesnt exist (i.e if Index(2) Item Member conditional is False)
'             *****if TopWebCut exist the next item Member(i.e Index(5)) will create the Flange Cut
'
' Index(5) - Flange Cut for Top Flange(if Exist) of Bounded Member
'            This item will be created only if the TopWebCut exist (i.e if Index(2) Item Member conditional is True)
'            index(5) item member is Vice versa of Index(4)Item Member
' Index(6) - Top Flange Cut for Top Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(7) - Bottom Flange Cut for Top Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(8) - Flange Cut for Bottom Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                     Flange Normal Cut ---> if Web Penetrated)
'             *****This Index(8) item will be created only if the BottomWebCut doesnt exist (i.e if Index(3) Item Member conditional is False)
'             *****if BottomWebCut exist the next item Member(i.e Index(9)) will create the Flange Cut
'
' Index(9) - Flange Cut for Bottom Flange(if Exist) of Bounded Member
'            This item will be created only if the BottomWebCut exist (i.e if Index(3) Item Member conditional is True)
'            index(9) item member is Vice versa of Index(8)Item Member
' Index(10) - Top Flange Cut for Bottom Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(11) - Bottom Flange Cut for Bottom Flange(if Exist) ---> Applied if (Web Non Penetrated and if End User Needs(Depending on Answer to some Selector Questions))
' Index(12) - Creates a Child Inset brace(Top/Bottom) ---> Applied if answer to the question "BraceType" is set to other than None.
' Index(13) - Creates an AC between ChildInsetBrace and Bounded of the parent AC ---> Creates AC whenever the Insetbrace is created.
' Index(14) - Creates an AC between ChildInsetBrace and Bounding of the parent AC ---> Creates AC whenever the Insetbrace is created.
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
    'if any more item members are needed to be added please add only at the bottom
    '******************************************************************************
       
    'should be responsible for creation of Web Cut(Center Cut or Normal Cut)
    lDispId = 1
    sDispId = "Axis_MbrToEdge_WebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    'should be responsible for creation of Top Web Cut
    lDispId = 2
    sDispId = "Axis_MbrToEdge_TopWebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute2", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    'should be responsible for creation of Bottom Web Cut
    lDispId = 3
    sDispId = "Axis_MbrToEdge_BtmWebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute3", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    'The next Item Member should be responsible for creation of Flange Cut on Top
    '    --->Center Cut or
    '    --->Normal Cut(when TopWebCut doesn't exist)
    'The item Member will create the Cut only when Top Web Cut doesnt exist
    lDispId = 4
    sDispId = "Axis_MbrToEdge_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'The next Item Member should be responsible for creation of Flange Cut on Top
    '    --->Normal Cut(Only when TopWebCut exist)
    'The item Member will create the Cut only when Top Web Cut exist(because TopWebCut will act as Input to Flange Cut)
    lDispId = 5
    sDispId = "Axis_MbrToEdge_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Top Flange Cut for Top Flange
    lDispId = 6
    sDispId = "Axis_MbrToEdge_TopFlangeCutForTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutForTopFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Bottom Flange Cut for Top Flange
    lDispId = 7
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutForTopFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Flange Cut on Bottom(Ceneter Cut or Normal Cut)
    lDispId = 8
    sDispId = "Axis_MbrToEdge_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Flange Cut on Bottom(Ceneter Cut or Normal Cut)
    lDispId = 9
    sDispId = "Axis_MbrToEdge_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Top Flange Cut for Bottom Flange
    lDispId = 10
    sDispId = "Axis_MbrToEdge_TopFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutForBtmFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'should be responsible for creation of Bottom Flange Cut for Bottom Flange
    lDispId = 11
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutForBtmFlangeNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'new item members added to create Top & Bottom Brace - B-52232
     ' should be responsible to create a new member as a child(brace)
    lDispId = 12
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructTopBrace", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsTopBraceNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    
     ' should be responsible to create AC between new member(child brace), bounded object
    lDispId = 13
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBraceAC1", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsTopBraceAC1Needed"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
     ' should be responsible to create AC between new member(child brace), bounded object
    lDispId = 14
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBraceAC2", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsTopBraceAC2Needed"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
     ' should be responsible to create a new member as a child(brace)
    lDispId = 15
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBottomBrace", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsBottomBraceNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    
     ' should be responsible to create AC between new member(child brace), bounded object
    lDispId = 16
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBraceAC1", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsBottomBraceAC1Needed"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
     ' should be responsible to create AC between new member(child brace), bounded object
    lDispId = 17
    sDispId = "Axis_MbrToEdge_BtmFlangeCutForBtmFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBraceAC2", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional CUSTOMERID + "MbrAC.MbrACDefCM", "CM_IsBottomBraceAC2Needed"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
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

'*********************************************************************************************
' Method      : CM_IsWebCutNeeded
' Description : Conditional method to determine if web cut is needed
'
'********************************************************************************************
Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    bIsNeeded = True
    

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsTopWebCutNeeded
' Description : Conditional method to determine if a Top web cut is needed
'
'********************************************************************************************
Public Sub CM_IsTopWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopWebCutNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           WebEndCut, _
                                           False, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "TopShapeAtFace", _
                                           vbNullString)

    ' -----------------------------------------------------------
    ' Use a helper that checks if one is needed for a child brace
    ' -----------------------------------------------------------
    If Not bIsNeeded Then
        bIsNeeded = oHelper.IsOutsideWebCutNeededForChildBrace(pMemberDescription.CAO, False, "TopBraceType")
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
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmWebCutNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           WebEndCut, _
                                           True, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "BottomShapeAtFace", _
                                           vbNullString)

    ' -----------------------------------------------------------
    ' Use a helper that checks if one is needed for a child brace
    ' -----------------------------------------------------------
    If Not bIsNeeded Then
        bIsNeeded = oHelper.IsOutsideWebCutNeededForChildBrace(pMemberDescription.CAO, True, "BottomBraceType")
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
                         
                             
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases
    
    If HasTopFlange(oBoundedData.MemberPart) Then
        If IsTubularMember(oBoundingData.MemberPart) Then
            'Bounding member is tube
            If (Not bPenetratesWeb) Then
                'Flange penetrated case
                bIsNeeded = IsBoundingOrthogonalToBoundedFlg(oBoundingData, oBoundedData, JXSEC_TOP)
            Else
                'Web penetrated case
                bIsNeeded = True
            End If
        Else
          bIsNeeded = True
        End If
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
    
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutForTopFlangeNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           TopFlangeEndCut, _
                                           False, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "TopShapeAtFace", _
                                           vbNullString)
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
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutForTopFlangeNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           TopFlangeEndCut, _
                                           True, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "BottomShapeAtFace", _
                                           vbNullString)
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
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    bIsNeeded = False
       
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

    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
    
    Dim bIsConnBTWBraceAndBounded As Boolean, bIsConnBTWBraceAndBounding As Boolean
    IsCutOnInsetBrace pMemberDescription.CAO, bIsConnBTWBraceAndBounded, bIsConnBTWBraceAndBounding
    
    If bIsConnBTWBraceAndBounded Or bIsConnBTWBraceAndBounding Then
        bIsNeeded = False
        Exit Sub
    End If
    
    If HasBottomFlange(oBoundedData.MemberPart) Then
        If IsTubularMember(oBoundingData.MemberPart) Then
            'Bounding member is tube
            If (Not bPenetratesWeb) Then
                'Flange penetrated case
                bIsNeeded = IsBoundingOrthogonalToBoundedFlg(oBoundingData, oBoundedData, JXSEC_BOTTOM)
            Else
                'Web penetrated case
                bIsNeeded = True
            End If
        Else
          bIsNeeded = True
        End If
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
    
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutForBtmFlangeNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           BottomFlangeEndCut, _
                                           False, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "TopShapeAtFace", _
                                           vbNullString)
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
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutForBtmFlangeNeeded"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -----------------------------------------------------------------------------------------------
    ' Use a helper that checks various conditions based on the question answers and relative position
    ' -----------------------------------------------------------------------------------------------
    Dim oHelper As New MbrACDefCM
    
    bIsNeeded = oHelper.IsOutsideCutNeeded(pMemberDescription.CAO, _
                                           BottomFlangeEndCut, _
                                           True, _
                                           eBoundingCondition.FaceAndOutside, _
                                           "BottomShapeAtFace", _
                                           vbNullString)
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_FinalConstructAggregator
' Description :
'********************************************************************************************
Public Sub CM_FinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::CM_FinalConstructAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Final Construct of Assembly Connection"
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_MigrateAggregator
' Description :
'********************************************************************************************
Public Sub CM_MigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, _
                                pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigrateAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Mirgrating Assembly Connection Inputs"
    MigrateAssemblyConnection pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_MigrateEndCut
' Description
'********************************************************************************************
Public Sub CM_MigrateEndCut(pMemberDesc As IJDMemberDescription, _
                            pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigrateEndCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' The EndCut Inputs have been Mirgrated at same time with the Assembly Connection Inputs
    ' Therefore, there should be nothing to do here
    sMsg = "Mirgrating EndCut Inputs"

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

    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName

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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description
'********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Implements ISPSFACInputHelper
'Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
'Property Get ISPSFACInputHelper_ValidateObjects(inputObjs As SP3DStructInterfaces.IJElements, _
'                                               relationObjs As SP3DStructInterfaces.IJElements) _
'                                               As SP3DStructInterfaces.SPSFACInputHelperStatus
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'*************************************************************************
'Function
'ISPSFACInputHelper_UserAttributeMgmt
'***************************************************************************
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() _
                            As SP3DStructInterfaces.IJUserAttributeMgmt
    Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_UserAttributeMgmt"
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
End Property

'*************************************************************************
'Function
'ISPSFACInputHelper_ValidateObjects
'***************************************************************************
Private Property Get ISPSFACInputHelper_ValidateObjects( _
                            ByVal oInputObjs As SP3DStructInterfaces.IJElements, _
                            oRelationObjs As SP3DStructInterfaces.IJElements) _
                            As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_ValidateObjects"
    
    On Error GoTo ErrorHandler
    ISPSFACInputHelper_ValidateObjects = InputHelper_ValidateObjects(oInputObjs, _
                                                                     oRelationObjs)
    Exit Property
ErrorHandler:
    HandleError MODULE, METHOD
End Property

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Implements IJUserAttributeMgmt
'Function OnAttributeChange(pIJDAttrs As IJDAttributes, _
'                           CollAllDisplayedValues As Object, _
'                           pAttrToChange As IJAttributeDescriptor, _
'                           varNewAttrValue) As String
'
'Function OnPreCommit(pIJDAttrs As IJDAttributes, CollAllDisplayedValues As Object) As String
'
'Function OnPreLoad(pIJDAttrs As IJDAttributes, CollAllDisplayedValues As Object) As String
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnAttributeChange
'***************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange( _
                        ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                        ByVal CollAllDisplayedValues As Object, _
                        ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, _
                        ByVal varNewAttrValue As Variant) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnAttributeChange"

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreCommit
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit( _
                                ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreCommit"

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreLoad
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad( _
                                ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreLoad"

End Function




