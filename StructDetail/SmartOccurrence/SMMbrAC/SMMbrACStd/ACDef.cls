VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014-16, Intergraph Corporation.  All rights reserved.
'
'  File        : ACDef.cls
'
'  Description :
'
'  Author      : Alligators
'  Revision History
'    04/Nov.14    NK  Adding a new item for creating slots through CR-CP-262343  Create AC ladder rung penetrating support
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
'    03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
'    27/Feb/2015 MDT/hgunturu
'         CR-265236 Added the properties to some of the item members to update the web and flangecuts for box connections when one box connection is changed
'                   to other box AC and also when "Flip Primary and Secondary" question answer is changed.
'                   Modified methods- CMNeedtoCompute(),CMNeedToComputeForSecondaryAttributes(),CMNeedToComputeForPrimaryAttributes()
'                   Added Methods - CMNeedToComputeForSecondaryFlangecut()
'    24/Mar/15  MDT/RPK  TR-269306 Handrail placed on builtup member, cannot create assembly connection, errors,
'                    Added  checks related to design/Builtup member as the bounding member in the CMConstructWebCut, CMConstructFlangeCut method.
'    12/May/2015  GH CR-260982 - Added Properties descriptions to all the Web, Flange and Corner feature item members and updated CMNeedtoCompute()
'                               Added new method UpdateFeatureForSelectionChange() to update the features whenever
'                               the feature smart items are changed.
'    16/Jun/2015   svsmylav   CR-271082 for 'ToEdgecase' AC, 'CMConstructFlangeCut' is
'                          is updated for dispids 7 and 10 to create flange cut with proper bounding edge port id.
'    27/11/2015    hgunturu    DI-282754: Updated GetSelections() Method to improve performance of Standard ACs
'    30/11/2015  - dsmamidi    DI-284513  Character limit restricts definition of Standard ACs
'                                                               modified GetStandardACAttribute() method
'    6/Jan/2016    svsmylav   DI-273986: PC was missing for few on-member cases as bounding port was
'                             filled using 'ReverseMap'. Fix uses mapped bounding port id as key to get real port.
'    28/Apr/2016   svsmylav   TR-290235: For Miter AC PC was missing AC duirng 'FlipPrimaryAndSecondary' answer from Yes to No (Default).
'                             So modified two propeties viz. 'CMNeedToComputeForPrimaryAttributes' and 'CMNeedToComputeForSecondaryAttributes'
'                             Fix checks if AC is Miter or Split and update the web cut for the selection change.
'    15/June/2016    knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'    12/Aug/2016   svsmylav   TR-274056: create physical connection argument is corrected from 'strFilter' to 'strSelection'.
'*********************************************************************************************

Const m_sClassName As String = "ACDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"
Dim m_sSymbolSharePath As String

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

    Dim oPD As IJDPropertyDescriptions
    Set oPD = pAD
    
    oPD.AddProperty "CachedData", 1, "IJDAttributes", "CM_UpdateCache", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_BEFORE_SYMBOL_UPDATE
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
' Index:
' End cuts, web-penetrated cases
' 1 - Left (or single) Center Web Cut
' 2 - Left (or single) Top Web Cut
' 3 - Left (or single) Bottom Web Cut
' 4 - Right Center Web Cut (future use)
' 5 - Right Top Web Cut (future use)
' 6 - Right Bottom Web Cut (future use)
' 7 - Flange Cut (Top flange, center web input)
' 8 - Flange Cut (Top flange, top web input)
' 9 - Flange Cut (Top flange, bottom web input)(On-member case)
' 10 - Flange Cut (Bottom flange, center web input)
' 11 - Flange Cut (Bottom flange, bottom web input)
' 12 - Flange Cut (Bottom flange, top web input)(On-member case)
' End cuts, flange-penetrated cases
' 13 - First Web Cut
' 14 - Second Web Cut (future use)
' 15 - Top Flange Center Cut
' 16 - Top Flange Top Cut
' 17 - Top Flange Bottom Cut
' 18 - Bottom Flange Center Cut
' 19 - Bottom Flange Top Cut
' 20 - Bottom Flange Bottom Cut
' Corner features, penetrated
' 21 - TopFaceTopCorner_LT
' 22 - TopFaceBottomCorner_LT
' 23 - TopEdgeOutsideCorner_LT
' 24 - TopEdgeInsideCorner_LT
' 25 - FaceTopCorner_LT
' 26 - FaceBottomCorner_LT
' 27 - BottomEdgeInsideCorner_LT
' 28 - BottomEdgeOutsideCorner_LT
' 29 - BottomFaceTopCorner_LT
' 30 - BottomFaceBottomCorner_LT
' 31 - TopFaceTopCorner_RB
' 32 - TopFaceBottomCorner_RB
' 33 - TopEdgeOutsideCorner_RB
' 36 - FaceBottomCorner_RB
' 37 - BottomEdgeInsideCorner_RB
' 38 - BottomEdgeOutsideCorner_RB
' 39 - BottomFaceTopCorner_RB
' 40 - BottomFaceBottomCorner_RB
' Corner Features, non-penetrated
' 41 - NPTopFlangeLeftInsideCorner
' 42 - NPTopFlangeRightInsideCorner
' 43 - NPBtmFlangeLeftInsideCorner
' 44 - NPBtmFlangeRightInsideCorner
' 45 - NPLeftWebTopInsideCorner
' 46 - NPLeftWebBtmInsideCorner
' 47 - NPRightWebTopInsideCorner
' 48 - NPRightWebBtmInsideCorner
' Split and Miter Attributes
' 49 - PrimaryLeftWeb
' 50 - PrimaryRightWeb
' 51 - PrimaryTopFlange
' 52 - PrimaryBottomFlange
' 53 - SecondaryLeftWeb
' 54 - SecondaryRightWeb
' 55 - SecondaryTopFlange
' 56 - SecondaryBottomFlange
' 57 - TopPC_LT
' 58 - TopPC_RB
' 59 - BtmPC_LT
' 60 - BtmPC_RB
' 61 - Bearing Plate
' 62 - Slot Feature
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sDispId As String
    
    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    
    ' ------------------------------
    ' End Cuts, web-penetrated cases
    ' ------------------------------
    ' Center web cut (left or single)
    lDispId = 1
    sDispId = "StdAC_CenterWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeCenterWebCut", 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top web cut (left or single)
    lDispId = 2
    sDispId = "StdAC_TopWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom web cut (left or single)
    lDispId = 3
    sDispId = "StdAC_BtmWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Right Center web cut
    lDispId = 4
    sDispId = "StdAC_RightCenterWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Right Top web cut
    lDispId = 5
    sDispId = "StdAC_RightTopWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Right Bottom web cut
    lDispId = 6
    sDispId = "StdAC_RightBtmWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, center web cut input
    lDispId = 7
    sDispId = "StdAC_TopFlangeCut_CenterWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeTopFlangeCut", 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, top web cut input
    lDispId = 8
    sDispId = "StdAC_TopFlangeCut_TopWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, bottom web cut input
    lDispId = 9
    sDispId = "StdAC_TopFlangeCut_BottomWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, center web cut input
    lDispId = 10
    sDispId = "StdAC_BtmFlangeCut_CenterWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeBtmFlangeCut", 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, bottom web cut input
    lDispId = 11
    sDispId = "StdAC_BtmFlangeCut_BottomWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, top web cut input
    lDispId = 12
    sDispId = "StdAC_BtmFlangeCut_TopWebCutInput_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' ---------------------------------
    ' End cuts, flange-penetrated cases
    ' ---------------------------------
    ' First web cut
    lDispId = 13
    sDispId = "StdAC_FirstWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' First web cut
    lDispId = 14
    sDispId = "StdAC_SecondWebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, center cut
    lDispId = 15
    sDispId = "StdAC_TopFlangeCut_CenterCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, top cut
    lDispId = 16
    sDispId = "StdAC_TopFlangeCut_TopCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Top flange cut, bottom cut
    lDispId = 17
    sDispId = "StdAC_TopFlangeCut_BottomCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, center cut
    lDispId = 18
    sDispId = "StdAC_BtmFlangeCut_CenterCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, top web cut input
    lDispId = 19
    sDispId = "StdAC_BtmFlangeCut_TopCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Bottom flange cut, bottom cut
    lDispId = 20
    sDispId = "StdAC_BtmFlangeCut_BottomCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' ---------------------------
    ' Corner features, penetrated
    ' ---------------------------
    ' TopFaceTopCorner_LT
    lDispId = 21
    sDispId = "StdAC_TopFaceTopCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' TopFaceBottomCorner_LT
    lDispId = 22
    sDispId = "StdAC_TopFaceBottomCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' TopEdgeOutsideCorner_LT
    lDispId = 23
    sDispId = "StdAC_TopEdgeOutsideCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' TopEdgeInsideCorner_LT
    lDispId = 24
    sDispId = "StdAC_TopEdgeInsideCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' FaceTopCorner_LT
    lDispId = 25
    sDispId = "StdAC_FaceTopCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' FaceBottomCorner_LT
    lDispId = 26
    sDispId = "StdAC_FaceBottomCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomEdgeInsideCorner_LT
    lDispId = 27
    sDispId = "StdAC_BottomEdgeInsideCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomEdgeOutsideCorner_LT
    lDispId = 28
    sDispId = "StdAC_BottomEdgeOutsideCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFaceTopCorner_LT
    lDispId = 29
    sDispId = "StdAC_BottomFaceTopCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFaceBottomCorner_LT
    lDispId = 30
    sDispId = "StdAC_BottomFaceBottomCorner_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' TopFaceTopCorner_RB
    lDispId = 31
    sDispId = "StdAC_TopFaceTopCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' TopFaceBottomCorner_RB
    lDispId = 32
    sDispId = "StdAC_TopFaceBottomCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' TopEdgeOutsideCorner_RB
    lDispId = 33
    sDispId = "StdAC_TopEdgeOutsideCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' TopEdgeInsideCorner_RB
    lDispId = 34
    sDispId = "StdAC_TopEdgeInsideCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    ' FaceTopCorner_RB
    lDispId = 35
    sDispId = "StdAC_FaceTopCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' FaceBottomCorner_RB
    lDispId = 36
    sDispId = "StdAC_FaceBottomCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomEdgeInsideCorner_RB
    lDispId = 37
    sDispId = "StdAC_BottomEdgeInsideCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomEdgeOutsideCorner_RB
    lDispId = 38
    sDispId = "StdAC_BottomEdgeOutsideCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFaceTopCorner_RB
    lDispId = 39
    sDispId = "StdAC_BottomFaceTopCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFaceBottomCorner_RB
    lDispId = 40
    sDispId = "StdAC_BottomFaceBottomCorner_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' -------------------------------
    ' Corner Features, non-penetrated
    ' -------------------------------
    ' TopFlangeLeftCorner
    lDispId = 41
    sDispId = "StdAC_TopFlangeLeftCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' TopFlangeRightCorner
    lDispId = 42
    sDispId = "StdAC_TopFlangeRightCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFlangeLeftCorner
    lDispId = 43
    sDispId = "StdAC_BottomFlangeLeftCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' BottomFlangeRightCorner
    lDispId = 44
    sDispId = "StdAC_BottomFlangeRightCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' LeftWebTopCorner
    lDispId = 45
    sDispId = "StdAC_LeftWebTopCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' LeftWebBottomCorner
    lDispId = 46
    sDispId = "StdAC_LeftWebBottomCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' RightWebTopCorner
    lDispId = 47
    sDispId = "StdAC_RightWebTopCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' RightWebBottomCorner
    lDispId = 48
    sDispId = "StdAC_RightWebBottomCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoCompute", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' --------------------------
    ' Split and Miter Attributes
    ' --------------------------
    ' PrimaryLeftWeb
    lDispId = 49
    sDispId = "StdAC_PrimaryLeftWeb_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePrimaryLeftWeb", 1, "IJSmartOccurrence", "CMNeedToComputeForPrimaryAttributes", imsCOOKIE_ID_USS_LIB
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    ' PrimaryRightWeb
    lDispId = 50
    sDispId = "StdAC_PrimaryRightWeb_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    ' PrimaryTopFlange
    lDispId = 51
    sDispId = "StdAC_PrimaryTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    ' PrimaryBottomFlange
    lDispId = 52
    sDispId = "StdAC_PrimaryBottomFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    ' SecondaryLeftWeb
    lDispId = 53
    sDispId = "StdAC_SecondaryLeftWeb_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeSecondaryLeftWeb", 1, "IJSmartOccurrence", "CMNeedToComputeForSecondaryAttributes", imsCOOKIE_ID_USS_LIB
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    ' SecondaryRightWeb
    lDispId = 54
    sDispId = "StdAC_SecondaryRightWeb_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    ' SecondaryTopFlange
    lDispId = 55
    sDispId = "StdAC_SecondaryTopFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeSecondaryTopFlange", 1, "IJSmartOccurrence", "CMNeedToComputeForSecondaryFlangecut", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' SecondaryBottomFlange
    lDispId = 56
    sDispId = "StdAC_SecondaryBottomFlange_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeSecondaryBottomFlange", 1, "IJSmartOccurrence", "CMNeedToComputeForSecondaryFlangecut", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Axis_MbrToEdge_SeatedPC_
    'Top SeatedPC: Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
    lDispId = 57
    sDispId = "Axis_MbrToEdge_SeatedPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysicalConnection", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'Axis_MbrToEdge_SeatedPC_
    'Top SeatedPC: Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrated
    lDispId = 58
    sDispId = "Axis_MbrToEdge_SeatedPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysicalConnection", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'Axis_MbrToEdge_SeatedPC_
    'Bottom SeatedPC: Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
    lDispId = 59
    sDispId = "Axis_MbrToEdge_SeatedPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysicalConnection", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    'Axis_MbrToEdge_SeatedPC_
    'Bottom SeatedPC: Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrated
    lDispId = 60
    sDispId = "Axis_MbrToEdge_SeatedPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysicalConnection", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing

     ' Bearing Plate
    lDispId = 61
    sDispId = "StdAC_BearingPlate_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBearingPlate", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeBearingPlate", 1, "IJSmartOccurrence", "CMNeedToCompute", imsCOOKIE_ID_USS_LIB
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    ' Slot feature
    lDispId = 62
    sDispId = "StdAC_SlotCutout_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructSlot", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsItemNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeSlot", 1, "IJSmartOccurrence", "CMNeedToCompute", imsCOOKIE_ID_USS_LIB
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************
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
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_UserAttributeMgmt"
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
End Property

'*************************************************************************
'Function
'ISPSFACInputHelper_ValidateObjects
'***************************************************************************
Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal oInputObjs As SP3DStructInterfaces.IJElements, _
                                                        oRelationObjs As SP3DStructInterfaces.IJElements) _
                                                        As SP3DStructInterfaces.SPSFACInputHelperStatus
Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_ValidateObjects"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lCount As Long
    Dim oPortCol As IJElements
    
    Dim ePortIdx As SPSMemberAxisPortIndex
    
    Dim oInputObj1 As Object
    Dim oInputObj2 As Object
    Dim oFrmConn As ISPSFrameConnection
    Dim oSupportedPort As ISPSSplitAxisPort
    Dim oSupportingPort As IJPort
    
    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Set oPortCol = New JObjectCollection
    'filter out ports to portCol
    For lCount = 1 To oInputObjs.Count
        If oInputObjs.Item(lCount) Is Nothing Then
        ElseIf TypeOf oInputObjs.Item(lCount) Is IJPort Then
            oPortCol.Add oInputObjs.Item(lCount)
        End If
    Next lCount
    
    '  make sure there are only two ports
    If oPortCol.Count <> 2 Then
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
        Exit Property
    End If
    ' Currently only supporting:
    '   Member End to Member Along Axis cases
    '   Member End to Member End cases (Split by Point)
    '   Member Along Axis to Member Along Axis cases (on Member case)
    ' First Port MUST BE ISPSSplitAxisEndPort or ISPSSplitAxisAlongPort
    Set oInputObj1 = oPortCol.Item(1)
    Set oInputObj2 = oPortCol.Item(2)
    
    'When Both ports are Axis Along
    If TypeOf oInputObj1 Is ISPSSplitAxisAlongPort And TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
        
        GetLogicalBoundedAndBounding_ForBothAlongPorts oInputObj1, oInputObj2, oSupportedPort, oSupportingPort
       
        Set oRelationObjs = New JObjectCollection
        oRelationObjs.Add oSupportedPort
        oRelationObjs.Add oSupportingPort
        Exit Property
    End If
    
    If Not TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
        Dim oTemp As Object
        Set oTemp = oInputObj1
        Set oInputObj1 = oInputObj2
        Set oInputObj2 = oTemp
    End If
    
    If Not TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        Exit Property
    End If
    
    ' Second Port MUST BE ISPSSplitAxisEndPort or ISPSSplitAxisAlongPort
    '   ISPSSplitAxisEndPort is expected for Member Split By Point cases
    '   ISPSSplitAxisAlongPort is expected for Bounded cases
    
    If Not TypeOf oInputObj2 Is ISPSSplitAxisEndPort Then
        If Not TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
            If Not TypeOf oInputObj2 Is IJPort Then
                ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
                Exit Property
            Else
                Dim oTempPort As IJPort
                Set oTempPort = oInputObj2
                
                If Not (TypeOf oTempPort.Connectable Is IJPlate Or TypeOf oTempPort.Connectable Is SPSSlabEntity Or TypeOf oTempPort.Connectable Is SPSWallPart) Then
                      ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
                      Exit Property
                End If
            End If
        End If
    End If
            
    ' Inputs are as expected: this is a valid Assembly Connection
    Set oSupportedPort = oInputObj1
    Set oSupportingPort = oInputObj2
                
    If oRelationObjs Is Nothing Then
        Set oRelationObjs = New JObjectCollection
    End If
    
    oRelationObjs.Clear
    oRelationObjs.Add oSupportedPort
    oRelationObjs.Add oSupportingPort
    
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
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                                       ByVal CollAllDisplayedValues As Object, _
                                                       ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, _
                                                       ByVal varNewAttrValue As Variant) As String
Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnAttributeChange"

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreCommit
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                                 ByVal CollAllDisplayedValues As Object) As String
Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreCommit"

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreLoad
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                               ByVal CollAllDisplayedValues As Object) As String
Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreLoad"

End Function

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************

'*********************************************************************************************
' Method      : CM_IsItemNeeded
' Description : Conditional method to determine if a Top web cut is needed
'
'********************************************************************************************
Public Sub CM_IsItemNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsItemNeeded"
    
    On Error GoTo ErrorHandler
    
    ' First we must copy the catalog attributes to the occurrence
    ' John Schwartz says this should probably not be necessary
    ' No TR filed because we're not certain
    Dim oCM As New CustomMethods
    oCM.CopyOccurrenceAttributes oMD, "IJUAStdACDefinition"
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    If Not GetSelection(oMD) = vbNullString Then
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsTopFlangeCutNeeded
' Description : Conditional method to determine if a particular flange cut is needed
'
'********************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutNeeded"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -------------------
    ' Get the input ports
    ' -------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    GetAssemblyConnectionInputs oMD.CAO, oBoundedPort, oBoundingPort
    
    ' ---------------------
    ' Skip if no top flange
    ' ---------------------
    If Not HasTopFlange(oBoundedPort.Connectable) Then
        Exit Sub
    End If
    
    ' ----------------------------------------------
    ' Skip if no selection identified by the AC item
    ' ----------------------------------------------
    If GetSelection(oMD) = vbNullString Then
        Exit Sub
    End If
    
    ' --------------------------------
    ' Get a handle for the other items
    ' --------------------------------
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oMD.CAO

    ' -------------------------------------------------------------------
    ' If this is for the flange cut that uses the center web cut as input
    ' -------------------------------------------------------------------
    If oMD.dispid = 7 Then
        ' -------------------------------------------------------------------------------------
        ' If the top web cut does not exist, it is needed unless the center also does not exist
        ' -------------------------------------------------------------------------------------
        If (oMemberObjects.ItemByDispid(2) Is Nothing) And (Not oMemberObjects.ItemByDispid(1) Is Nothing) Then
            bIsNeeded = True
        End If
    ' ----------------------------------------------------------------
    ' If this is for the flange cut that uses the top web cut as input
    ' ----------------------------------------------------------------
    ElseIf oMD.dispid = 8 Then
        ' ---------------------------------------
        ' If the top web cut exists, it is needed
        ' ---------------------------------------
        If Not oMemberObjects.ItemByDispid(2) Is Nothing Then
            bIsNeeded = True
        End If
    ' -------------------------------------------------------------------
    ' If this is for the flange cut that uses the bottom web cut as input
    ' -------------------------------------------------------------------
    ElseIf oMD.dispid = 9 Then
        ' --------------------------------------------------------------
        ' If both the top and center web cuts are not used, it is needed
        ' --------------------------------------------------------------
        If (oMemberObjects.ItemByDispid(2) Is Nothing) And (oMemberObjects.ItemByDispid(1) Is Nothing) Then
            bIsNeeded = True
        End If
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsBtmFlangeCutNeeded
' Description : Conditional method to determine if a particular flange cut is needed
'
'********************************************************************************************
Public Sub CM_IsBtmFlangeCutNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutNeeded"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' -------------------
    ' Get the input ports
    ' -------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    GetAssemblyConnectionInputs oMD.CAO, oBoundedPort, oBoundingPort
    
    ' ------------------------
    ' Skip if no bottom flange
    ' ------------------------
    If Not HasBottomFlange(oBoundedPort.Connectable) Then
        Exit Sub
    End If
    
    ' ----------------------------------------------
    ' Skip if no selection identified by the AC item
    ' ----------------------------------------------
    If GetSelection(oMD) = vbNullString Then
        Exit Sub
    End If
    
    ' --------------------------------
    ' Get a handle for the other items
    ' --------------------------------
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oMD.CAO

    ' -------------------------------------------------------------------
    ' If this is for the flange cut that uses the center web cut as input
    ' -------------------------------------------------------------------
    If oMD.dispid = 10 Then
        ' ----------------------------------------------------------------------------------------
        ' If the bottom web cut does not exist, it is needed unless the center also does not exist
        ' ----------------------------------------------------------------------------------------
        If (oMemberObjects.ItemByDispid(3) Is Nothing) And (Not oMemberObjects.ItemByDispid(1) Is Nothing) Then
            bIsNeeded = True
        End If
    ' -------------------------------------------------------------------
    ' If this is for the flange cut that uses the bottom web cut as input
    ' -------------------------------------------------------------------
    ElseIf oMD.dispid = 11 Then
        ' ------------------------------------------
        ' If the bottom web cut exists, it is needed
        ' ------------------------------------------
        If Not oMemberObjects.ItemByDispid(3) Is Nothing Then
            bIsNeeded = True
        End If
    ' ----------------------------------------------------------------
    ' If this is for the flange cut that uses the top web cut as input
    ' ----------------------------------------------------------------
    ElseIf oMD.dispid = 12 Then
        ' -----------------------------------------------------------------
        ' If both the bottom and center web cuts are not used, it is needed
        ' -----------------------------------------------------------------
        If (oMemberObjects.ItemByDispid(3) Is Nothing) And (oMemberObjects.ItemByDispid(1) Is Nothing) Then
            bIsNeeded = True
        End If
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
' Method:
' CMConstructWebCut
'
' Description:
' Custom method for creating a web cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructWebCut(ByVal oMD As IJDMemberDescription, _
                                  ByVal oResourceManager As IUnknown, _
                                  ByRef oObject As Object)
    
    Const METHOD = m_DefinitionProgid & "::CMConstructWebCut"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating WebCut ...oMD.index = " & Str(oMD.Index)
     
    Dim strSelection As String
    strSelection = GetSelection(oMD)
     
    Dim oMbrACDefCM As Object
    
    Dim eCase As eMemberBoundingCase
    eCase = GetMemberBoundingCase(oMD.CAO, , , False)
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim oBdngPort As IJPort
    Dim oAppConnection As IJAppConnection
    Dim lStatus As Long
        
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = oMD.CAO
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBdngPort
        
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBdngPort, oBoundedPort)
    
    Dim lSectionAlias As Long
    Dim oBoundingObject As Object
    Dim oBoundingPort As Object
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    
    If (Not bPenetratesWeb) And (eCase = OnMemberBottom Or eCase = OnMemberTop) And Not (IsTubularMember(oBoundedPort.Connectable) And IsTubularMember(oBdngPort.Connectable)) Then
        Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBdngPort, oBoundedPort, lSectionAlias)
        Set oBoundingObject = oBdngPort.Connectable
                
        If eCase = OnMemberTop Then
        
            Select Case lSectionAlias
                Case 6, 8, 9, 12, 13, 19
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oBoundingObject, oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_LEFT)))
                Case Else
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oBoundingObject, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
            End Select
            
        ElseIf eCase = OnMemberBottom Then
        
            Select Case lSectionAlias
                Case 7, 10, 11, 12, 13, 19
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oBoundingObject, oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_LEFT)))
                Case Else
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oBoundingObject, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
            End Select
        End If
    End If

    Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID + "MbrAC.MbrACDefCM")
    
    If oMD.dispid < 49 Then
        oMbrACDefCM.CreateAxisWebCut oMD, oResourceManager, strSelection, oObject, oBoundingPort
    ElseIf oMD.dispid < 53 Then
        ' -------------------
        ' Get the input ports
        ' -------------------
        Dim oSupportedPort As IJPort
        Dim oSupportingPort As IJPort
    
        GetAssemblyConnectionInputs oMD.CAO, oSupportedPort, oSupportingPort

        Dim oSD_Webcut As New StructDetailObjects.WebCut
        If (TypeOf oSupportingPort.Connectable Is IJPlate) Or (TypeOf oSupportingPort.Connectable Is SPSSlabEntity) _
                                                        Or (TypeOf oSupportingPort.Connectable Is SPSWallPart) Then
        
            oSD_Webcut.Create oResourceManager, oSupportingPort, oSupportedPort, strSelection, oMD.CAO
            Set oObject = oSD_Webcut.object
            
        '** Checking for the builtup member*******
        ElseIf TypeOf oSupportingPort.Connectable Is ISPSDesignedMember Then
            Dim oBounded As ISPSMemberPartPrismatic
            Dim oBoundedStart As IJPoint
            Dim oBoundedEnd As IJPoint
            Dim oBoundedRefPos As New DPosition
            Dim oBoundingBU As ISPSDesignedMember
            Set oBoundingBU = oSupportingPort.Connectable
            Dim dx As Double
            Dim dy As Double
            Dim dz As Double
            
            'getting the bounded far end as ref postion********
            If TypeOf oSupportedPort Is ISPSSplitAxisPort Then
                Dim oBoundedMemberPort As ISPSSplitAxisPort
                Set oBoundedMemberPort = oSupportedPort
                Dim oBoundedMemberPart As ISPSMemberPartCommon
                Set oBoundedMemberPart = oSupportedPort.Connectable
                If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
                    Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
                        oBoundedEnd.GetPoint dx, dy, dz
                        oBoundedRefPos.Set dx, dy, dz
                    ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
                        Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
                        oBoundedStart.GetPoint dx, dy, dz
                        oBoundedRefPos.Set dx, dy, dz
               End If
                 GetNearestboundingBUPort oBoundedRefPos, oBoundedPort, oBoundingBU, oSupportingPort
            End If
            
            oSD_Webcut.Create oResourceManager, oSupportingPort, oSupportedPort, strSelection, oMD.CAO
            Set oObject = oSD_Webcut.object
        Else
            oMbrACDefCM.CM_CreateEndToEndWebCut oMD, oResourceManager, strSelection, True, False, 1, oObject
        End If
    Else
        If IsBoxCut(oMD.CAO) Then
            oMbrACDefCM.CM_CreateEndToEndWebCut oMD, oResourceManager, strSelection, True, True, 2, oObject
        Else
            oMbrACDefCM.CM_CreateEndToEndWebCut oMD, oResourceManager, strSelection, True, False, 2, oObject
        End If
        
    End If
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Set oSDOWebCut.object = oObject
    
    ' ---------------------------------------------------------------------------------
    ' In order to place corner features the web cuts must already be attached to the AE
    ' ---------------------------------------------------------------------------------
    Dim oStructProfilePart As IJStructProfilePart
    
    If TypeOf oSDOWebCut.Bounded Is IJStructProfilePart Then
        
        Set oStructProfilePart = oSDOWebCut.Bounded
        
        oStructProfilePart.PlaceEndCutFeature oObject, oSDOWebCut.BoundedPort, eEndCutTypes.WebCut
    
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' *******************************************************************************************
' Method:
' CMConstructFlangeCut
'
' Description:
' Custom method for creating a Flange cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructFlangeCut(ByVal oMD As IJDMemberDescription, _
                                     ByVal oResourceManager As IUnknown, _
                                     ByRef oObject As Object)
    
    Const METHOD = m_DefinitionProgid & "::CMConstructFlangeCut"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating FlangeCut ...oMD.index = " & Str(oMD.Index)
    
    ' -------------------
    ' Get the input ports
    ' -------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    GetAssemblyConnectionInputs oMD.CAO, oBoundedPort, oBoundingPort
    
    ' -----------------------------------
    ' Get the web cut to be used as input
    ' -----------------------------------
    Dim oWebCut As Object
    Dim oMemberObjects As IJDMemberObjects
    
    Set oMemberObjects = oMD.CAO

    ' ---------------------------------------------------------------
    ' Identify the web cut that should be the input to the flange cut
    ' ---------------------------------------------------------------
    Select Case oMD.dispid
        ' Flange cuts that use web center cuts as input
        Case 7, 10
            Set oWebCut = oMemberObjects.ItemByDispid(1)
        ' Flange cuts that use web top cuts as input
        Case 8, 12
            Set oWebCut = oMemberObjects.ItemByDispid(2)
        Case 9, 11
            Set oWebCut = oMemberObjects.ItemByDispid(3)
        Case 15, 16, 17, 18, 19, 20
            Set oWebCut = oMemberObjects.ItemByDispid(13)
        Case 51, 52
            Set oWebCut = oMemberObjects.ItemByDispid(49)
        Case 55, 56
            Set oWebCut = oMemberObjects.ItemByDispid(53)
    End Select
    
    If oWebCut Is Nothing Then
        GoTo ErrorHandler
    End If
        
    ' ---------------------
    ' Create the flange cut
    ' ---------------------
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long

    Dim eCase As eMemberBoundingCase
    eCase = GetMemberBoundingCase(oMD.CAO, , , False)
    Dim bCreateAxisFlangeCut As Boolean
    bCreateAxisFlangeCut = False 'Initialize
    If eCase = TopEdge Or eCase = BottomEdge Then
        bCreateAxisFlangeCut = True
        Dim strSelection As String
        Dim oMbrACDefCM As Object
        
        strSelection = GetSelection(oMD)
        Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID + "MbrAC.MbrACDefCM")
    End If
    
    If bCreateAxisFlangeCut And (oMD.dispid = 7 Or oMD.dispid = 10) Then
        'FlangeCuts that use web center cut as input
        If oMD.dispid = 7 Then
            oMbrACDefCM.CreateAxisFlangeCut oMD, oResourceManager, strSelection, oObject, False
        ElseIf oMD.dispid = 10 Then
            oMbrACDefCM.CreateAxisFlangeCut oMD, oResourceManager, strSelection, oObject, True
        End If
        'Initailize FlangeCut wrapper instance so that bounded part can be accessed (at bottom portion of this method)
        Set oSDO_FlangeCut.object = oObject
    ElseIf oMD.dispid < 51 Then
        oSDO_FlangeCut.Create oResourceManager, oBoundingPort, oBoundedPort, oWebCut, GetSelection(oMD), oMD.CAO
        Set oObject = oSDO_FlangeCut.object
    Else
        Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID + "MbrAC.MbrACDefCM")
    
        If oMD.dispid < 55 Then
            If (TypeOf oBoundingPort.Connectable Is IJPlate) Or (TypeOf oBoundingPort.Connectable Is SPSSlabEntity) Or _
                                                                (TypeOf oBoundingPort.Connectable Is SPSWallPart) Then
                oSDO_FlangeCut.Create oResourceManager, oBoundingPort, oBoundedPort, oWebCut, GetSelection(oMD), oMD.CAO
                Set oObject = oSDO_FlangeCut.object
                
            '** Checking for the builtup member*******
            
            ElseIf TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
                Dim oBounded As ISPSMemberPartPrismatic
                Dim oBoundingBU As ISPSDesignedMember
                Set oBoundingBU = oBoundingPort.Connectable
                Dim oBoundingMbrpart As ISPSMemberPartCommon
                Set oBoundingMbrpart = oBoundingPort.Connectable
              
                Dim oBoundedStart As IJPoint
                Dim oBoundedEnd As IJPoint
                Dim oBoundedRefPos As New DPosition
                
                Dim dx As Double
                Dim dy As Double
                Dim dz As Double
                
                'getting the bounded far end as ref postion
                
                If TypeOf oBoundedPort Is ISPSSplitAxisPort Then
                    Dim oBoundedMemberPort As ISPSSplitAxisPort
                    Set oBoundedMemberPort = oBoundedPort
                    Dim oBoundedMemberPart As ISPSMemberPartCommon
                    Set oBoundedMemberPart = oBoundedPort.Connectable
                    If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
                       Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
                                            oBoundedEnd.GetPoint dx, dy, dz
                                            oBoundedRefPos.Set dx, dy, dz
                    ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
                        Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
                                            oBoundedStart.GetPoint dx, dy, dz
                                            oBoundedRefPos.Set dx, dy, dz
                    End If
                    GetNearestboundingBUPort oBoundedRefPos, oBoundedPort, oBoundingBU, oBoundingPort
                    
                End If
                oSDO_FlangeCut.Create oResourceManager, oBoundingPort, oBoundedPort, oWebCut, GetSelection(oMD), oMD.CAO
                Set oObject = oSDO_FlangeCut.object
            Else
                oMbrACDefCM.CM_CreateEndToEndFlangeCut oMD, oResourceManager, GetSelection(oMD), True, False, 1, oObject
            End If
        Else
            If IsBoxCut(oMD.CAO) Then
                CM_CreateEndToEndFlangeCut oMD, oResourceManager, GetSelection(oMD), True, True, 2, oWebCut, oObject
            Else
                oMbrACDefCM.CM_CreateEndToEndFlangeCut oMD, oResourceManager, GetSelection(oMD), True, False, 2, oObject
            End If
        End If
        
        Set oSDO_FlangeCut.object = oObject
    End If
    
    ' ---------------------------------------------------------------------------------
    ' In order to place corner features the end cuts must already be attached to the AE
    ' ---------------------------------------------------------------------------------
    Dim oStructProfilePart As IJStructProfilePart
    
    If TypeOf oSDO_FlangeCut.Bounded Is IJStructProfilePart Then
        
        Set oStructProfilePart = oSDO_FlangeCut.Bounded
    
        Select Case oMD.dispid
            ' Top flange cuts
            Case 7, 8, 9, 15, 16, 17, 51, 55
                oStructProfilePart.PlaceEndCutFeature oObject, oSDO_FlangeCut.BoundedPort, eEndCutTypes.FlangeCutTop
            ' Bottom flange cuts
            Case 10, 11, 12, 18, 19, 20, 52, 56
                oStructProfilePart.PlaceEndCutFeature oObject, oSDO_FlangeCut.BoundedPort, eEndCutTypes.FlangeCutBottom
        End Select

    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' *******************************************************************************************
' Method:
' CMCopyQuestionsToFlangeCut
'
' Description: It is necessary to tell the flange cut if it is to be applied to the top or
'              bottom flange
' *******************************************************************************************
Public Sub CMCopyQuestionsToFlangeCut(oMD As IJDMemberDescription)
    
    Const METHOD = m_DefinitionProgid & "::CMCopyQuestionsToFlangeCut"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    sMsg = "Set 'bottom flange' question"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = oMD
    
    Select Case oMD.dispid
        Case 7, 8, 9, 15, 16, 17, 51, 55
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrACStd.FlangeCutSel", "BottomFlange", "No"
        Case 10, 11, 12, 18, 19, 20, 52, 56
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrACStd.FlangeCutSel", "BottomFlange", "Yes"
    End Select

    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : GetSelection
' Description :
'
'********************************************************************************************
Private Function GetSelection(oMD As IJDMemberDescription) As String
    
    Const METHOD = m_DefinitionProgid & "::GetSelection"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    ' -----------------------------
    ' Check if connected end-to-end
    ' -----------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = oMD.CAO
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
     
    Dim oBoundedPart As Object
    Set oBoundedPart = oBoundedPort.Connectable
    
    Dim lDispId As Long
    lDispId = oMD.dispid
        
    'Item members 57,58,59 and 60 are to related to Seated PC between bounded and bounding for On Member case
    If lDispId >= 57 And lDispId <= 60 Then
    
        Dim cWLOrTop As ConnectedEdgeInfo
        Dim cWROrBtm As ConnectedEdgeInfo
        Dim cTFIOrFL As ConnectedEdgeInfo
        Dim cBFIOrFR As ConnectedEdgeInfo
        
        'Get Bounding Case
        Dim eCase As eMemberBoundingCase
        eCase = GetMemberBoundingCase(oAppConnection, , , False)
        
        If Not eCase = OnTubeMember And Not IsTubularMember(oBoundingPort.Connectable) Then
            GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, cWLOrTop, cWROrBtm, cTFIOrFL, cBFIOrFR
        End If
            
        Dim bHastopFlange As Boolean
        Dim bHasBottomFlange As Boolean
        Dim bTFL As Boolean
        Dim bTFR As Boolean
        Dim bBFR As Boolean
        Dim bBFL As Boolean
        
        'Get Cross section Flanges data
        bHastopFlange = HasTopFlange(oBoundedPart, bTFL, bTFR)
        bHasBottomFlange = HasBottomFlange(oBoundedPart, bBFL, bBFR)
                
        Dim bRightWeb As Boolean
        bRightWeb = HasRightWeb(oBoundedPart)
    End If
    
    ' ---------------------------------------------------------------------
    ' Determine if CaseA or CaseB, as described in the bulkload spreadsheet
    ' ---------------------------------------------------------------------
    Dim oCustomMethods As New CustomMethods
    Dim isCaseA As Boolean
    Dim strCaseSuffix As String
    isCaseA = oCustomMethods.IsBoundingCaseA(oAppConnection, strCaseSuffix)
    
    ' -------------------------------------
    ' Determine if web or flange penetrated
    ' -------------------------------------

    Dim bPenetratesWeb As Boolean
    GetEdgeMap oAppConnection, oBoundingPort, oBoundedPort, , bPenetratesWeb
    ' -------------------------------------
    ' Center web cut for single or left web
    ' -------------------------------------
    ' Items 1-12 are web-penetrated only
    Select Case lDispId
    
    Case 1
        If bPenetratesWeb Then
            GetSelection = GetStandardACAttribute(oAppConnection, "LeftWebCenterCut" & strCaseSuffix)
        End If
    ' ------------------------------
    ' Top cut for single or left web
    ' ------------------------------
    Case 2
        If bPenetratesWeb Then
            GetSelection = GetStandardACAttribute(oAppConnection, "LeftWebTopCut" & strCaseSuffix)
        End If
    ' ---------------------------------
    ' Bottom cut for single or left web
    ' ---------------------------------
    Case 3
        If bPenetratesWeb Then
            GetSelection = GetStandardACAttribute(oAppConnection, "LeftWebBtmCut" & strCaseSuffix)
        End If
    ' ----------------------------
    ' Center web cut for right web
    ' ----------------------------
    ' Ignore the method ConsiderRightWeb when proper rule changes were made to take Rightweb into consideration.
    Case 4
        If bPenetratesWeb And HasRightWeb(oBoundedPart) And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "RightWebCenterCut" & strCaseSuffix)
        End If
    ' -------------------------
    ' Top web cut for right web
    ' -------------------------
    Case 5
        If bPenetratesWeb And HasRightWeb(oBoundedPart) And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "RightWebTopCut" & strCaseSuffix)
        End If
    ' ------------------------
    ' Bottom cut for right web
    ' ------------------------
    Case 6
        If bPenetratesWeb And HasRightWeb(oBoundedPart) And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "RightWebBtmCut" & strCaseSuffix)
        End If
    ' ------------------------------------------------------------
    ' For the top flange cut that uses the center web cut as input
    ' ------------------------------------------------------------
    Case 7
        If bPenetratesWeb And HasTopFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ---------------------------------------------------------
    ' For the top flange cut that uses the top web cut as input
    ' --------------------------------------------------------
    Case 8
        If bPenetratesWeb And HasTopFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ------------------------------------------------------------
    ' For the top flange cut that uses the bottom web cut as input
    ' ------------------------------------------------------------
    Case 9
        If bPenetratesWeb And HasTopFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ---------------------------------------------------------------
    ' For the bottom flange cut that uses the center web cut as input
    ' ---------------------------------------------------------------
    Case 10
        If bPenetratesWeb And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ---------------------------------------------------------------
    ' For the bottom flange cut that uses the bottom web cut as input
    ' ---------------------------------------------------------------
    Case 11
        If bPenetratesWeb And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ------------------------------------------------------------
    ' For the bottom flange cut that uses the top web cut as input
    ' ------------------------------------------------------------
    Case 12
        If bPenetratesWeb And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetNonPenFlangeCutSelection(oMD)
        End If
    ' ----------------------
    ' For single or left web
    ' ----------------------
    ' Items 13-20 are flange-penetrated only
    Case 13
        If Not bPenetratesWeb Then
            GetSelection = GetStandardACAttribute(oAppConnection, "NPLeftWebCut" & strCaseSuffix)
        End If
    ' -------------
    ' For right web
    ' -------------
    Case 14
        ' --------------------
        ' Skip if no right web
        ' --------------------
        If (Not bPenetratesWeb) And HasRightWeb(oBoundedPart) And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "NPRightWebCut" & strCaseSuffix)
        End If
    ' -------------------------
    ' For top flange center cut
    ' -------------------------
    Case 15
        ' ---------------------
        ' Skip if no top flange
        ' ---------------------
        If (Not bPenetratesWeb) And HasTopFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopFlangeCenterCut" & strCaseSuffix)
        End If
    ' ----------------------
    ' For top flange top cut
    ' ----------------------
    Case 16
        ' ---------------------
        ' Skip if no top flange
        ' ---------------------
        If (Not bPenetratesWeb) And HasTopFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopFlangeTopCut" & strCaseSuffix)
        End If
    ' -------------------------
    ' For top flange bottom cut
    ' -------------------------
    Case 17
        ' ---------------------
        ' Skip if no top flange
        ' ---------------------
        If (Not bPenetratesWeb) And HasTopFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopFlangeBtmCut" & strCaseSuffix)
        End If
    ' ----------------------------
    ' For bottom flange center cut
    ' ----------------------------
    Case 18
        ' ------------------------
        ' Skip if no bottom flange
        ' ------------------------
        If (Not bPenetratesWeb) And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmFlangeCenterCut" & strCaseSuffix)
        End If
    ' -------------------------
    ' For bottom flange Top cut
    ' -------------------------
    Case 19
        ' ------------------------
        ' Skip if no bottom flange
        ' ------------------------
        If (Not bPenetratesWeb) And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmFlangeTopCut" & strCaseSuffix)
        End If
    ' ----------------------------
    ' For bottom flange bottom cut
    ' ----------------------------
    Case 20
        ' ------------------------
        ' Skip if no bottom flange
        ' ------------------------
        If (Not bPenetratesWeb) And HasBottomFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmFlangeBtmCut" & strCaseSuffix)
        End If
    ' ----------------------------------------
    ' TopFaceTopCorner, left web or top flange
    ' ----------------------------------------
    Case 21
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceTopCorner_LT" & strCaseSuffix)
    ' -------------------------------------------
    ' TopFaceBottomCorner, left web or top flange
    ' -------------------------------------------
    Case 22
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceBtmCorner_LT" & strCaseSuffix)
    ' --------------------------------------------
    ' TopEdgeOutsideCorner, left web or top flange
    ' --------------------------------------------
    Case 23
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopEdgeOutsideCorner_LT" & strCaseSuffix)
    ' -------------------------------------------
    ' TopEdgeInsideCorner, left web or top flange
    ' -------------------------------------------
    Case 24
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopEdgeInsideCorner_LT" & strCaseSuffix)
    ' -------------------------------------
    ' FaceTopCorner, left web or top flange
    ' -------------------------------------
    Case 25
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceCorner_LT" & strCaseSuffix)
    ' ----------------------------------------
    ' FaceBottomCorner, left web or top flange
    ' ----------------------------------------
    Case 26
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceCorner_LT" & strCaseSuffix)
    ' ----------------------------------------------
    ' BottomEdgeInsideCorner, left web or top flange
    ' ----------------------------------------------
    Case 27
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmEdgeInsideCorner_LT" & strCaseSuffix)
    ' -----------------------------------------------
    ' BottomEdgeOutsideCorner, left web or top flange
    ' -----------------------------------------------
    Case 28
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmEdgeOutsideCorner_LT" & strCaseSuffix)
    ' -------------------------------------------
    ' BottomFaceTopCorner, left web or top flange
    ' -------------------------------------------
    Case 29
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceTopCorner_LT" & strCaseSuffix)
    ' ----------------------------------------------
    ' BottomFaceBottomCorner, left web or top flange
    ' ----------------------------------------------
    Case 30
        ' -------------------------------------------
        ' Skip if flange-penetrated and no top flange
        ' -------------------------------------------
        If Not bPenetratesWeb And Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceBtmCorner_LT" & strCaseSuffix)
    ' --------------------------------------------
    ' TopFaceTopCorner, right web or bottom flange
    ' --------------------------------------------
    Case 31
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceTopCorner_RB" & strCaseSuffix)
    ' -----------------------------------------------
    ' TopFaceBottomCorner, right web or bottom flange
    ' -----------------------------------------------
    Case 32
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceBtmCorner_RB" & strCaseSuffix)
    ' ------------------------------------------------
    ' TopEdgeOutsideCorner, right web or bottom flange
    ' ------------------------------------------------
    Case 33
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopEdgeOutsideCorner_RB" & strCaseSuffix)
    ' -----------------------------------------------
    ' TopEdgeInsideCorner, right web or bottom flange
    ' -----------------------------------------------
    Case 34
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopEdgeInsideCorner_RB" & strCaseSuffix)
    ' -----------------------------------------
    ' FaceTopCorner, right web or bottom flange
    ' -----------------------------------------
    Case 35
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "TopFaceCorner_RB" & strCaseSuffix)
    ' --------------------------------------------
    ' FaceBottomCorner, right web or bottom flange
    ' --------------------------------------------
    Case 36
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceCorner_RB" & strCaseSuffix)
    ' --------------------------------------------------
    ' BottomEdgeInsideCorner, right web or bottom flange
    ' --------------------------------------------------
    Case 37
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmEdgeInsideCorner_RB" & strCaseSuffix)
    ' ---------------------------------------------------
    ' BottomEdgeOutsideCorner, right web or bottom flange
    ' ---------------------------------------------------
    Case 38
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmEdgeOutsideCorner_RB" & strCaseSuffix)
    ' -----------------------------------------------
    ' BottomFaceTopCorner, right web or bottom flange
    ' -----------------------------------------------
    Case 39
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceTopCorner_RB" & strCaseSuffix)
    ' --------------------------------------------------
    ' BottomFaceBottomCorner, right web or bottom flange
    ' --------------------------------------------------
    Case 40
        ' -------------------------------------------
        ' Skip if flange-penetrated and no btm flange
        ' Skip if web-penetrated and not right web
        ' -------------------------------------------
        If (bPenetratesWeb And Not HasRightWeb(oBoundedPart)) Or _
           (Not bPenetratesWeb And Not HasBottomFlange(oBoundedPart)) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "BtmFaceBtmCorner_RB" & strCaseSuffix)
    ' -------------------
    ' TopFlangeLeftCorner
    ' -------------------
    Case 41
        ' ----------------------------------------------
        ' Skip if not flange-penetrated or no top flange
        ' ----------------------------------------------
        If Not bPenetratesWeb Or Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPTopFlangeLeftInsideCorner" & strCaseSuffix)
    ' --------------------
    ' TopFlangeRightCorner
    ' --------------------
    Case 42
        ' ----------------------------------------------
        ' Skip if not flange-penetrated or no top flange
        ' ----------------------------------------------
        If Not bPenetratesWeb Or Not HasTopFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPTopFlangeRightInsideCorner" & strCaseSuffix)
    ' ----------------------
    ' BottomFlangeLeftCorner
    ' ----------------------
    Case 43
        ' -------------------------------------------------
        ' Skip if not flange-penetrated or no bottom flange
        ' -------------------------------------------------
        If Not bPenetratesWeb Or Not HasBottomFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPBtmFlangeLeftInsideCorner" & strCaseSuffix)
    ' -----------------------
    ' BottomFlangeRightCorner
    ' -----------------------
    Case 44
        ' -------------------------------------------------
        ' Skip if not flange-penetrated or no bottom flange
        ' -------------------------------------------------
        If Not bPenetratesWeb Or Not HasBottomFlange(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPBtmFlangeRightInsideCorner" & strCaseSuffix)
    ' ----------------
    ' LeftWebTopCorner
    ' ----------------
    Case 45
        ' --------------------------
        ' Skip if not web-penetrated
        ' --------------------------
        If bPenetratesWeb Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPLeftWebTopInsideCorner" & strCaseSuffix)
    ' -------------------
    ' LeftWebBottomCorner
    ' -------------------
    Case 46
        ' --------------------------
        ' Skip if not web-penetrated
        ' --------------------------
        If bPenetratesWeb Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPLeftWebBtmInsideCorner" & strCaseSuffix)
    ' -----------------
    ' RightWebTopCorner
    ' -----------------
    Case 47
        ' ------------------------------------------
        ' Skip if not web-penetrated or no right web
        ' ------------------------------------------
        If bPenetratesWeb Or Not HasRightWeb(oBoundedPart) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPRightWebTopInsideCorner" & strCaseSuffix)
    ' --------------------
    ' RightWebBottomCorner
    ' --------------------
    Case 48
        ' ------------------------------------------
        ' Skip if not web-penetrated or no right web
        ' ------------------------------------------
        If bPenetratesWeb Or Not HasRightWeb(oBoundedPart) Then
            Exit Function
        ElseIf bPenetratesWeb And HasRightWeb(oBoundedPart) And Not ConsiderRightWebAttributes(oBoundedPart) Then
            Exit Function
        End If
        
        GetSelection = GetStandardACAttribute(oAppConnection, "NPRightWebBtmInsideCorner" & strCaseSuffix)
    ' --------------------------------
    ' For split/miter primary left web
    ' --------------------------------
    Case 49
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        GetSelection = GetStandardACAttribute(oAppConnection, "PrimaryLeftWeb")
    ' ---------------------------------
    ' For split/miter primary right web
    ' ---------------------------------
    Case 50
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "PrimaryRightWeb")
        End If
    ' ----------------------------------
    ' For split/miter primary top flange
    ' ----------------------------------
    Case 51
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If HasTopFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "PrimaryTopFlange")
        End If
    ' -------------------------------------
    ' For split/miter primary bottom flange
    ' -------------------------------------
    Case 52
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If HasBottomFlange(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "PrimaryBottomFlange")
        End If
    ' ----------------------------------
    ' For split/miter secondary left web
    ' ----------------------------------
    Case 53
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        GetSelection = GetStandardACAttribute(oAppConnection, "SecondaryLeftWeb")
    ' -----------------------------------
    ' For split/miter secondary right web
    ' -----------------------------------
    Case 54
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If Not oBoundingPort Is Nothing Then 'Needed this check as when boudning to Slab or Plates BoundingPort is Nothing, as nothing is filled by MemmebrConnectionData
            If ConsiderRightWebAttributes(oBoundingPort.Connectable) Then
                GetSelection = GetStandardACAttribute(oAppConnection, "SecondaryRightWeb")
            End If
        End If
    ' ------------------------------------
    ' For split/miter secondary top flange
    ' ------------------------------------
    Case 55
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If Not oBoundingPort Is Nothing Then 'Needed this check as when boudning to Slab or Plates BoundingPort is Nothing, as nothing is filled by MemmebrConnectionData
            If HasTopFlange(oBoundingPort.Connectable) Then
                GetSelection = GetStandardACAttribute(oAppConnection, "SecondaryTopFlange")
            End If
        End If
    ' ---------------------------------------
    ' For split/miter secondary bottom flange
    ' ---------------------------------------
    Case 56
        If DoesValueToAttributeExist(oAppConnection, "BearingPlate") Then Exit Function
        If Not oBoundingPort Is Nothing Then 'Needed this check as when boudning to Slab or Plates BoundingPort is Nothing, as nothing is filled by MemmebrConnectionData
            If HasBottomFlange(oBoundingPort.Connectable) Then
               GetSelection = GetStandardACAttribute(oAppConnection, "SecondaryBottomFlange")
            End If
        End If
        
    ' ---------------------------------------
    ' For Top Seated PC
    ' ---------------------------------------
    ' Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
    Case 57
        
        If Not (eCase = OnMemberBottom Or eCase = OnMemberTop) Then Exit Function
        
        If IsTubularMember(oBoundedPart) Then Exit Function
        
        If bPenetratesWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopPC_LT" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopPC_LT" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And bTFR And cBFIOrFR.IntersectingEdge = eBounding_Edge.Top Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopPC_LT" & strCaseSuffix)
        End If
        
    ' ---------------------------------------
    ' For Top Seated PC
    ' ---------------------------------------
    ' Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrated
    Case 58
        
        If Not (eCase = OnMemberBottom Or eCase = OnMemberTop) Then Exit Function
            
        If IsTubularMember(oBoundedPart) Then Exit Function
                        
        If bPenetratesWeb And bRightWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopPC_RB" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And bBFR And cBFIOrFR.IntersectingEdge = eBounding_Edge.Top Then
            GetSelection = GetStandardACAttribute(oAppConnection, "TopPC_RB" & strCaseSuffix)
        End If
    ' ---------------------------------------
    ' For Bottom Seated PC
    ' ---------------------------------------
    ' Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
    Case 59
        
        If Not (eCase = OnMemberBottom Or eCase = OnMemberTop) Then Exit Function
    
        If IsTubularMember(oBoundedPart) Then Exit Function

        If bPenetratesWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmPC_LT" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmPC_LT" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And bTFL And cTFIOrFL.IntersectingEdge = eBounding_Edge.Bottom Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmPC_LT" & strCaseSuffix)
        End If
                
    ' ---------------------------------------
    ' For Bottom Seated PC
    ' ---------------------------------------
    ' Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrated
    Case 60
        
        If Not (eCase = OnMemberBottom Or eCase = OnMemberTop) Then Exit Function
    
        If IsTubularMember(oBoundedPart) Then Exit Function
        
        If bPenetratesWeb And bRightWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom And ConsiderRightWebAttributes(oBoundedPart) Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmPC_RB" & strCaseSuffix)
        ElseIf Not bPenetratesWeb And bBFL And cTFIOrFL.IntersectingEdge = eBounding_Edge.Bottom Then
            GetSelection = GetStandardACAttribute(oAppConnection, "BtmPC_RB" & strCaseSuffix)
        End If
        
    ' ---------------------------------------
    ' For Creating Bearing Plate
    ' ---------------------------------------
    Case 61

        GetSelection = GetStandardACAttribute(oAppConnection, "BearingPlate")
        
    ' ---------------------------------------
    ' For Creating Slot feature when member penetrates through member
    ' ---------------------------------------
    Case 62

        GetSelection = GetStandardACAttribute(oAppConnection, "SlotFeature")
        
    End Select

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Private Function GetNonPenFlangeCutSelection(oMD As IJDMemberDescription) As String

    Const METHOD = m_DefinitionProgid & "::GetNonPenFlangeCutSelection"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    Dim sSelection As String

    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim eCase As eMemberBoundingCase
    
    GetAssemblyConnectionInputs oMD.CAO, oBoundedPort, oBoundingPort

    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oMD.CAO, strCaseSuffix)
    
    Select Case oMD.dispid
        
        Case 7, 8, 9
            
            ' ------------------------------
            ' Skip if there is no top flange
            ' ------------------------------
            If Not HasTopFlange(oBoundedPort.Connectable) Then
                Exit Function
            End If
                
            GetNonPenFlangeCutSelection = GetStandardACAttribute(oMD.CAO, "NPTopFlangeCut" & strCaseSuffix)

        Case 10, 11, 12
        
            ' ---------------------------------
            ' Skip if there is no bottom flange
            ' ---------------------------------
            If Not HasBottomFlange(oBoundedPort.Connectable) Then
                Exit Function
            End If
        
            GetNonPenFlangeCutSelection = GetStandardACAttribute(oMD.CAO, "NPBtmFlangeCut" & strCaseSuffix)
 
    End Select
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' *******************************************************************************************
' Method:
' CMConstructBrace
'
' Description:
'
' *******************************************************************************************
Public Function CMConstructBrace(ByVal oMD As IJDMemberDescription, _
                                 ByVal oResourceManager As IUnknown, _
                                 ByRef oObject As Object)
    
    Const METHOD = "MbrACDefCM::CMConstructBrace"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Brace ...oMD.index = " & Str(oMD.Index)
        
    Set oObject = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub CMConstructCornerFeature(ByVal oMD As IJDMemberDescription, _
                                    ByVal oResourceManager As IUnknown, _
                                    ByRef oObject As Object)
                           
    Const METHOD = m_DefinitionProgid & "::CMConstructCornerFeature"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating corner feature ...oMD.index = " & Str(oMD.Index)
                           
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = oMD.CAO
            
    ' ----------------------------------
    ' Get the bounding and bounded ports
    ' ----------------------------------
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    
    GetAssemblyConnectionInputs oMD.CAO, oBoundedPort, oBoundingPort
    
    ' -----------------
    ' Get the selection
    ' -----------------
    Dim strSelection As String
    strSelection = GetSelection(oMD)
    
    ' ------------------------------------------------------------------------------------------------------------------
    ' Input ports are dependent on which feature is being created and where the web/flange intersect the bounding object
    ' ------------------------------------------------------------------------------------------------------------------
    ' A selector class is not used
    ' For cases where the corner is between two sub faces on the same end cut, the sub IDs are determined and
    ' then processed at the end
    ' Other cases are processed as they are found
    Dim ePortId1 As JXSEC_CODE
    Dim ePortId2 As JXSEC_CODE
    
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrFL As ConnectedEdgeInfo
    Dim eIBFOrFR As ConnectedEdgeInfo
    
    Dim dispID1 As Long
    Dim dispID2 As Long
    Dim oMbrACDefCM As Object
    
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
    
    Select Case oMD.dispid
        ' --------------------------------------------
        ' Top Face, Top Corner, Left Web or Top Flange
        ' --------------------------------------------
        Case 21, 31
            
            ePortId1 = JXSEC_WEB_RIGHT_TOP
            ePortId2 = JXSEC_TOP
        
        ' -----------------------------------------------
        ' Top Face, Bottom Corner, Left Web or Top Flange
        ' -----------------------------------------------
        Case 22, 32
                    
            ePortId1 = JXSEC_TOP_FLANGE_RIGHT_TOP
            ePortId2 = JXSEC_WEB_RIGHT_TOP
        ' -----------------------------------------------
        ' Top Edge Outside Corner, Left Web or Top Flange
        ' -----------------------------------------------
        Case 23, 33
                    
            ePortId1 = JXSEC_TOP_FLANGE_RIGHT
            ePortId2 = JXSEC_TOP_FLANGE_RIGHT_TOP
        ' ----------------------------------------------
        ' Top Edge Inside Corner, Left Web or Top Flange
        ' ----------------------------------------------
        Case 24, 34
                        
            ePortId1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
            ePortId2 = JXSEC_TOP_FLANGE_RIGHT
        
        ' ---------------------------------------
        ' Face Top Corner, Left Web or Top Flange
        ' ---------------------------------------
        Case 25
            
            ' -----------------------------------------------------------------------------------------------
            ' Ports to be used if bounded top overlaps bounding edge, creating a corner between face and edge
            ' -----------------------------------------------------------------------------------------------
            ePortId1 = JXSEC_WEB_RIGHT
            ePortId2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM

            ' ---------------------------------------------------------------
            ' If web-penetrated and the top completely intersects the face...
            ' ---------------------------------------------------------------
            Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID & "MbrAC.MbrACDefCM")
                    
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If bPenetratesWeb And eTopOrWL.IntersectingEdge = Web_Right And eITFOrFL.IntersectingEdge = Web_Right Then
                
                dispID1 = 1 ' Future use - might be right web cut (item 4)
                
                ' ----------------------------------------------------------------------------------------
                ' If there is a top flange, we must instead use the corner between the web and flange cuts
                ' ----------------------------------------------------------------------------------------
                If HasTopFlange(oBoundedPort.Connectable) Then
                    
                    If Not oMemObjs.ItemByDispid(7) Is Nothing Then
                        dispID2 = 7
                    ElseIf Not oMemObjs.ItemByDispid(8) Is Nothing Then
                        dispID2 = 8
                    ElseIf Not oMemObjs.ItemByDispid(9) Is Nothing Then
                        dispID2 = 9
                    Else
                        Exit Sub ' Unexpected
                    End If
                
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           dispID1, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' ----------------------------------------------------------------------------------------------
                ' If there is no top flange, we must use the corner between the web cut and the bounded top port
                ' ----------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(dispID1), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_TOP)
                    Exit Sub
                End If
            ' -------------------------------------------------------------------
            ' If flange-penetrated and the left completely intersects the face...
            ' -------------------------------------------------------------------
            '
            ' |                          |                             |
            ' |---------------FL         |--------------------WL/FL    |----------------------FL
            ' |---------------WL         |--------------------WR       |
            ' |                          |                             |
            ' |                          |                             |----------------------WL
            ' |                          |                             |----------------------WR
            ' |---------------WR         |                             |
            ' |---------------FR         |--------------------FR       |
            ' |                          |                             |----------------------FR
            ' |                          |                             |
            '
            ElseIf (Not bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And eITFOrFL.IntersectingEdge = Web_Right And eTopOrWL.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable) And Not bTFL And eITFOrFL.IntersectingEdge = Web_Right And eBtmOrWR.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And bTFL And eITFOrFL.IntersectingEdge = Web_Right) Then
                
                dispID2 = 13 ' Future use - could be item 14
                
                ' ------------------------------------------------------------------------------------------
                ' If there is no left flange, we must instead use the corner between the web and flange cuts
                ' ------------------------------------------------------------------------------------------
                If Not bTFL Then
                    
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           15, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' -----------------------------------------------------------------------------------------------------
                ' If there is a left flange, we must use the corner between the flange cut and the top flange left port
                ' -----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(15), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_TOP_FLANGE_LEFT)
                    Exit Sub
                End If
         
            End If
            
        ' ------------------------------------------
        ' Face Bottom Corner, Left Web or Top Flange
        ' ------------------------------------------
        Case 26
        
            ' --------------------------------------------------------------------------------------------------
            ' Ports to be used if bounded bottom overlaps bounding edge, creating a corner between face and edge
            ' --------------------------------------------------------------------------------------------------
            ePortId1 = JXSEC_WEB_RIGHT
            ePortId2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP

            ' ------------------------------------------------------------------
            ' If web-penetrated and the bottom completely intersects the face...
            ' ------------------------------------------------------------------
            Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID & "MbrAC.MbrACDefCM")
                    
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If bPenetratesWeb And eBtmOrWR.IntersectingEdge = Web_Right And eIBFOrFR.IntersectingEdge = Web_Right Then
                
                dispID1 = 1 ' Future use - might be right web cut (item 4)
                
                ' ----------------------------------------------------------------------------------------
                ' If there is a top flange, we must instead use the corner between the web and flange cuts
                ' ----------------------------------------------------------------------------------------
                If HasBottomFlange(oBoundedPort.Connectable) Then
                    
                    If Not oMemObjs.ItemByDispid(10) Is Nothing Then
                        dispID2 = 10
                    ElseIf Not oMemObjs.ItemByDispid(11) Is Nothing Then
                        dispID2 = 11
                    ElseIf Not oMemObjs.ItemByDispid(12) Is Nothing Then
                        dispID2 = 12
                    Else
                        Exit Sub ' Unexpected
                    End If
                
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           dispID1, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' ----------------------------------------------------------------------------------------------------
                ' If there is no bottom flange, we must use the corner between the web cut and the bounded bottom port
                ' ----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(dispID1), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_BOTTOM)
                    Exit Sub
                End If
            ' --------------------------------------------------------------------
            ' If flange-penetrated and the right completely intersects the face...
            ' --------------------------------------------------------------------
            '
            ' |                          |                             |
            ' |---------------FL         |--------------------FL       |----------------------FL
            ' |---------------WL         |                             |
            ' |                          |                             |
            ' |                          |                             |----------------------WL
            ' |                          |                             |----------------------WR
            ' |---------------WR         |--------------------WL       |
            ' |---------------FR         |--------------------WR/FR    |
            ' |                          |                             |----------------------FR
            ' |                          |                             |
            '
            ElseIf (Not bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And eIBFOrFR.IntersectingEdge = Web_Right And eBtmOrWR.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable) And Not bTFR And eIBFOrFR.IntersectingEdge = Web_Right And eTopOrWL.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And bTFR And eIBFOrFR.IntersectingEdge = Web_Right) Then
                    
                dispID2 = 13 ' Future use - could be item 14
                
                ' -------------------------------------------------------------------------------------------
                ' If there is no right flange, we must instead use the corner between the web and flange cuts
                ' -------------------------------------------------------------------------------------------
                If Not bTFR Then
                    
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           15, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' -----------------------------------------------------------------------------------------------------
                ' If there is a left flange, we must use the corner between the flange cut and the top flange left port
                ' -----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(15), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_TOP_FLANGE_RIGHT)
                    Exit Sub
                End If
            End If
            
        ' -------------------------------------------------
        ' Bottom Edge Inside Corner, Left Web or Top Flange
        ' -------------------------------------------------
        Case 27, 37
            
            ePortId1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
            ePortId2 = JXSEC_BOTTOM_FLANGE_RIGHT
            
        ' --------------------------------------------------
        ' Bottom Edge Outside Corner, Left Web or Top Flange
        ' --------------------------------------------------
        Case 28, 38
        
            ePortId1 = JXSEC_BOTTOM_FLANGE_RIGHT
            ePortId2 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
                        
        ' -----------------------------------------------
        ' Bottom Face, Top Corner, Left Web or Top Flange
        ' -----------------------------------------------
        Case 29, 39
        
            ePortId1 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
            ePortId2 = JXSEC_WEB_RIGHT_BOTTOM

        ' --------------------------------------------------
        ' Bottom Face, Bottom Corner, Left Web or Top Flange
        ' --------------------------------------------------
        Case 30, 40
        
            ePortId1 = JXSEC_WEB_RIGHT_BOTTOM
            ePortId2 = JXSEC_BOTTOM
                    
        Case 35
            
            ' -----------------------------------------------------------------------------------------------
            ' Ports to be used if bounded top overlaps bounding edge, creating a corner between face and edge
            ' -----------------------------------------------------------------------------------------------
            ePortId1 = JXSEC_WEB_RIGHT
            ePortId2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM

            ' ---------------------------------------------------------------
            ' If web-penetrated and the top completely intersects the face...
            ' ---------------------------------------------------------------
            Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID & "MbrAC.MbrACDefCM")
                    
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If bPenetratesWeb And eTopOrWL.IntersectingEdge = Web_Right And eITFOrFL.IntersectingEdge = Web_Right Then
                
                dispID1 = 1 ' Future use - might be right web cut (item 4)
                
                ' ----------------------------------------------------------------------------------------
                ' If there is a top flange, we must instead use the corner between the web and flange cuts
                ' ----------------------------------------------------------------------------------------
                If HasTopFlange(oBoundedPort.Connectable) Then
                    
                    If Not oMemObjs.ItemByDispid(7) Is Nothing Then
                        dispID2 = 7
                    ElseIf Not oMemObjs.ItemByDispid(8) Is Nothing Then
                        dispID2 = 8
                    ElseIf Not oMemObjs.ItemByDispid(9) Is Nothing Then
                        dispID2 = 9
                    Else
                        Exit Sub ' Unexpected
                    End If
                
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           dispID1, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' ----------------------------------------------------------------------------------------------
                ' If there is no top flange, we must use the corner between the web cut and the bounded top port
                ' ----------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(dispID1), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_TOP)
                    Exit Sub
                End If
            ' -------------------------------------------------------------------
            ' If flange-penetrated and the left completely intersects the face...
            ' -------------------------------------------------------------------
            '
            ' |                          |                             |
            ' |---------------FL         |--------------------WL/FL    |----------------------FL
            ' |---------------WL         |--------------------WR       |
            ' |                          |                             |
            ' |                          |                             |----------------------WL
            ' |                          |                             |----------------------WR
            ' |---------------WR         |                             |
            ' |---------------FR         |--------------------FR       |
            ' |                          |                             |----------------------FR
            ' |                          |                             |
            '
            ElseIf (Not bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And eITFOrFL.IntersectingEdge = Web_Right And eTopOrWL.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable) And Not bBFL And eITFOrFL.IntersectingEdge = Web_Right And eBtmOrWR.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And bTFL And eITFOrFL.IntersectingEdge = Web_Right) Then
                
                dispID2 = 13 ' Future use - could be item 14
                
                ' ------------------------------------------------------------------------------------------
                ' If there is no left flange, we must instead use the corner between the web and flange cuts
                ' ------------------------------------------------------------------------------------------
                If Not bBFL Then
                    
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           18, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' -----------------------------------------------------------------------------------------------------
                ' If there is a left flange, we must use the corner between the flange cut and the top flange left port
                ' -----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(18), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_BOTTOM_FLANGE_LEFT)
                    Exit Sub
                End If
         
            End If
            
        ' ------------------------------------------
        ' Face Bottom Corner, Left Web or Top Flange
        ' ------------------------------------------
        Case 36
        
            ' --------------------------------------------------------------------------------------------------
            ' Ports to be used if bounded bottom overlaps bounding edge, creating a corner between face and edge
            ' --------------------------------------------------------------------------------------------------
            ePortId1 = JXSEC_WEB_RIGHT
            ePortId2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP

            ' ------------------------------------------------------------------
            ' If web-penetrated and the bottom completely intersects the face...
            ' ------------------------------------------------------------------
            Set oMbrACDefCM = SP3DCreateObject(CUSTOMERID & "MbrAC.MbrACDefCM")
                    
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If bPenetratesWeb And eBtmOrWR.IntersectingEdge = Web_Right And eIBFOrFR.IntersectingEdge = Web_Right Then
                
                dispID1 = 1 ' Future use - might be right web cut (item 4)
                
                ' ----------------------------------------------------------------------------------------
                ' If there is a top flange, we must instead use the corner between the web and flange cuts
                ' ----------------------------------------------------------------------------------------
                If HasBottomFlange(oBoundedPort.Connectable) Then
                    
                    If Not oMemObjs.ItemByDispid(10) Is Nothing Then
                        dispID2 = 10
                    ElseIf Not oMemObjs.ItemByDispid(11) Is Nothing Then
                        dispID2 = 11
                    ElseIf Not oMemObjs.ItemByDispid(12) Is Nothing Then
                        dispID2 = 12
                    Else
                        Exit Sub ' Unexpected
                    End If
                
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           dispID1, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' ----------------------------------------------------------------------------------------------------
                ' If there is no bottom flange, we must use the corner between the web cut and the bounded bottom port
                ' ----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(dispID1), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_BOTTOM)
                    Exit Sub
                End If
            ' --------------------------------------------------------------------
            ' If flange-penetrated and the right completely intersects the face...
            ' --------------------------------------------------------------------
            '
            ' |                          |                             |
            ' |---------------FL         |--------------------FL       |----------------------FL
            ' |---------------WL         |                             |
            ' |                          |                             |
            ' |                          |                             |----------------------WL
            ' |                          |                             |----------------------WR
            ' |---------------WR         |--------------------WL       |
            ' |---------------FR         |--------------------WR/FR    |
            ' |                          |                             |----------------------FR
            ' |                          |                             |
            '
            ElseIf (Not bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And eIBFOrFR.IntersectingEdge = Web_Right And eBtmOrWR.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable) And Not bBFR And eIBFOrFR.IntersectingEdge = Web_Right And eTopOrWL.IntersectingEdge = Web_Right) Or _
                   (Not bPenetratesWeb And bTFR And eIBFOrFR.IntersectingEdge = Web_Right) Then
                    
                dispID2 = 13 ' Future use - could be item 14
                
                ' -------------------------------------------------------------------------------------------
                ' If there is no right flange, we must instead use the corner between the web and flange cuts
                ' -------------------------------------------------------------------------------------------
                If Not bBFR Then
                    
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                           oResourceManager, _
                                                                                           18, _
                                                                                           dispID2, _
                                                                                           strSelection)
                    Exit Sub
                ' -----------------------------------------------------------------------------------------------------
                ' If there is a left flange, we must use the corner between the flange cut and the top flange left port
                ' -----------------------------------------------------------------------------------------------------
                Else
                    Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(18), _
                                                                                          oResourceManager, _
                                                                                          JXSEC_BOTTOM_FLANGE_RIGHT)
                    Exit Sub
                End If
            End If
                    
        ' ------------------------------------------
        ' Non-Penetrating left web top inside corner
        ' ------------------------------------------
        Case 41
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")

            If HasTopFlange(oBoundedPort.Connectable) Then
                ' 13 is first web cut
                ' 15 is top flange center cut - this corner should not be at a top or bottom flange cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       13, _
                                                                                       15, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(13), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_TOP)
                Exit Sub
            End If
        ' ---------------------------------------------
        ' Non-Penetrating left web bottom inside corner
        ' ---------------------------------------------
        Case 42
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
                        
            If HasBottomFlange(oBoundedPort.Connectable) Then
                ' 13 is first web cut
                ' 15 is top flange center cut - this corner should not be at a top or bottom flange cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       13, _
                                                                                       18, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(13), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_BOTTOM)
                Exit Sub
            End If
        
        ' -------------------------------------------
        ' Non-Penetrating right web top inside corner
        ' -------------------------------------------
        Case 43
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            If HasTopFlange(oBoundedPort.Connectable) Then
                ' 13 is first web cut.  Future use - change to second (right) web cut if/when implemented
                ' 15 is top flange center cut - this corner should not be at a top or bottom flange cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       13, _
                                                                                       15, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(13), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_TOP)
                Exit Sub
            End If

        ' ----------------------------------------------
        ' Non-Penetrating right web bottom inside corner
        ' ----------------------------------------------
        Case 44
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            If HasBottomFlange(oBoundedPort.Connectable) Then
                ' 13 is first web cut.  Future use - change to second (right) web cut if/when implemented
                ' 15 is top flange center cut - this corner should not be at a top or bottom flange cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       13, _
                                                                                       18, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(13), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_BOTTOM)
                Exit Sub
            End If

        ' ---------------------------------------------
        ' Non-Penetrating top flange left inside corner
        ' ---------------------------------------------
        Case 45
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If Not bTFL Then
                ' 1 is first web cut(in future, could be item 4)
                ' 7 is top flange with center web cut as input - this corner should not be at a top or bottom web cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       1, _
                                                                                       7, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(7), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_TOP_FLANGE_LEFT)
                Exit Sub
            End If
            
        ' ----------------------------------------------
        ' Non-Penetrating top flange right inside corner
        ' ----------------------------------------------
        Case 46
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If Not bTFR Then
                ' 1 is first web cut(in future, could be item 4)
                ' 7 is top flange with center web cut as input - this corner should not be at a top or bottom web cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       1, _
                                                                                       7, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(7), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_TOP_FLANGE_RIGHT)
                Exit Sub
            End If

        ' ------------------------------------------------
        ' Non-Penetrating bottom flange left inside corner
        ' ------------------------------------------------
        Case 47
        
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If Not bBFL Then
                ' 1 is first web cut(in future, could be item 4)
                ' 10 is bottom flange with center web cut as input - this corner should not be at a top or bottom web cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       1, _
                                                                                       10, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(10), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_BOTTOM_FLANGE_LEFT)
                Exit Sub
            End If
            
        ' ----------------------------------------------
        ' Non-Penetrating top flange right inside corner
        ' ----------------------------------------------
        Case 48
            
            Set oMbrACDefCM = SP3DCreateObject(m_sStdACProjectName & "MbrAC.MbrACDefCM")
            
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            If Not bBFR Then
                ' 1 is first web cut(in future, could be item 4)
                ' 10 is bottom flange with center web cut as input - this corner should not be at a top or bottom web cut
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenTwoEndCutsByDispID(oMD, _
                                                                                       oResourceManager, _
                                                                                       1, _
                                                                                       10, _
                                                                                       strSelection)
                Exit Sub
            Else
                ' This item is really intended to be placed between web and flange cuts, but just in case...
                Set oObject = oMbrACDefCM.CreateCornerFeatureBetweenEndAndLateralPort(oMemObjs.ItemByDispid(10), _
                                                                                      oResourceManager, _
                                                                                      JXSEC_BOTTOM_FLANGE_RIGHT)
                Exit Sub
            End If

    End Select
    
    Dim oParentEndCut As Object
    Dim bIsBottomFlange As Boolean
    
    Select Case oMD.dispid
        Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
            If bPenetratesWeb Then
                Set oParentEndCut = oMemObjs.ItemByDispid(1)
            Else
                Set oParentEndCut = oMemObjs.ItemByDispid(15)
                bIsBottomFlange = False
            End If
        Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
            If bPenetratesWeb Then
                Set oParentEndCut = oMemObjs.ItemByDispid(4)
            Else
                Set oParentEndCut = oMemObjs.ItemByDispid(18)
                bIsBottomFlange = True
            End If
        Case Else
            ' unexpected
            Exit Sub
    End Select
    
    Dim oFacePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    
    'MbrCornerFeatureDataByObject oParentEndCut, ePortID1, ePortID2, oFacePort, oEdgePort1, oEdgePort2
    
    Dim oBoundedPart As Object
    Set oBoundedPart = oBoundedPort.Connectable
    
    If bPenetratesWeb Then
        Set oFacePort = GetLateralSubPortBeforeTrim(oBoundedPart, JXSEC_WEB_LEFT)
    Else
        If bIsBottomFlange Then
            Set oFacePort = GetLateralSubPortBeforeTrim(oBoundedPart, JXSEC_BOTTOM)
        Else
            Set oFacePort = GetLateralSubPortBeforeTrim(oBoundedPart, JXSEC_TOP)
        End If
    End If

    GetMbrEndCutCornerPorts oParentEndCut, oBoundedPart, ePortId1, ePortId2, oFacePort, oEdgePort1, oEdgePort2
    
    
    If Not oFacePort Is Nothing And Not oEdgePort1 Is Nothing And Not oEdgePort2 Is Nothing Then
        
        sMsg = "Creating Member Corner Feature"
        Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
        Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
        oSDO_CornerFeature.Create oResourceManager, _
                                  oFacePort, _
                                  oEdgePort1, _
                                  oEdgePort2, _
                                  strSelection, _
                                  oParentEndCut
    
        sMsg = "Returning Member CornerFeature just created"
        
        Set oObject = oSDO_CornerFeature.object
    
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'---------------------------------------------------------------------------------------------------------------
'Function: GetStandardACAttribute
'
'DESCRIPTION:
' get attribute value for the given attrbute name from xml

'Input:
'SmartOccurence:

'Output:  the attribute value as string
'
'---------------------------------------------------------------------------------------------------------------
Public Function GetStandardACAttribute(oSO As IJSmartOccurrence, attrName As String)

    Const METHOD = "::GetStandardACAttribute"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Getting the standard AC attribute for the given attribute name from xml "
    
    Dim xmldoc As IXMLDOMDocument
    Dim XmlMainNode As IXMLDOMNode
    
    
    Dim oRefDataMiddleServices As RefDataMiddleServices.StructXMLCacheServices
    Set oRefDataMiddleServices = New RefDataMiddleServices.StructXMLCacheServices
    
    If m_sSymbolSharePath = vbNullString Then
        m_sSymbolSharePath = GetSymbolShare()
    End If
    
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSO.ItemObject
    ' Get attribute
    Dim sXMLPathAttribute As String
    sXMLPathAttribute = GetCustomAttribute(oSmartItem, "IJUAStdACDefinition", "AttributeList")
    If sXMLPathAttribute = vbNullString Then Exit Function
    Dim oFileSysObject As FileSystemObject
    Set oFileSysObject = New FileSystemObject
    
    Dim sPath As String
    sPath = oFileSysObject.BuildPath(m_sSymbolSharePath, "\" & sXMLPathAttribute)
    If Not oFileSysObject.FileExists(sPath) Then
        sMsg = "xml path or file doesn't exist"
        GoTo ErrorHandler ' log the error
    End If
    Set xmldoc = oRefDataMiddleServices.LoadXML(sPath)
    Set oRefDataMiddleServices = Nothing
    ' Disable asynchronous loading
    xmldoc.async = False

    If Not xmldoc Is Nothing Then
        Dim sAttrNameInQuote As String
        Dim sSmartItemNameInQuote As String
        sAttrNameInQuote = "'" & attrName & "'" 'represents  'attrName'
        sSmartItemNameInQuote = "'" & oSmartItem.Name & "'" 'Example  'EE0001'
        ' search for the xml node which has this attribute name
        Set XmlMainNode = xmldoc.selectSingleNode("//StandardACs/StandardAC[@Name=" & sSmartItemNameInQuote & "]/Attribute[@Name=" & sAttrNameInQuote & "]")
        If Not XmlMainNode Is Nothing Then
            GetStandardACAttribute = XmlMainNode.Attributes.getNamedItem("Value").nodeValue
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub CM_UpdateCache(oPD As IJDPropertyDescription, oObject As Object)

    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    'setting the FlipPrimaryandSecondary Attribute value based on the condition
    Set_FlipAttributeValue oPD.CAO
    GetAssemblyConnectionInputs oObject, oBoundedPort, oBoundingPort

    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    Dim sectionAlias As Long
    Dim oEdgeMapColl As JCmnShp_CollectionAlias

    Dim bPenetratesWeb As Boolean

    If Not GetMemberBoundingCase(oObject, , , True) = OnTubeMember And Not IsTubularMember(oBoundingPort.Connectable) Then
        ' This also calculates and caches the egde map and section alias
        GetConnectedEdgeInfo oObject, oBoundedPort, oBoundingPort, oTopORWL, oBottomOrWR, oTFL, oTFR, , , , , , True
    End If
End Sub

'*************************************************************************
' Function
' GetAxisACInputs
'
' Abstract
'
'***************************************************************************
Public Sub GetACInputs(oAppConnection As IJAppConnection, oBoundedPort As IJPort, oBoundingPort As IJPort)
                                    
    ' This method is no more valid, use GetAssemblyConnectionInputs() to get the AC inputs
     

End Sub


Public Function CMConstructPhysicalConnection(ByVal pMD As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    ByRef pObject As Object)
                                 
    Const METHOD = "ACDef::CMConstructPhysicalConnection"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Creating seated PC...pMD.index = " & Str(pMD.Index)
    
    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    
    Set oAppConnection = pMD.CAO
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
    
      
    ' --------------------------
    ' Get the relative positions
    ' --------------------------
    Dim cWLOrTop As ConnectedEdgeInfo
    Dim cWROrBtm As ConnectedEdgeInfo
    Dim cTFIOrFL As ConnectedEdgeInfo
    Dim cBFIOrFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, cWLOrTop, cWROrBtm, cTFIOrFL, cBFIOrFR

    ' -----------------
    ' If web penetrated
    ' -----------------
    Dim boundingConnPortXid As JXSEC_CODE
    Dim boundedConnPortXid As JXSEC_CODE
    
    boundingConnPortXid = JXSEC_UNKNOWN
    boundedConnPortXid = JXSEC_UNKNOWN
    
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)

    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    Dim strPCAttribute As String
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oAppConnection, strCaseSuffix)
    
    Dim oBoundedPart As Object
    Set oBoundedPart = oBoundedPort.Connectable
    
    Dim bRightWeb  As Boolean
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    Dim bTopFlange As Boolean
    Dim bBottomFlange As Boolean
    
    bTopFlange = HasTopFlange(oBoundedPart, bTFL, bTFR)
    bBottomFlange = HasBottomFlange(oBoundedPart, bBFL, bBFR)
    bRightWeb = HasRightWeb(oBoundedPart)
       
    Select Case pMD.dispid
    'Axis_MbrToEdge_SeatedPC_
    'Top SeatedPC: Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
     Case 57
        strPCAttribute = "TopPC_LT" & strCaseSuffix
        
        If bPenetratesWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_BOTTOM
            boundingConnPortXid = JXSEC_TOP
        
        ElseIf Not bPenetratesWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_WEB_RIGHT
            boundingConnPortXid = JXSEC_TOP
            
        ElseIf Not bPenetratesWeb And bTFR And cBFIOrFR.IntersectingEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_TOP_FLANGE_RIGHT
            boundingConnPortXid = JXSEC_TOP
        End If
    ' ---------------------------------------
    ' For Top Seated PC
    ' ---------------------------------------
    ' Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrat
    Case 58
        strPCAttribute = "TopPC_RB" & strCaseSuffix
        If bPenetratesWeb And bRightWeb And cWROrBtm.IntersectingEdge = eBounding_Edge.Top And ConsiderRightWebAttributes(oBoundedPart) Then
            boundedConnPortXid = JXSEC_RIGHT_WEB_BOTTOM
            boundingConnPortXid = JXSEC_TOP
            
        ElseIf Not bPenetratesWeb And bBFR And cBFIOrFR.IntersectingEdge = eBounding_Edge.Top Then

            boundedConnPortXid = JXSEC_BOTTOM_FLANGE_RIGHT
            boundingConnPortXid = JXSEC_TOP
        End If
    ' ---------------------------------------
    ' For Bottom Seated PC
    ' ---------------------------------------
    ' Left Web PC for Web Penetrated/ Top Flange PC for Flange Penetrated
    Case 59
        strPCAttribute = "BtmPC_LT" & strCaseSuffix
        If bPenetratesWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom Then
            
            boundedConnPortXid = JXSEC_TOP
            boundingConnPortXid = JXSEC_BOTTOM
            
        ElseIf Not bPenetratesWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom Or cWLOrTop.IntersectingEdge = eBounding_Edge.Below Then
            
            boundedConnPortXid = JXSEC_WEB_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
            
        ElseIf Not bPenetratesWeb And bTFL And cTFIOrFL.IntersectingEdge = eBounding_Edge.Bottom Then
            
            boundedConnPortXid = JXSEC_TOP_FLANGE_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
            
        End If
    ' ---------------------------------------
    ' For Bottom Seated PC
    ' ---------------------------------------
    ' Right Web PC for Web Penetrated/ Bottom Flange PC for Flange Penetrated
    Case 60
        strPCAttribute = "BtmPC_RB" & strCaseSuffix
        If bPenetratesWeb And bRightWeb And cWLOrTop.IntersectingEdge = eBounding_Edge.Bottom And ConsiderRightWebAttributes(oBoundedPart) Then
            boundedConnPortXid = JXSEC_RIGHT_WEB_TOP
            boundingConnPortXid = JXSEC_BOTTOM
            
        ElseIf Not bPenetratesWeb And bBFL And cTFIOrFL.IntersectingEdge = eBounding_Edge.Bottom Then
            
            boundedConnPortXid = JXSEC_BOTTOM_FLANGE_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
        End If
    End Select
    
    ' --------------------
    ' Get the edge mapping
    ' --------------------
    'Get the Edge Mapping and Section Alias
    Dim pSectionAlias As Long
    Dim oEdgeMap As JCmnShp_CollectionAlias
    Set oEdgeMap = New Collection

    Set oEdgeMap = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, pSectionAlias, bPenetratesWeb)
    ' ------------------------------------------
    ' Get connectable ports on the last geometry
    ' ------------------------------------------
    If boundedConnPortXid = JXSEC_UNKNOWN Or boundingConnPortXid = JXSEC_UNKNOWN Then
        sMsg = "Could not identify ports for connection"
        GoTo ErrorHandler
    End If
    
    Dim oBoundedConnPort As IJPort
    Dim oBoundingConnPort As IJPort
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim oLateBindingPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oStructSymbolTools As IJStructSymbolTools
    Set oStructSymbolTools = New StructSymbolTools
    
    If pHelper.ObjectType(oBoundedPort.Connectable) = SDOBJECT_STIFFENER Then
        Set oProfilePart.object = oBoundedPort.Connectable
        Set oBoundedConnPort = oProfilePart.SubPort(boundedConnPortXid)
    ElseIf pHelper.ObjectType(oBoundedPort.Connectable) = SDOBJECT_MEMBER Then
        Set oBoundedConnPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, boundedConnPortXid)
        Set oStructPort = oBoundedConnPort
        oStructSymbolTools.GetLateBindingPort oBoundedPort.Connectable, oStructPort.ProxyType, oStructPort.OperationID, oStructPort.OperatorID, oStructPort.ContextID, oStructPort.OperatorID, oLateBindingPort
        Set oBoundedConnPort = oLateBindingPort
    End If

    Dim RealEdgeID As JXSEC_CODE
    If ItemExists(boundingConnPortXid, oEdgeMap, RealEdgeID) Then
        boundingConnPortXid = RealEdgeID
        If pHelper.ObjectType(oBoundingPort.Connectable) = SDOBJECT_STIFFENER Then
            Set oProfilePart.object = oBoundingPort.Connectable
            Set oBoundingConnPort = oProfilePart.SubPort(boundingConnPortXid)
        ElseIf pHelper.ObjectType(oBoundingPort.Connectable) = SDOBJECT_MEMBER Then
            Set oBoundingConnPort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, boundingConnPortXid)
            Set oBoundingConnPort = pHelper.GetEquivalentLastPort(oBoundingConnPort)
        End If
    End If

    If oBoundedConnPort Is Nothing Or oBoundingConnPort Is Nothing Then
        sMsg = "Could not identify ports for connection"
        GoTo ErrorHandler
    End If
    
    Dim strSelection As String
    Dim strFilter As String
    
    strSelection = GetStandardACAttribute(oAppConnection, strPCAttribute)
        
    oCustomMethods.GetPCSelection pMD, strPCAttribute, strSelection, strFilter
    
    If strSelection = vbNullString Then
        strSelection = "LapWeld"
    End If
    
    ' --------------------------
    ' Create physical connection
    ' --------------------------
    Dim oPhysicalConnection As New PhysicalConn
    sMsg = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oBoundedConnPort, _
                               oBoundingConnPort, _
                               strSelection, _
                               oAppConnection, _
                               ConnectionStandard
    
    If Not oPhysicalConnection Is Nothing Then
        Set pObject = oPhysicalConnection.object
    Else
        sMsg = "Creation of physical connection failed"
        GoTo ErrorHandler
    End If
    
    ' ------------------------------
    ' Set filter progID on PC object
    ' ------------------------------
    oCustomMethods.AddPCFilter pObject, strFilter

  Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

Public Function DoesValueToAttributeExist(oSO As IJSmartOccurrence, attrName As String, Optional attrValue As String = "") As Boolean

    Const METHOD = m_DefinitionProgid & "::DoesValueToAttributeExist"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    attrValue = GetStandardACAttribute(oSO, attrName)
    
    If attrValue = vbNullString Then
        DoesValueToAttributeExist = False
    Else
        DoesValueToAttributeExist = True
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function



' *******************************************************************************************
' Method:
' CMConstructBearingPlate
'
' Description:
' Custom method for creating a BearingPlate
' *******************************************************************************************
Public Function CMConstructBearingPlate(ByVal oMD As IJDMemberDescription, _
                                  ByVal oResourceManager As IUnknown, _
                                  ByRef oObject As Object)
    
    Const METHOD = m_DefinitionProgid & "::CMConstructBearingPlate"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String

    sMsg = "Creating EndCut ...oMD.index = " & Str(oMD.Index)
    lDispId = oMD.dispid
    
    ' Create BearingPlate/Web Cut
    sMsg = "Create Bearing Plate on Root Selection Rule"
        sEndCutSel = GetSelection(oMD)
        
        Create_BearingPlate oMD, oResourceManager, _
                            sEndCutSel, False, oObject


    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the respective smart occurences based on need
' *******************************************************************************************
Public Sub CMNeedToCompute(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "ACDef::CMNeedToCompute"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a smart occurence"

    If pObject Is Nothing Then Exit Sub
    
    Dim oSmartOccurance As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    
    Set oSmartOccurance = pObject
    Set oSmartItem = oSmartOccurance.ItemObject
    
    If oSmartItem Is Nothing Then Exit Sub
    
    If TypeOf pObject Is IJSmartPlate Then
        Dim eSmartPlateType As SmartPlateTypes
        Dim oSmartPlate As IJSmartPlate
    
        Set oSmartPlate = pObject
    
        If oSmartPlate.SmartPlateType = SmartPlateTypes.spType_BEARING Then
    
             Dim strBPattribute As String
             strBPattribute = GetStandardACAttribute(pPropertyDescriptions.CAO, "BearingPlate")
             If (oSmartItem.Name <> strBPattribute) Then
                 Dim oGraphicInputs As JCmnShp_CollectionAlias
                 Dim oSPAttributes As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
                 Dim oBearingPlate As IJSmartPlate
                 Dim oSPDefinition As GSCADSDCreateModifyUtilities.IJSDSmartPlateDefinition

                 Set oBearingPlate = pObject
                 Set oSPAttributes = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils

                 Set oGraphicInputs = New Collection
                 oSPAttributes.GetInputs_BearingPlate oBearingPlate, oGraphicInputs

                 Set oSPDefinition = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils

                 Set oBearingPlate = oSPDefinition.CreateBearingPlatePart(GetResourceManagerFromObject(pObject), _
                                                                         strBPattribute, _
                                                                        oGraphicInputs, _
                                                                         pPropertyDescriptions.CAO, oBearingPlate)
            End If
        End If
    ' When member cross-section increases then The slot should update accordingly. This is added to support such modification cases.
    ElseIf TypeOf pObject Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = pObject
        
        Dim eFeatureType As StructFeatureTypes
        eFeatureType = oFeature.get_StructFeatureType
        
        If eFeatureType = SF_Slot Then
            oSmartOccurance.Update
        ElseIf eFeatureType = SF_WebCut Or eFeatureType = SF_FlangeCut Then
        
            Dim bIsBoxCut As Boolean
            Dim eBoundingCase As eMemberBoundingCase
            bIsBoxCut = IsBoxCut(pPropertyDescriptions.CAO)
            
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            Dim oMemberPart As ISPSMemberPartCommon
            ' This recreation of webcut in place of existing webcut should happen only when the below cases are happened
            
            If bIsBoxCut Then
            
                GetAssemblyConnectionInputs pPropertyDescriptions.CAO, oBoundedPort, oBoundingPort
                
                If eFeatureType = SF_WebCut Then
                
                    Dim oSD_Webcut As New StructDetailObjects.WebCut
                    Set oSD_Webcut.object = pObject
                    
                    ' Check the old bounded from webcut and flipped bounded (if happened) is same or not.In this case.
                    ' In this case, Replacing or recreating the webcut is not needed, other wise needed.
                    If oBoundedPort.Connectable Is oSD_Webcut.Bounded Then
                        Exit Sub
                    End If
                    
                    'Get the bouindind along and bounded end port and create the webcut with this ports
                    GetBoundingEndPort oSD_Webcut.BoundedPort, oSD_Webcut.BoundingPort, oBoundedPort
                    Set oMemberPart = oSD_Webcut.Bounded
                    Set oBoundingPort = oMemberPart.AxisPort(SPSMemberAxisAlong)
                    
                    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
                    Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
                                            
                    Set pObject = oFeatureUtils.CreateWebCut(GetResourceManagerFromObject(pObject), _
                                                    oBoundingPort, oBoundedPort, oSD_Webcut.ItemName, _
                                                    pPropertyDescriptions.CAO, oSD_Webcut.object)
                                                    
                ElseIf eFeatureType = SF_FlangeCut Then
                
                    Dim oSD_FlangeCut As New StructDetailObjects.FlangeCut
                    Set oSD_FlangeCut.object = pObject
                    
                    ' Check the old bounded from webcut and flipped bounded (if happened) is same or not.In this case.
                    ' In this case, Replacing or recreating the webcut is not needed, other wise needed.
                    If oBoundedPort.Connectable Is oSD_FlangeCut.Bounded Then
                         Exit Sub
                    End If
                
                    GetBoundingEndPort oSD_FlangeCut.BoundedPort, oSD_FlangeCut.BoundingPort, oBoundedPort
                    Set oMemberPart = oSD_FlangeCut.Bounded
                    
                    Set oBoundingPort = oMemberPart.AxisPort(SPSMemberAxisAlong)
                    
                    Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
                    Set pObject = oFeatureUtils.CreateFlangeCut(GetResourceManagerFromObject(pObject), oBoundingPort, _
                                                            oBoundedPort, oSD_FlangeCut.WebCut, oSD_FlangeCut.ItemName, _
                                                            oSD_FlangeCut.WebCut, oSD_FlangeCut.object)
                End If
            Else
                UpdateFeatureForSelectionChange pPropertyDescriptions, pObject
                
            End If
        ElseIf eFeatureType = SF_CornerFeature Then
            UpdateFeatureForSelectionChange pPropertyDescriptions, pObject
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

Public Sub CMConstructSlot(pMemberDescription As IJDMemberDescription, pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    Const METHOD = "CMConstructSlot"
    
    Dim oSlot As New StructDetailObjects.Slot
    Dim strStartClass As String
    Dim pSystemParent As IJSystemChild
    Dim oPenetratingPort As IJPort
    Dim oPenetratedPort As IJPort

    GetAssemblyConnectionInputs pMemberDescription.CAO, oPenetratingPort, oPenetratedPort
    strStartClass = GetSelection(pMemberDescription)
    Set pSystemParent = pMemberDescription.CAO
    Call oSlot.Create(pResourceManager, oPenetratingPort.Connectable, oPenetratedPort.Connectable, strStartClass, pSystemParent)
    'strError = "Setting Slot to private variable"
    Set pObject = oSlot.object
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Could not create Slot CutOut").Number
End Sub


'*************************************************************************
'Function
'CM_CreateEndToEndFlangeCut
'
'Abstract
'   Create FlangeCut given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'   lEndCutSet
'   oSDO_WebCut
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub CM_CreateEndToEndFlangeCut(pMemberDescription As IJDMemberDescription, _
                            pResourceManager As IUnknown, _
                            sEndCutSelRule As String, _
                            bUseBoundingEndPort As Boolean, _
                            bUseBoundedAlongPort As Boolean, _
                            lEndCutSet As Long, _
                            oSDO_WebCut As Object, _
                            pEndCutObject As Object)
    Const METHOD = "ACDef::CM_CreateEndToEndFlangeCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    Dim nWebCuts As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oTempBddPort As IJPort
    Dim oTempBdngPort As IJPort
    Dim oBddpart As ISPSMemberPartPrismatic
    Dim oPort As Object
    Dim oWebCut As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
        GetAssemblyConnectionInputs oAppConnection, oTempBddPort, oTempBdngPort
    
    Set oBoundedPort = oTempBddPort
    If bUseBoundedAlongPort Then
        ' By default, the Bounded Port is the AxisStart or AxisEnd Port
        ' For ShortBox cases, need the Bounded AxisLong Port
        Set oBddpart = oTempBddPort.Connectable
        Set oBoundedPort = oBddpart.AxisPort(SPSMemberAxisAlong)
    End If
    
    Set oBoundingPort = oTempBdngPort
    If bUseBoundingEndPort Then
        ' By default, the Bounding Port is the AlongAxis Port
        ' For End to End cases, need the Bounding End Port
        GetBoundingEndPort oTempBddPort, oTempBdngPort, oBoundingPort
    End If
    

    If lEndCutSet = 2 Then
        ' For the Second set of End Cuts
        ' Switch the Bounding and Bounded Ports used to create the EndCut
        Set oPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oPort
        Set oPort = Nothing
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    
    If oSDO_WebCut Is Nothing Then
        nWebCuts = 0
        Set oMemberObjects = oAppConnection
        For iIndex = 1 To oMemberObjects.Count
            If Not oMemberObjects.Item(iIndex) Is Nothing Then
                If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                    Set oStructFeature = oMemberObjects.Item(iIndex)
                    If oStructFeature.get_StructFeatureType = SF_WebCut Then
                        nWebCuts = nWebCuts + 1
                        If nWebCuts = lEndCutSet Then
                            Set oWebCut = oStructFeature
                            Exit For
                        End If
                    End If
                End If
            End If
        Next iIndex
    Else
        If TypeOf oSDO_WebCut Is IJStructFeature Then
            Set oStructFeature = oSDO_WebCut
            If oStructFeature.get_StructFeatureType = SF_WebCut Then
                Set oWebCut = oSDO_WebCut
            End If
        End If
    End If
    
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oWebCut, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pEndCutObject = oSDO_FlangeCut.object
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

Public Function IsBoxCut(oACObject As IJAppConnection) As Boolean
                                    
    Const METHOD = "ACDef::IsBoxCut"
    On Error GoTo ErrorHandler
    
    Dim strPrimaryLeftWeb As String
    Dim strSecondaryLeftWeb As String
    
    IsBoxCut = False
    
    strPrimaryLeftWeb = GetStandardACAttribute(oACObject, "PrimaryLeftWeb")
    strSecondaryLeftWeb = GetStandardACAttribute(oACObject, "SecondaryLeftWeb")
    
    If (strPrimaryLeftWeb = vbNullString) And (strSecondaryLeftWeb <> vbNullString) Then
        IsBoxCut = True
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

'******************************************************************************************
' Method:
' CMNeedToComputeForPrimaryAttributes
'
' Description: Updates the respective webcut based on need
' *******************************************************************************************
Public Sub CMNeedToComputeForPrimaryAttributes(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "ACDef::CMNeedToComputeForPrimaryAttributes"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the Webcut "

     If TypeOf pObject Is IJStructFeature Then
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        If Not oSmartItem Is Nothing Then
            Dim oFeature As IJStructFeature
            Set oFeature = pObject
            If oFeature.get_StructFeatureType = SF_WebCut Then
                Dim strPrimaryLWebattribute As String
                                
                Dim strAnswer As String
                Dim oBoundedPort As IJPort
                Dim oBoundingPort As IJPort
                Dim oSD_Webcut As New StructDetailObjects.WebCut
                Set oSD_Webcut.object = pObject
                
                strPrimaryLWebattribute = GetStandardACAttribute(pPropertyDescriptions.CAO, "PrimaryLeftWeb")
                GetAssemblyConnectionInputs pPropertyDescriptions.CAO, oBoundedPort, oBoundingPort
                
                If IsSplitOrMiter(pPropertyDescriptions.CAO) = True And _
                     (StrComp(oSmartItem.Name, strPrimaryLWebattribute, vbTextCompare) = 0) Then
                    'Update Split/Miter Endcut and exit
                    UpdateFeatureForSelectionChange pPropertyDescriptions, pObject
                    Exit Sub
                End If
                
                If (oSD_Webcut.Bounded Is oBoundedPort.Connectable) And (StrComp(oSmartItem.Name, strPrimaryLWebattribute, vbTextCompare) = 0) Then
                    Exit Sub
                ElseIf Not oSD_Webcut.Bounded Is oBoundedPort.Connectable Then
                    If TypeOf oSD_Webcut.BoundingPort Is ISPSSplitAxisEndPort Then
                        Set oBoundedPort = oSD_Webcut.BoundingPort
                        Set oBoundingPort = oSD_Webcut.BoundedPort
                    Else
                        GetBoundingEndPort oSD_Webcut.BoundedPort, oSD_Webcut.BoundingPort, oBoundedPort
                        Dim oMemberPart As ISPSMemberPartCommon
                        Set oMemberPart = oSD_Webcut.Bounded
                        Set oBoundingPort = oMemberPart.AxisPort(SPSMemberAxisAlong)
                    End If
                ElseIf (oSmartItem.Name <> strPrimaryLWebattribute) Then
                    Set oBoundedPort = oSD_Webcut.BoundedPort
                    Set oBoundingPort = oSD_Webcut.BoundingPort
                Else
                    Exit Sub
                End If
                
                Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
                Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
                Dim oParentObject As IJSystem
                
                Set pObject = oFeatureUtils.CreateWebCut(GetResourceManagerFromObject(pObject), oBoundingPort, oBoundedPort, _
                                                                           strPrimaryLWebattribute, oParentObject, oSD_Webcut.object)
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'******************************************************************************************
' Method:
' CMNeedToComputeForSecondaryAttributes
'
' Description: Updates the respective Webcuts based on need
' *******************************************************************************************
Public Sub CMNeedToComputeForSecondaryAttributes(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "ACDef::CMNeedToComputeForSecondaryAttributes"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the Webcut "

     If TypeOf pObject Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Dim oSmartOccurance As IJSmartOccurrence
        
        Dim oSmartItem As IJSmartItem
        Set oSmartOccurance = pObject
        
        Dim oSD_Webcut As New StructDetailObjects.WebCut
        
        Set oSD_Webcut.object = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        If Not oSmartItem Is Nothing Then
            Set oFeature = pObject
            If oFeature.get_StructFeatureType = SF_WebCut Then
                Dim strSecondaryLeftWebattribute As String
                
                Dim strAnswer As String
                Dim oBoundedPort As IJPort
                Dim oBoundingPort As IJPort
                
                strSecondaryLeftWebattribute = GetStandardACAttribute(pPropertyDescriptions.CAO, "SecondaryLeftWeb")
                GetAssemblyConnectionInputs pPropertyDescriptions.CAO, oBoundedPort, oBoundingPort
                
                If IsSplitOrMiter(pPropertyDescriptions.CAO) = True And _
                     (StrComp(oSmartItem.Name, strSecondaryLeftWebattribute, vbTextCompare) = 0) Then
                    'Update Split/Miter Endcut and exit
                    UpdateFeatureForSelectionChange pPropertyDescriptions, pObject
                    Exit Sub
                End If
                'This recreation of webcut in place of existing webcut is needed when the old bounded (On webcut) and updated bounded (if flipped) is not same.
                ' This is needed when S0001 is changed to S0002 AC.
                If (oSD_Webcut.Bounded Is oBoundingPort.Connectable) And (StrComp(oSmartItem.Name, strSecondaryLeftWebattribute, vbTextCompare) = 0) Then
                    Exit Sub
                ElseIf Not oSD_Webcut.Bounded Is oBoundingPort.Connectable Then
                    If TypeOf oSD_Webcut.BoundingPort Is ISPSSplitAxisEndPort Then
                        Set oBoundedPort = oSD_Webcut.BoundingPort
                        Set oBoundingPort = oSD_Webcut.BoundedPort
                    Else
                        GetBoundingEndPort oSD_Webcut.BoundedPort, oSD_Webcut.BoundingPort, oBoundedPort
                        Dim oMemberPart As ISPSMemberPartCommon
                        Set oMemberPart = oSD_Webcut.Bounded
                        Set oBoundingPort = oMemberPart.AxisPort(SPSMemberAxisAlong)
                    End If
                ElseIf (oSmartItem.Name <> strSecondaryLeftWebattribute) Then
                    Set oBoundedPort = oSD_Webcut.BoundedPort
                    Set oBoundingPort = oSD_Webcut.BoundingPort
                Else
                    Exit Sub
                End If
                
                Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
                Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
                Dim oParentObject As IJSystem
                
                Set pObject = oFeatureUtils.CreateWebCut(GetResourceManagerFromObject(pObject), oBoundingPort, oBoundedPort, _
                                                                           strSecondaryLeftWebattribute, oParentObject, oSD_Webcut.object)
                End If

            End If
        End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'******************************************************************************************
' Method:
' CMNeedToComputeForSecondaryFlangecut
'
' Description: Updates the respective Webcuts based on need
' *******************************************************************************************
Public Sub CMNeedToComputeForSecondaryFlangecut(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "ACDef::CMNeedToComputeForSecondaryFlangecut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the Flangecut "
    
     If TypeOf pObject Is IJStructFeature Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        If Not oSmartItem Is Nothing Then
            Dim oFeature As IJStructFeature
            Set oFeature = pObject
            
            If oFeature.get_StructFeatureType = SF_FlangeCut Then
                                
                Dim oSD_FlangeCut As New StructDetailObjects.FlangeCut
                Set oSD_FlangeCut.object = pObject
                
                Dim oBoundedPort  As IJPort
                Dim oBoundingPort As IJPort
                Dim strAnswer As String
                
                GetAssemblyConnectionInputs pPropertyDescriptions.CAO, oBoundedPort, oBoundingPort
                'This recreation of webcut in place of existing webcut is needed only when the old bounded (On webcut) and updated bounded (if flipped)
                ' is not same.
                                
                If oBoundingPort.Connectable Is oSD_FlangeCut.Bounded Or Not IsBoxCut(pPropertyDescriptions.CAO) Then
                    Exit Sub
                End If
                
                GetBoundingEndPort oSD_FlangeCut.BoundedPort, oSD_FlangeCut.BoundingPort, oBoundedPort
                                
                Dim oMemberPart As ISPSMemberPartCommon
                Set oMemberPart = oSD_FlangeCut.Bounded
                
                Set oBoundingPort = oMemberPart.AxisPort(SPSMemberAxisAlong)
                
                Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
                Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
                Set pObject = oFeatureUtils.CreateFlangeCut(GetResourceManagerFromObject(pObject), oBoundingPort, _
                                                        oBoundedPort, oSD_FlangeCut.WebCut, oSD_FlangeCut.ItemName, _
                                                        oSD_FlangeCut.WebCut, oSD_FlangeCut.object)
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Private Sub UpdateFeatureForSelectionChange(pPropertyDescriptions As IJDPropertyDescription, oFeature As IJStructFeature)

  Const METHOD = "ACDef::UpdateFeatureForSelectionChange"
  On Error GoTo ErrorHandler
  
  Dim sMsg As String
  sMsg = "Updating Feature for Smart item change"
    
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim eFeatureType As StructFeatureTypes
    
    eFeatureType = oFeature.get_StructFeatureType
    
    'Get Feature Type
    If eFeatureType = SF_WebCut Then
        iStart = 1
        iEnd = 14
    ElseIf eFeatureType = SF_FlangeCut Then
        iStart = 7
        iEnd = 20
    ElseIf eFeatureType = SF_CornerFeature Then
        iStart = 21
        iEnd = 48
    Else
        Exit Sub 'Not handled
    End If
        
    Dim strSelection As String
    Dim iCount As Integer
    
    Dim oMemberObjects As IJDMemberObjects
    Dim oMD As IJDMemberDescription
    Dim oMemberDescriptions As IJDMemberDescriptions
    
    'Get all the Item Objects
    Set oMemberObjects = pPropertyDescriptions.CAO
    Set oMemberDescriptions = oMemberObjects.MemberDescriptions
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pPropertyDescriptions.CAO

    Dim oSmartOccurance As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
        
    'Get Smart Item Name
    Set oSmartOccurance = oFeature
    Set oSmartItem = oSmartOccurance.ItemObject
    
    Dim bIsSmartItemChange As Boolean
    bIsSmartItemChange = False
        
    For iCount = iStart To iEnd
        Set oMD = oMemberDescriptions.ItemByDispid(iCount)
                
        Dim lDispId As Long
        
        If oMD.object Is oFeature Then
            strSelection = GetSelection(oMD)
            lDispId = oMD.dispid
            
            If InStr(1, strSelection, oSmartItem.Name, vbTextCompare) = 0 And _
                InStr(1, strSelection, oSmartOccurance.RootSelection, vbTextCompare) = 0 Then
                
                bIsSmartItemChange = True
            End If
            Exit For
        End If
    Next
        
    'Updating web and Flange cuts if there is change in AC smart item and no endcut change so that welds will get updated
    If (eFeatureType = SF_WebCut Or eFeatureType = SF_FlangeCut) And Not bIsSmartItemChange Then

        Dim oSelectionRuleInfo As IJSelectionRuleInfo
        Set oSelectionRuleInfo = pPropertyDescriptions.CAO

        Dim eSelectionModificationType As SelectionRuleModificationType
        eSelectionModificationType = oSelectionRuleInfo.IsSelectionRuleModified

        If eSelectionModificationType = SelectionRule_CurrentRuleModified Then
            'selection rule is modified, need to update the object on Interface IJStructGeometry
            oSmartOccurance.Update
            Exit Sub
        End If
    End If
    
    If Not bIsSmartItemChange Then Exit Sub
        
    Dim oSDFeatureUtils As New GSCADSDCreateModifyUtilities.SDFeatureUtils
    
    If eFeatureType = SF_WebCut Then
        Dim oSDO_WebCut As New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oFeature
        
        'Recompute WebCut
        oSDFeatureUtils.CreateWebCut GetResourceMgr(), oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, _
                                    strSelection, oAppConnection, oFeature
        
    ElseIf eFeatureType = SF_FlangeCut Then
        
        Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
        Set oSDO_FlangeCut.object = oFeature
        
        'Recompute Flangecut
        oSDFeatureUtils.CreateFlangeCut GetResourceMgr(), oSDO_FlangeCut.BoundingPort, oSDO_FlangeCut.BoundedPort, _
                                        oSDO_FlangeCut.WebCut, strSelection, oAppConnection, oFeature
        
    ElseIf eFeatureType = SF_CornerFeature Then
        
        Dim oFeatureAttributes As IJSDFeatureAttributes
        Set oFeatureAttributes = New SDFeatureUtils
        
        Dim oFacePort As IJPort
        Dim oEdgePort1 As IJPort
        Dim oEdgePort2 As IJPort
        
        'Get CF inputs
        oFeatureAttributes.get_CornerCutInputs oFeature, oFacePort, oEdgePort1, oEdgePort2
        
        'Remcompute the CF with updated smart item
        oSDFeatureUtils.CreateCornerCut GetResourceMgr(), oFacePort, oEdgePort1, oEdgePort2, _
                                        strSelection, oAppConnection, oFeature
        
    End If
    
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub
'***********************************************************************
' METHOD:  IsSplitOrMiter
'
' DESCRIPTION: Method to determine given standard AC is split or miter
'
' INPUTS:  AC object
'
' OUPUT: Retuns True if current AC is split or miter, other wise false
'***********************************************************************
Private Function IsSplitOrMiter(oACObject As IJAppConnection) As Boolean
                                    
    Const METHOD = "ACDef::IsSplitOrMiter"
    On Error GoTo ErrorHandler
    
    Dim strPrimaryLeftWeb As String
    Dim strSecondaryLeftWeb As String
    
    IsSplitOrMiter = False
    
    strPrimaryLeftWeb = GetStandardACAttribute(oACObject, "PrimaryLeftWeb")
    strSecondaryLeftWeb = GetStandardACAttribute(oACObject, "SecondaryLeftWeb")
    
    If (strPrimaryLeftWeb <> vbNullString) And (strSecondaryLeftWeb <> vbNullString) Then
        IsSplitOrMiter = True
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function
