VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StifEndToOnTubeMemberDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : StifEndToOnTubeMemberDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/NOV/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "StifEndToOnTubeMemberDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
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
' Index(1) - Web Cut for Bounded Member
'
' Index(2) - Flange Cut for Top Flange(if Exist) of Bounded Member
'
' Index(3) - Flange Cut for Bottom Flange(if Exist) of Bounded Member
'
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
    sDispId = "StiffEndToMbrFace_WebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'The next Item Member should be responsible for creation of Flange Cut on Top
    lDispId = 2
    sDispId = "StiffEndToMbrFace_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    Set oMemDesc = Nothing
        
    'should be responsible for creation of Flange Cut on Bottom
    lDispId = 3
    sDispId = "StiffEndToMbrFace_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
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

'*********************************************************************************************
' Method      : CM_IsWebCutNeeded
' Description : Conditional method to determine if web cut is needed
'
'********************************************************************************************
Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
bIsNeeded = False
    Const METHOD = m_DefinitionProgid & "::CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
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
bIsNeeded = False
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

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR

    'Check whether Top Flange exisis
    If (bTopFL) Or (bTopFR) Then
       bIsNeeded = True
    Else
       bIsNeeded = False
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

    Set oAppConnection = pMemberDescription.CAO
    Dim lStatus As Long

    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg

    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR

    'Check whether bottom Flange exist
    If bBtmFL Or bBtmFR Then
       bIsNeeded = True
    Else
       bIsNeeded = False
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
'
'' *******************************************************************************************
'' Method:
'' CMConstructWebCut
''
'' Description:
'' Custom method for creating a web cut for a Axis Assembly connection
'' *******************************************************************************************
Public Function CMConstructWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
'    Const METHOD = "MbrACDefCM::CMConstructWebCut"
'    On Error GoTo ErrorHandler
'    Dim sMsg As String
'    Dim sEndCutSel As String
'
'    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
'
'    ' Create Web Cut based on the Root Selection Rule
'    sMsg = "Create WebCut based on Root Selection Rule"
'
'    'if the code is here we expect(not necessarily) the dispid should be = 1
'    'dispid = 1(Center Cut or Normal Cut)
'    sEndCutSel = "MbrAxisWebSel"
'
'    Dim clsMbrACDefCM As New MbrACDefCM
'    clsMbrACDefCM.CreateAxisWebCut pMemberDescription, pResourceManager, sEndCutSel, pObject
'
'    Exit Function
'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' Method:
' CMConstructFlangeCut
'
' Description:
' Custom method for creating a Flange cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                         ByVal pResourceManager As IUnknown, _
                                         ByRef pObject As Object)
'    Const METHOD = "MbrACDefCM::CMConstructFlangeCut"
'    On Error GoTo ErrorHandler
'
'    Dim sMsg As String
'    Dim lDispId As Long
'
'    lDispId = pMemberDescription.dispid
'
'    Dim sEndCutSel As String
'
'    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
'
'    Dim clsMbrACDefCM As New MbrACDefCM
'
'    'if Dispid = 4 or 5 should create Top Flange Cut(Center Cut or Normal Cut)
'    sEndCutSel = "MbrAxisFlangeSel"
'    clsMbrACDefCM.CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, False
'
'    Exit Function
'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

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
    Dim pCAFactory As CAFactory
    Set pCAFactory = New CAFactory
    
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

