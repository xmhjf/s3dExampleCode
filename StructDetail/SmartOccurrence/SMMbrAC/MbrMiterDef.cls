VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrMiterDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrMiterDef.cls
'
'  Description : Control creation of the Member Items for the MbrMiterDef Class
'
'  Author      : Alligators
'
'  History     :
'    03/MAY/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "MbrMiterDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"


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
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description : List any graphic Inputs that the Definition has here
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
'           ****Miter EndCuts can consist of (2) Web Cut and (4) Flange Cuts
'
'                 Index(1) - Web Cut for Bounded Member
'                 Index(2) - Top Flange Cut for Bounded Member
'                 Index(3) - Bottom Flange Cut for Bounded Member
'                 Index(4) - Web Cut for Bounding Member
'                 Index(5) - Top Flange Cut for Bounding Member
'                 Index(6) - Bottom Flange Cut for Bounding Member
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
    ' if any more item members are needed to be added please add only at the bottom
    '******************************************************************************
    
    lDispId = 1 '- Web Cut for Bounded Member
    sDispId = "Miter_WebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    lDispId = 2 '- Top Flange Cut for Bounded Member
    sDispId = "Miter_FlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oMemDesc = Nothing
    
    lDispId = 3 '- Bottom Flange Cut for Bounded Member
    sDispId = "Miter_FlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oMemDesc = Nothing
    
    lDispId = 4 '- Web Cut for Bounding Member
    sDispId = "Miter_WebCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    lDispId = 5 '- Top Flange Cut for Bounding Member
    sDispId = "Miter_FlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oMemDesc = Nothing
    
    lDispId = 6 '- Bottom Flange Cut for Bounding Member
    sDispId = "Miter_FlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************

'*********************************************************************************************
' Method      : CM_IsEndCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsEndCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsEndCutNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    Dim sAnswer As String
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    
    ' Based on DispId number, determine if WebCut/FlangeCut is needed
    ' DispId(1) - Web Cut for Bounded Member
    ' DispId(2) - Top Flange Cut for Bounded Member
    ' DispId(3) - Bottom Flange Cut for Bounded Member
    ' DispId(4) - Web Cut for Bounding Member
    ' DispId(5) - Top Flange Cut for Bounding Member
    ' DispId(6) - Bottom Flange Cut for Bounding Member
    If lDispId = 1 Then
        bIsNeeded = True
        
    ElseIf lDispId < 4 Then
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
        CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
        
        If lDispId = 2 Then
            If (bTFL) Or (bTFR) Then
                bIsNeeded = True
            End If
        ElseIf lDispId = 3 Then
            If (bBFL) Or (bBFR) Then
                bIsNeeded = True
            End If
        End If

    ElseIf lDispId = 4 Then
        bIsNeeded = True
        
    Else
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
        CrossSection_Flanges oBoundingData.MemberPart, bTFL, bBFL, bTFR, bBFR

        If lDispId = 5 Then
            If (bTFL) Or (bTFR) Then
                bIsNeeded = True
            End If
        ElseIf lDispId = 6 Then
            If (bBFL) Or (bBFR) Then
                bIsNeeded = True
            End If
        End If
        
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : CM_CopyQuestionsToMember
' Description :
'
'*********************************************************************************************
Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
    Const METHOD = m_DefinitionProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sEndCutProgId As String
    
    Dim lDispId As Long

    sMsg = "Initializing Web Cut data"
    lDispId = pMemberDescription.dispid

    sMsg = "Setting Answers on EndCut from Assembly Connection"
    If lDispId = 1 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrMiterWebSel"
        Set_WebCutQuestions pMemberDescription, sEndCutProgId, False, False
        
    ElseIf lDispId = 2 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutProgId, False, False
        
    ElseIf lDispId = 3 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutProgId, False, True
        
    ElseIf lDispId = 4 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrMiterWebSel"
        Set_WebCutQuestions pMemberDescription, sEndCutProgId, True, False
        
    ElseIf lDispId = 5 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutProgId, True, False
        
    ElseIf lDispId = 6 Then
        sEndCutProgId = CUSTOMERID + "MbrEndCut.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutProgId, True, True
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructEndCut
' Description :
'
'*********************************************************************************************
Public Function CM_ConstructEndCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructEndCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String

    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    Dim clsMbrACDefCM As New MbrACDefCM
    
    ' Create Web Cut or Flange Cut base on the Root Selection Rule
    sMsg = "Create EndCut based on Root Selection Rule"
    If lDispId = 1 Then
        sEndCutSel = "MbrMiterWebSel"
        clsMbrACDefCM.CM_CreateEndToEndWebCut pMemberDescription, pResourceManager, _
                      sEndCutSel, True, False, 1, pObject
    ElseIf lDispId = 2 Then
        sEndCutSel = "MbrEndFlangeSel"
        clsMbrACDefCM.CM_CreateEndToEndFlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, False, 1, pObject
    ElseIf lDispId = 3 Then
        sEndCutSel = "MbrEndFlangeSel"
        clsMbrACDefCM.CM_CreateEndToEndFlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, False, 1, pObject
    ElseIf lDispId = 4 Then
        sEndCutSel = "MbrMiterWebSel"
        clsMbrACDefCM.CM_CreateEndToEndWebCut pMemberDescription, pResourceManager, _
                      sEndCutSel, True, False, 2, pObject
    ElseIf lDispId = 5 Then
        sEndCutSel = "MbrEndFlangeSel"
        clsMbrACDefCM.CM_CreateEndToEndFlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, False, 2, pObject
    ElseIf lDispId = 6 Then
        sEndCutSel = "MbrEndFlangeSel"
        clsMbrACDefCM.CM_CreateEndToEndFlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, False, 2, pObject
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

'*********************************************************************************************
' Method      : CM_FinalConstructAggregator
' Description :
'
'*********************************************************************************************
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
'
'*********************************************************************************************
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

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************


'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
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
' Description :
'
'*********************************************************************************************
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
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : ISPSFACInputHelper_UserAttributeMgmt
' Description :
'
'*********************************************************************************************
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() _
                            As SP3DStructInterfaces.IJUserAttributeMgmt
    Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_UserAttributeMgmt"
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
End Property

'*********************************************************************************************
' Method      : ISPSFACInputHelper_ValidateObjects
' Description :
'
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnAttributeChange
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange( _
                        ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                        ByVal CollAllDisplayedValues As Object, _
                        ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, _
                        ByVal varNewAttrValue As Variant) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnAttributeChange"

End Function

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnPreCommit
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit( _
                                ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreCommit"

End Function

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnPreLoad
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad( _
                                ByVal pIJDAttrs As SPSMembers.IJDAttributes, _
                                ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreLoad"

End Function


