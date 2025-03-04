VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gen_2B_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenAssyConnSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    16/May/2012 - Created
'*********************************************************************************************
Option Explicit

Const m_sClassName As String = "Gen_2B_Def"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'*Require by S:\StructuralModeling\Client\AsmConnCommands
Implements ISPSFACInputHelper
Implements IJUserAttributeMgmt

Implements IJDUserSymbolServices
'

' List any graphic Inputs that the Definition has here
Public Sub ItemInputs(pIH As IJDInputsHelper)
Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' List any graphic Inputs that the Definition has here
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

' List all the Item members that this SmartClass can create
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)

    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oMemDesc As IJDMemberDescription
    'lDispId 1 is for web cut
    Set oMemDesc = pMDs.AddMember("WebCut1", 1, "CM_CreateGenericWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    'lDispId 2 is for Top Flange Cut
    Set oMemDesc = pMDs.AddMember("TopFlangeCut", 2, "CM_ConstructTopFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CM_CopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    'lDispId 3 is for Bottom Flange Cut
    Set oMemDesc = pMDs.AddMember("BottomFlangeCut", 3, "CM_ConstructBottomFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CM_CopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_FinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = m_DefinitionProgid & "::CM_FinalConstructAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Final Construct of Assembly Connection"
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
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

' *******************************************************************************************
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

' ********************************************************************************************
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

' ********************************************************************************************
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

' ********************************************************************************************
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

Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_DefinitionProgid & "::CM_CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    'By Default making the web cut as true
    bIsNeeded = True

 Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

  On Error GoTo ErrorHandler
  Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutNeeded"
  
    bIsNeeded = False
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long
    
    Set oAppConnection = pMemberDescription.CAO
    
    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim bTopFL As Boolean
    Dim bTopFR As Boolean
    Dim bBtmFL As Boolean
    Dim bBtmFR As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTopFL, bBtmFL, bTopFR, bBtmFR
    
    'Check whether Top Flange exists
    If (bTopFL) Or (bTopFR) Then
       bIsNeeded = True
    End If
  
  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


Public Sub CM_IsBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

  On Error GoTo ErrorHandler
  Const METHOD = m_DefinitionProgid & "::CM_IsBottomFlangeCutNeeded"
    
    Dim sMsg As String
    Dim lStatus As Long
    bIsNeeded = False
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    Set oAppConnection = pMemberDescription.CAO
    
    'Set the Bounded and Bounding Port
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
