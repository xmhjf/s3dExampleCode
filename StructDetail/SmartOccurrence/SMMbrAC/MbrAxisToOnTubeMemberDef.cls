VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisToOnTubeMemberDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAxisToOnTubeMemberDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "MbrAxisToOnTubeMemberDef"
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
' Index(4) - Flange Cut for Top Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                  Flange Normal Cut ---> if Web Penetrated)
'             *****This Index(4) item will be created only if the TopWebCut doesnt exist (i.e if Index(2) Item Member conditional is False)
'             *****if TopWebCut exist the next item Member(i.e Index(5)) will create the Flange Cut
'
' Index(8) - Flange Cut for Bottom Flange(if Exist) of Bounded Member(Flange Center Cut ---> if Web Non Penetrated OR
'                                                                     Flange Normal Cut ---> if Web Penetrated)
'             *****This Index(8) item will be created only if the BottomWebCut doesnt exist (i.e if Index(3) Item Member conditional is False)
'             *****if BottomWebCut exist the next item Member(i.e Index(9)) will create the Flange Cut
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
    sDispId = "Axis_MbrToEdge_WebCut1_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToCompute1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID + "MbrAC.MbrACDefCM"
    Set oMemDesc = Nothing
    Set oPropDesc = Nothing
    
    'The next Item Member should be responsible for creation of Flange Cut on Top
    '--->Center Cut or
    '--->Normal Cut(when TopWebCut doesn't exist)
    'The item Member will create the Cut only when Top Web Cut doesnt exist
    lDispId = 2
    sDispId = "Axis_MbrToEdge_TopFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oMemDesc = Nothing
        
    'should be responsible for creation of Flange Cut on Bottom(Ceneter Cut or Normal Cut)
    lDispId = 3
    sDispId = "Axis_MbrToEdge_BtmFlangeCut_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBtmFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CMCopyQuestionsToBottomFlangeCut"
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
    Const METHOD = m_DefinitionProgid & "::CM_IsBtmFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    bIsNeeded = False
    
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

' *******************************************************************************************
' Method:
' CMConstructWebCut
'
' Description:
' Custom method for creating a web cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM::CMConstructWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sEndCutSel As String
      
    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
   
    ' Create Web Cut based on the Root Selection Rule
    sMsg = "Create WebCut based on Root Selection Rule"
    
    'if the code is here we expect(not necessarily) the dispid should be = 1
    'dispid = 1(Center Cut or Normal Cut)
    sEndCutSel = "MbrAxisWebSel"
    
    Dim clsMbrACDefCM As New MbrACDefCM
    clsMbrACDefCM.CreateAxisWebCut pMemberDescription, pResourceManager, sEndCutSel, pObject
   
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
Public Function CMConstructFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                         ByVal pResourceManager As IUnknown, _
                                         ByRef pObject As Object)
    Const METHOD = "MbrACDefCM::CMConstructFlangeCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    
    lDispId = pMemberDescription.dispid
    
    Dim sEndCutSel As String
    
    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
  
    Dim clsMbrACDefCM As New MbrACDefCM
   
    'if Dispid = 4 or 5 should create Top Flange Cut(Center Cut or Normal Cut)
    sEndCutSel = "MbrAxisFlangeSel"
    clsMbrACDefCM.CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, False

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

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



