VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndLongBoxDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrEndLongBoxDef.cls
'
'Author : D.A. Trent
'
'Description :
'   Control creation of the Member Items for the End LongBox Selector Class
'
'
'History:
'
'********************************************************************

Option Explicit

Const m_sClassName As String = "MbrEndLongBoxDef"
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
    Dim sDispId As String
    
    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    ' End to End LongBox EndCuts can consist of:
    '   (1) LongBox Web Cut, (1) Web Cut and (2) Flange Cuts
    ' Index(1) - LongBox (Free) Web Cut on Bounded Member
    ' Index(2) - Web Cut for Bounding Member
    ' Index(3) - Top Flange Cut for Bounding Member
    ' Index(4) - Bottom Flange Cut for Bounding Member
    For lDispId = 1 To 4
        sDispId = "LongBoxPlate_" & Trim(Str(lDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructEndCut", imsCOOKIE_ID_USS_LIB)
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsEndCutNeeded"
        oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
        oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
        Set oMemDesc = Nothing
    Next lDispId
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
Public Sub CM_EndCutCount(pMemberDescription As IJDMemberDescription, nCount As Long)
Const METHOD = m_DefinitionProgid & "::CM_EndCutCount"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lStatus As Long
    nCount = 1
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_IsEndCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = m_DefinitionProgid & "::CM_IsEndCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sSelection As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim bTopFlange As Boolean
    Dim bBottomFlange As Boolean
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    
    ' Based on DispId number:
    '   determine if LongBox/WebCut/FlangeCut is needed
    ' Index(1) - LongBox (Free) Web Cut on Bounded Member
    ' Index(2) - Web Cut for Bounding Member
    ' Index(3) - Top Flange Cut for Bounding Member
    ' Index(4) - Bottom Flange Cut for Bounding Member
    If lDispId = 1 Then
        bIsNeeded = True
    ElseIf lDispId = 2 Then
        bIsNeeded = True
    Else
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                                 lStatus, sMsg
        EndCut_IdealizedWebFlanges oBoundingData.MemberPart, eIdealized_Unk, _
                                   bTopFlange, bBottomFlange
        If lDispId = 3 Then
            If bTopFlange Then
                bIsNeeded = True
            End If
        ElseIf lDispId = 4 Then
            If bBottomFlange Then
                bIsNeeded = True
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
Const METHOD = m_DefinitionProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sEndCutSel As String
    Dim sEndCutProgId As String
    Dim sAssemblyProgId As String
    
    Dim lDispId As Long

    sMsg = "Initializing Web Cut data"
    lDispId = pMemberDescription.dispid
    sAssemblyProgId = CUSTOMERID + "MBR_AsmbConn.MbrEndLongBoxSel"
    
    sMsg = "Setting Answers on EndCut from Assembly Connection"
    If lDispId = 1 Then
        sEndCutSel = "MBR_LongBoxEndSel"
        sEndCutProgId = CUSTOMERID + "Mbr_EndCuts.MbrLongBoxEndSel"
        Set_WebCutQuestions pMemberDescription, sEndCutSel, _
                            sEndCutProgId, sAssemblyProgId, True, True
    
    ElseIf lDispId = 2 Then
        sEndCutSel = "MBR_AxisWebSel"
        sEndCutProgId = CUSTOMERID + "Mbr_EndCuts.MbrAxisWebSel"
        Set_WebCutQuestions pMemberDescription, sEndCutSel, _
                            sEndCutProgId, sAssemblyProgId, False, True

    ElseIf lDispId = 3 Then
        sEndCutSel = "MBR_EndFlangeSel"
        sEndCutProgId = CUSTOMERID + "Mbr_EndCuts.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutSel, _
                               sEndCutProgId, sAssemblyProgId, False, False
    ElseIf lDispId = 4 Then
        sEndCutSel = "MBR_EndFlangeSel"
        sEndCutProgId = CUSTOMERID + "Mbr_EndCuts.MbrEndFlangeSel"
        Set_FlangeCutQuestions pMemberDescription, sEndCutSel, _
                               sEndCutProgId, sAssemblyProgId, False, True
    End If
    

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Function CM_ConstructEndCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
Const METHOD = m_DefinitionProgid & "::CM_ConstructEndCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String

    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Create LongBox WebCut, WebCut, FlangeCuts with the Root Selection Rule
    sMsg = "Create EndCut based on Root Selection Rule"
    If lDispId = 1 Then
        sEndCutSel = "MBR_LongBoxEndSel"
        Create_WebCut pMemberDescription, pResourceManager, _
                      sEndCutSel, False, False, 1, pObject
    ElseIf lDispId = 2 Then
        sEndCutSel = "MBR_AxisWebSel"
        Create_WebCut pMemberDescription, pResourceManager, _
                      sEndCutSel, True, True, 2, pObject
    ElseIf lDispId = 3 Then
        sEndCutSel = "MBR_EndFlangeSel"
        Create_FlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, True, 2, pObject
    ElseIf lDispId = 4 Then
        sEndCutSel = "MBR_EndFlangeSel"
        Create_FlangeCut pMemberDescription, pResourceManager, _
                         sEndCutSel, True, True, 2, pObject
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

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
    pIH.Definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
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




