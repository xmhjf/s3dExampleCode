VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generic_LongBoxCutsDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : Generic_LongBoxCutsDef.cls
'
'  Description : Control creation of the Member Items for the Generic_LongBoxCutsDef Selector Class
'
'  Author      : Alligators
'
'  History     :
'    22/Nov/2012 - Created new Definition file so as to select longbox cuts when member is bounded by more than one boudings
'
'*********************************************************************************************

Const m_sClassName As String = "Generic_LongBoxCutsDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'*Require by S:\StructuralModeling\Client\AsmConnCommands
Implements ISPSFACInputHelper
Implements IJUserAttributeMgmt

Implements IJDUserSymbolServices
'

'*********************************************************************************************
' Method      : ItemInputs
' Description : ' List any graphic Inputs that the Definition has here
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
' Description : ' List all the Item members that this SmartClass can create
'               lDispId 1 through 3 are for web cuts
'               lDispId 4 and 5 are for top flange cut and bottom flange cut respectively
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)

        Const METHOD = m_DefinitionProgid & "::ItemMembers"
        On Error GoTo ErrorHandler
        
        Dim sMsg As String
        Dim sDispId As String
        Dim lDispId As Long
        Dim oMemDesc As IJDMemberDescription

   For lDispId = 1 To 5
        sDispId = "Generic_Web_" & Trim(Str(lDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructCut", imsCOOKIE_ID_USS_LIB)
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsCutNeeded"
        oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
        oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
        Set oMemDesc = Nothing
   Next lDispId

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
Public Sub CM_IsCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
  Const METHOD = m_DefinitionProgid & "::CM_IsCutNeeded"
     On Error GoTo ErrorHandler
        
        Dim sMsg As String
        Dim lDispId As Long
        Dim lStatus As Long
        Dim bTFL As Boolean, bBFL As Boolean, bTFR As Boolean, bBFR As Boolean
        Dim oAppConnection As IJAppConnection
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim oEditJDArgument As IJDEditJDArgument
        Dim oReferencesCollection As IJDReferencesCollection
        Dim oMemberObjects As IJDMemberObjects
        
        Set oAppConnection = pMemberDescription.CAO
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        Dim oBoundingObjectColl As IJElements
        Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        Set oMemberObjects = pMemberDescription.CAO

        bIsNeeded = False
        lDispId = pMemberDescription.dispid
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                                 lStatus, sMsg
        CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR

     If lDispId = 1 Then
           bIsNeeded = True
     ElseIf lDispId = 2 Or lDispId = 3 Then
           Dim oPort As IJPort
           If lDispId = 2 And (bTFL Or bTFR) Then
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
              If TypeOf GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort) Is IJPlatePart Then
                 bIsNeeded = True
              End If
           ElseIf lDispId = 3 And (bBFL Or bBFR) Then
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
              If TypeOf GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort) Is IJPlatePart Then
                 bIsNeeded = True
              End If
           End If
     
     ElseIf lDispId = 4 Then
           If bTFL Or bTFR Then
              If oMemberObjects.Item(2) Is Nothing Then
                 bIsNeeded = True
              End If
           End If
     ElseIf lDispId = 5 Then
           If bBFL Or bBFR Then
              If oMemberObjects.Item(3) Is Nothing Then
                 bIsNeeded = True
              End If
           End If
     End If
      
 Exit Sub
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
Const METHOD = m_DefinitionProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDispId As Long

    sMsg = "Initializing Web Cut data"
    lDispId = pMemberDescription.dispid
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    sMsg = "Setting Answers on EndCut from Assembly Connection"
        If lDispId = 1 Then
           '
        ElseIf lDispId = 2 Then
           '
        ElseIf lDispId = 3 Then
           '
        ElseIf lDispId = 4 Then
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrEndFlangeSel", "BottomFlange", "No"
        ElseIf lDispId = 5 Then
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrEndFlangeSel", "BottomFlange", "Yes"
        End If

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
Public Function CM_ConstructCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
Const METHOD = m_DefinitionProgid & "::CM_ConstructCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String

    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
     Dim oClsMbrACDefCM As New MbrACDefCM
    ' Create LongBox WebCut, WebCut, FlangeCuts with the RootSelection Rule
    ' Index(1) - LongBox (Free) Web Cut on Bounded Member
    ' Index(2) - Top Web Cut for Bounded Member
    ' Index(3) - Bottom Web Cut for Bounded Member
    ' Index(4) - Top Flange Cut for Bounded Member
    ' Index(5) - Bottom Flange Cut for Bounded Member
    sMsg = "Create EndCut based on Root Selection Rule"
    If lDispId = 1 Then
        sEndCutSel = "MbrLongBoxEndSel"
        oClsMbrACDefCM.CM_CreateGenericWebCut pMemberDescription, pResourceManager, pObject
    ElseIf lDispId = 2 Then
        sEndCutSel = "RootGenWebSel"
        oClsMbrACDefCM.CM_CreateGenericWebCut pMemberDescription, pResourceManager, pObject
    ElseIf lDispId = 3 Then
        sEndCutSel = "RootGenWebSel"
        oClsMbrACDefCM.CM_CreateGenericWebCut pMemberDescription, pResourceManager, pObject
    ElseIf lDispId = 4 Then
        sEndCutSel = "MbrEndFlangeSel"
        oClsMbrACDefCM.CM_ConstructTopFlangeCut pMemberDescription, pResourceManager, pObject
    ElseIf lDispId = 5 Then
        sEndCutSel = "MbrEndFlangeSel"
        oClsMbrACDefCM.CM_ConstructBottomFlangeCut pMemberDescription, pResourceManager, pObject
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

'*********************************************************************************************
' Method      : CM_MigrateEndCut
' Description :
'
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique

    IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid

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
