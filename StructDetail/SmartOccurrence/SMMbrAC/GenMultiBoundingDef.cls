VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenMultiBoundingDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenMultiBoundingDef.cls
'
'  Description : Control creation of the Member Items for the GenAssyConnDef Selector Class
'
'  Author      : Alligators
'
'  History     :
'    05/MAY/2014 - Created
'    22/April/2015 - modified MDT TR-271041.Added properties to update the end cuts when bounded member is rotated after AC placement
'*********************************************************************************************

Const m_sClassName As String = "GenMultiBoundingDef"
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
'               lDispId 1 : For creating Web Cut
'               lDispId 2 : For creating Top Flange Cut
'               lDispId 3 : For creating Bottom Flange Cut
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)

    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Item Members"
    
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    
    Set oMemDesc = pMDs.AddMember("Generic_MB_WebCut", 1, "CM_CreateGenericMultiBoundingWebCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsMultiBoundingWebCutNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeWebCut", 1, "IJSmartOccurrence", "CMNeedtoComputeWebCut", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
   
    Set oMemDesc = pMDs.AddMember("Generic_MB_TopFlangeCut", 2, "CM_CreateGenericMultiBoundingTopFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsMultiBoundingTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeTopFlangeCut", 1, "IJSmartOccurrence", "CMNeedtoComputeTopFlangeCut", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    Set oMemDesc = pMDs.AddMember("Generic_MB_BtmFlangeCut", 3, "CM_CreateGenericMultiBoundingBottomFlangeCut", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsMultiBoundingBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateEndCut"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeBottomFlangeCut", 1, "IJSmartOccurrence", "CMNeedtoComputeBottomFlangeCut", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

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
' *******************************************************************************************
' Method:
' CM_IsMultiBoundingWebCutNeeded
'
' Description: Conditional method to determine if a Mutli Boundnig Single Web cut is needed
' *******************************************************************************************
Public Sub CM_IsMultiBoundingWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsMultiBoundingWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Conditional : Decide if Multi Bounding Web Cut is needed or not"
    
    bIsNeeded = True
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub
' *******************************************************************************************
' Method:
' CM_CreateGenericMultiBoundingWebCut
'
' Description:
' Custom method for creating a web cut for a Generic Multi Bounding Assembly connection
' *******************************************************************************************
Public Function CM_CreateGenericMultiBoundingWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_CreateGenericMultiBoundingWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Constructing Multi Bounding Generic Web Cut"

    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oRefPortsCollection As Collection
    Dim oMappedPortsCollection As Collection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim strEndCutSel As String
    
    ' ---------------------
    ' Get the Bounded Port
    ' ---------------------
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' ---------------------
    ' Get the Bounding Port
    ' ---------------------
    GetMultiBoundingEdgeMap oAppConnection, WebCut, oMappedPortsCollection
    
    Set oBoundingPort = oMappedPortsCollection.Item(1)
    strEndCutSel = "GenWebMultiBoundingSel"
    
    ' ---------------------
    ' Set the parent system
    ' ---------------------
    sMsg = "Retreiving Parent System for WebCut"
    
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' ------------------
    ' Create the Web Cut
    ' ------------------
    sMsg = "Creating WebCut"
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, strEndCutSel, oSystemParent
    Set pObject = oSDO_WebCut.object
    
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

' *******************************************************************************************
' Method:
' CM_IsMultiBoundingTopFlangeCutNeeded
'
' Description: Conditional method to determine if a Mutli Boundnig Top Flange cut is needed
' *******************************************************************************************
Public Sub CM_IsMultiBoundingTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsMultiBoundingTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Conditional : Decide if Multi Bounding Top Flange Cut is needed or not"
    
    bIsNeeded = False
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long

    ' ---------------------
    ' Get the Bounded Part
    ' ---------------------
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    If (bTopFlangeLeft Or bTopFlangeRight) Then
        bIsNeeded = True
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
' Method:
' CM_CreateGenericMultiBoundingTopFlangeCut
'
' Description:
' Custom method for creating a Top FLange Cut for a Generic Multi Bounding Assembly connection
' *******************************************************************************************
Public Function CM_CreateGenericMultiBoundingTopFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_CreateGenericMultiBoundingTopFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Constructing Multi Bounding Generic Top Flange Cut"

    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oMappedPortsCollection As Collection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim strEndCutSel As String
    
    ' ---------------------
    ' Get the Bounded Port
    ' ---------------------
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' ---------------------
    ' Get the Bounding Port
    ' ---------------------
    
    GetMultiBoundingEdgeMap oAppConnection, FlangeCutTop, oMappedPortsCollection
    
    Set oBoundingPort = oMappedPortsCollection.Item(1)
    strEndCutSel = "GenFlgMultiBoundingSel"
        
    'Create the Flange Cut
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oFeature As IJStructFeature
    Dim iIndex As Integer
    
    For iIndex = 1 To pMemberObjects.Count
        Set oFeature = pMemberObjects.Item(iIndex)
        If (Not (oFeature Is Nothing)) Then
            If (oFeature.get_StructFeatureType = SF_WebCut) Then
                Exit For
            End If
        End If
    Next iIndex
    
     ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oFeature, _
                       strEndCutSel, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pObject = oSDO_FlangeCut.object
    
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

' *******************************************************************************************
' Method:
' CM_IsMultiBoundingBottomFlangeCutNeeded
'
' Description: Conditional method to determine if a Mutli Boundnig Btm Flange cut is needed
' *******************************************************************************************
Public Sub CM_IsMultiBoundingBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsMultiBoundingBottomFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Conditional : Decide if Multi Bounding Bottom Flange Cut is needed or not"
    
    bIsNeeded = False
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long

    ' ---------------------
    ' Get the Bounded Part
    ' ---------------------
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    If (bBottomFlangeLeft Or bBottomFlangeRight) Then
        bIsNeeded = True
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
' Method:
' CM_CreateGenericMultiBoundingBottomFlangeCut
'
' Description:
' Custom method for creating a Top FLange Cut for a Generic Multi Bounding Assembly connection
' *******************************************************************************************
Public Function CM_CreateGenericMultiBoundingBottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_CreateGenericMultiBoundingBottomFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Constructing Multi Bounding Generic Bottom Flange Cut"

    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oMappedPortsCollection As Collection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    Dim strEndCutSel As String
    
    ' ---------------------
    ' Get the Bounded Port
    ' ---------------------
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' ---------------------
    ' Get the Bounding Port
    ' ---------------------
    GetMultiBoundingEdgeMap oAppConnection, FlangeCutBottom, oMappedPortsCollection
    
    Set oBoundingPort = oMappedPortsCollection.Item(oMappedPortsCollection.Count)
    strEndCutSel = "GenFlgMultiBoundingSel"
        
    'Create the Flange Cut
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oFeature As IJStructFeature
    Dim iIndex As Integer
    
    For iIndex = 1 To pMemberObjects.Count
        Set oFeature = pMemberObjects.Item(iIndex)
        If (Not (oFeature Is Nothing)) Then
            If (oFeature.get_StructFeatureType = SF_WebCut) Then
                Exit For
            End If
        End If
    Next iIndex
    
     ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oFeature, _
                       strEndCutSel, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pObject = oSDO_FlangeCut.object
    

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

' *******************************************************************************************
' Method:
' CM_CopyQuestionsToFlangeCut
'
' Description: copy questions to end cut rules to set BottomFlange to
'   - Yes; if the cut is to be applied on bottom flange
'   - No; if the cut is to be applied on Top Flange
' *******************************************************************************************
Public Sub CM_CopyQuestionsToFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsToFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Copy question answers to Flange Cut rules"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    Dim lDispId  As Long
    lDispId = pMemberDescription.dispid
    
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    If lDispId = 2 Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.GenMBFlangeSel", "BottomFlange", "No"
    Else
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.GenMBFlangeSel", "BottomFlange", "Yes"
    End If
    
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





'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the respective webcut based on previous and updated bounding port variation
' *******************************************************************************************
Public Sub CMNeedtoComputeWebCut(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "GenMultiBoundingDef::CMNeedtoComputeWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating Webcut"


    If Not pObject Is Nothing Then
    
        Dim oWebCut As New StructDetailObjects.WebCut
        Set oWebCut.object = pObject
        Dim oSO As IJSmartOccurrence
        Dim oSI As IJSmartItem
     
        Set oSO = pObject
        Set oSI = oSO.ItemObject
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim oBoundedPort As IJPort
        Dim oBoundingPort As IJPort
        Dim lStatus As Long
        Dim oAppConnection As IJAppConnection
        Set oAppConnection = pPropertyDescriptions.CAO
        ' ---------------------
        ' Get the Bounded Port
        ' ---------------------
        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
        Set oBoundedPort = oBoundedData.AxisPort
     
        If Not oSI Is Nothing Then
            ' Pass the appropriate bounding port
            ' Replace the existing  Webcut if any with new Webcut with updated bounding port
            Dim oMappedPortColl As JCmnShp_CollectionAlias
            GetMultiBoundingEdgeMap pPropertyDescriptions.CAO, WebCut, oMappedPortColl

            Set oBoundingPort = oMappedPortColl.Item(1)
            If Not oBoundingPort Is oWebCut.BoundingPort Then
                'Need to Update the endcut with updated bounding port
                ' Following method will update the endcut with the passed bouning port by replacing the existing end cut
                CMNeedToCompute oBoundingPort, oBoundedPort, pObject
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the Top flangecut based on previous and updated bounding port variation
' *******************************************************************************************
Public Sub CMNeedtoComputeTopFlangeCut(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "GenMultiBoundingDef::CMNeedtoComputeTopFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating Top Flangecut"


    If Not pObject Is Nothing Then
    
        Dim oFlangeCut As New StructDetailObjects.FlangeCut
        Set oFlangeCut.object = pObject
        Dim oSO As IJSmartOccurrence
        Dim oSI As IJSmartItem
     
        Set oSO = pObject
        Set oSI = oSO.ItemObject
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim oBoundedPort As IJPort
        Dim oBoundingPort As IJPort
        Dim lStatus As Long
        Dim strEndCutSel As String
        Dim oAppConnection As IJAppConnection
        Set oAppConnection = pPropertyDescriptions.CAO
        ' ---------------------
        ' Get the Bounded Port
        ' ---------------------
    
        If Not oSI Is Nothing Then
            InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
            Set oBoundedPort = oBoundedData.AxisPort
            
            ' Pass the appropriate bounding port
            ' Replace the existing  flangecut if any with new flangecut with updated bounding port
            Dim oMappedPortColl As JCmnShp_CollectionAlias
            GetMultiBoundingEdgeMap pPropertyDescriptions.CAO, FlangeCutTop, oMappedPortColl
                            
            Set oBoundingPort = oMappedPortColl.Item(1)
            If Not oBoundingPort Is oFlangeCut.BoundingPort Then
                'Need to Update the endcut with updated bounding port
                ' Following method will update the endcut with the passed bouning port by replacing the existing end cut
                CMNeedToCompute oBoundingPort, oBoundedPort, pObject
            End If
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'******************************************************************************************
' Method:
' CMNeedtoComputeBottomFlangeCut
'
' Description: Updates the Bottom flangecut based on previous and updated bounding port variation
' *******************************************************************************************
Public Sub CMNeedtoComputeBottomFlangeCut(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "GenMultiBoundingDef::CMNeedtoComputeBottomFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating Btm Flangecut"


    If Not pObject Is Nothing Then
    
        Dim oFlangeCut As New StructDetailObjects.FlangeCut
        Set oFlangeCut.object = pObject
        Dim oSO As IJSmartOccurrence
        Dim oSI As IJSmartItem
        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim lStatus As Long
     
        Set oSO = pObject
        Set oSI = oSO.ItemObject
        If Not oSI Is Nothing Then
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            Dim oACObject As Object
            AssemblyConnection_SmartItemName pPropertyDescriptions.CAO, , oACObject
            InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg
            Set oBoundedPort = oBoundedData.AxisPort
            
            ' Pass the appropriate bounding port
            ' Replace the existing flangecut if any with new flangecut with updated bounding port
            Dim oMappedPortColl As JCmnShp_CollectionAlias
            GetMultiBoundingEdgeMap pPropertyDescriptions.CAO, FlangeCutBottom, oMappedPortColl
            
            
            Set oBoundingPort = oMappedPortColl.Item(1)
            If Not oBoundingPort Is oFlangeCut.BoundingPort Then
                'Need to Update the endcut with updated bounding port
                ' Following method will update the endcut with the passed bouning port by replacing the existing end cut
                CMNeedToCompute oBoundingPort, oBoundedPort, pObject
            End If
        End If
    
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub


'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the respective endcuts
' *******************************************************************************************
Private Sub CMNeedToCompute(oBoundingPort As Object, oBoundedPort As Object, pObject As Object)
    Const METHOD = "GenMultiBoundingDef::CMNeedToCompute"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating Top/Btm Flangecut"

    Dim lStatus As Long
    Dim strEndCutSel As String
    Dim eStructFeatureType As StructFeatureTypes
    Dim oStructFeature As IJStructFeature
    If TypeOf pObject Is IJStructFeature Then
        Set oStructFeature = pObject
        eStructFeatureType = oStructFeature.get_StructFeatureType
        Dim ParentSystem As IJSystem
        Dim oSD_WebCut As New StructDetailObjects.WebCut
        Dim oSD_FlangeCut As New StructDetailObjects.FlangeCut
        Dim oFeatureUtils As New GSCADSDCreateModifyUtilities.SDFeatureUtils
        If eStructFeatureType = SF_WebCut Then
            Set oSD_WebCut.object = pObject
            strEndCutSel = "GenWebMultiBoundingSel"
            ' Existing Webcut Object ---- oSD_WebCut.object
            ' Below method will replace the existing endcut and creates with new bounding port passed
            Set pObject = oFeatureUtils.CreateWebCut(GetResourceMgr(), oBoundingPort, _
                                oBoundedPort, strEndCutSel, ParentSystem, oSD_WebCut.object)
        Else
            Set oSD_FlangeCut.object = pObject
            strEndCutSel = "GenFlgMultiBoundingSel"
            ' Existing Flangecut Object ---- oSD_FlangeCut.object
            ' Below method will replace the existing endcut and creates with new bounding port passed
            Set pObject = oFeatureUtils.CreateFlangeCut(GetResourceMgr(), oBoundingPort, oBoundedPort, _
                                            oSD_FlangeCut.WebCut, strEndCutSel, ParentSystem, oSD_FlangeCut.object)
            
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

