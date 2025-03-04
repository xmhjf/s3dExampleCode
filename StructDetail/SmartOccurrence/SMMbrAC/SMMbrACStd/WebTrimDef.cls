VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebTrimDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrACStd
'  File        : WebTrimDef.cls
'
'  Description :
'
'  Author      : Alligators
'
' 12/May/15   GH    CR-260982 Added Property descriptions to all the Physical connection Item members
'                   Created new method CMNeedtoComputePC() to handle smart item change.
'*********************************************************************************************
Const m_sClassName As String = "WebTrimDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

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
    sMsg = "Defining Definition Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
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
    
    pAD.UserTypeClsid = CA_WEBCUT
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
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
    
    pMDs.RemoveAll  ' Remove all the previous cached member descriptions(best practice)
        
    '******************************************************************************
    'if any more item members are needed to be added please add only at the bottom
    '******************************************************************************
        
    ' Creates a Physical Connection with the bounding (mapped) top port
    lDispId = 1
    sDispId = "Std_EndCut_TopPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right top port
    lDispId = 2
    sDispId = "Std_EndCut_FaceTopPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right top port
    lDispId = 3
    sDispId = "Std_EndCut_TopEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right port
    lDispId = 4
    sDispId = "Std_EndCut_TopEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right bottom port
    lDispId = 5
    sDispId = "Std_EndCut_TopEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right port
    lDispId = 6
    sDispId = "Std_EndCut_FacePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeFacePC", 1, "IJSmartOccurrence", "CMNeedtoComputeForFacePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right top
    lDispId = 7
    sDispId = "Std_EndCut_BottomEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right port
    lDispId = 8
    sDispId = "Std_EndCut_BottomEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right bottom port
    lDispId = 9
    sDispId = "Std_EndCut_BottomEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right bottom port
    lDispId = 10
    sDispId = "Std_EndCut_FaceBottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom port
    lDispId = 11
    sDispId = "Std_EndCut_BottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top port
    lDispId = 12
    sDispId = "Std_EndCut_TopPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right top port
    lDispId = 13
    sDispId = "Std_EndCut_FaceTopPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right top port
    lDispId = 14
    sDispId = "Std_EndCut_TopEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right port
    lDispId = 15
    sDispId = "Std_EndCut_TopEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right bottom port
    lDispId = 16
    sDispId = "Std_EndCut_TopEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right port
    lDispId = 17
    sDispId = "Std_EndCut_FacePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right top
    lDispId = 18
    sDispId = "Std_EndCut_BottomEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right port
    lDispId = 19
    sDispId = "Std_EndCut_BottomEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right bottom port
    lDispId = 20
    sDispId = "Std_EndCut_BottomEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right bottom port
    lDispId = 21
    sDispId = "Std_EndCut_FaceBottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom port
    lDispId = 22
    sDispId = "Std_EndCut_BottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    
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

Public Sub CMNeedtoComputeForPC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object, _
                                            eBoundingSubPort As JXSEC_CODE, eBoundedSubPort As JXSEC_CODE)
                                            
    Const METHOD = "WebTrimDef::CMNeedtoComputeForPC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
            
     If TypeOf pObject Is IJStructPhysicalConnection Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject

        Dim oACDef As New ACDef
        Dim oAppConnection As IJAppConnection
        AssemblyConnection_SmartItemName pPropertyDescriptions.CAO, , oAppConnection
                    
        If Not oSmartItem Is Nothing And oACDef.IsBoxCut(oAppConnection) Then
                     
            Dim oSD_PhysicalConn As New StructDetailObjects.PhysicalConn
            Set oSD_PhysicalConn.object = pObject
            
            Dim oSD_Webcut As New StructDetailObjects.WebCut
            Set oSD_Webcut.object = pPropertyDescriptions.CAO
            
            Dim oBoundedObject As Object
            Set oBoundedObject = oSD_Webcut.Bounded
            
            If oBoundedObject Is oSD_PhysicalConn.ConnectedObject1 Then
                Exit Sub
            End If
                    
            ' --------------------
            ' Get the edge mapping
            ' --------------------
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Set oEdgeMapColl = GetEdgeMap(pPropertyDescriptions.CAO, oSD_Webcut.BoundingPort, oSD_Webcut.BoundedPort)
            
            eBoundingSubPort = oEdgeMapColl.Item(CStr(eBoundingSubPort))
            
            ' -------------------------------
            ' Get the PC selection and filter
            ' -------------------------------
            Dim strSelection As String
            Dim strFilter As String
            
            Dim oBoundingPort As IJPort
            Dim pHelper As New StructDetailObjects.Helper
    
            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSD_Webcut.Bounding, eBoundingSubPort)
            Set oBoundingPort = pHelper.GetEquivalentLastPort(oBoundingPort)
    
            Dim oStructProfilePart As IJStructProfilePart
            Dim oStructEndCutUtil As IJStructEndCutUtil
            Dim oBoundedPort As IJPort
            
            Set oStructProfilePart = oSD_Webcut.Bounded
            Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
            oStructEndCutUtil.GetLatePortForFeatureSegment oSD_Webcut.object, eBoundedSubPort, oBoundedPort
        
            Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
            Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
            
            oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysicalConn.Port1, oBoundedPort
            oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysicalConn.Port2, oBoundingPort
        ElseIf Not oSmartItem Is Nothing Then
            'Update the PC if smart item is changed
            CMNeedtoComputePC pPropertyDescriptions, pObject
            
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CMNeedtoComputeForFacePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "WebTrimDef::CMNeedtoComputeForFacePC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the FacePC "
    
    Dim eBoundingSubPort As JXSEC_CODE
    Dim eBoundedSubPort As JXSEC_CODE
    
    If pObject Is Nothing Then Exit Sub
    
    eBoundedSubPort = JXSEC_WEB_RIGHT
    eBoundingSubPort = JXSEC_WEB_RIGHT

    CMNeedtoComputeForPC pPropertyDescriptions, pObject, eBoundingSubPort, eBoundedSubPort

  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


Public Sub CMNeedtoComputePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "WebTrimDef::CMNeedtoComputePC"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Modifying the PC "
    
    If pObject Is Nothing Then Exit Sub
    
    Dim oSmartOccurance As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    
    Set oSmartOccurance = pObject
    Set oSmartItem = oSmartOccurance.ItemObject
    
    'Exit while creating the PC
    If oSmartItem Is Nothing Then
        Exit Sub
    End If
    
    'Check object Type
    If TypeOf pObject Is IJStructPhysicalConnection Then
    
        Dim oMemberDesc As IJDMemberDescription
        Set oMemberDesc = GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions, pObject)
        
        If Not oMemberDesc Is Nothing Then
            
            Dim eBoundingSubPort As JXSEC_CODE
            Dim eBoundedEdgeSubPort As JXSEC_CODE
            
            Dim strACAttribute As String
            Dim strPCSelection As String
            Dim oCM As New CustomMethods
            Dim strNewFilter As String
            
            'Get PC attribute Name
            strACAttribute = oCM.GetPCAttributeForCenterCut(oMemberDesc, eBoundingSubPort, eBoundedEdgeSubPort)
            
            '**********************************
            
            Dim oPrevBoundingPort As IJStructPort
            Dim oPrevBoundedPort As IJStructPort
            Dim oCurrentBoundingPort As IJPort
            Dim oCurrentBoundedPort As IJPort
            Dim oSD_PhysConn As New StructDetailObjects.PhysicalConn
            Set oSD_PhysConn.object = pObject
            Set oPrevBoundingPort = oSD_PhysConn.Port2
            Set oPrevBoundedPort = oSD_PhysConn.Port1
            Dim oPCAsChild As IJDesignChild
            Set oPCAsChild = pObject
            Dim oWCObj As Object
            Set oWCObj = oPCAsChild.GetParent
            Dim oSD_Webcut As New StructDetailObjects.WebCut
            Set oSD_Webcut.object = oWCObj
            
            Dim oStructProfilePart As IJStructProfilePart
            Dim oStructEndCutUtil As IJStructEndCutUtil
            
            Dim pHelper As New StructDetailObjects.Helper
            Dim oBoundingPort As IJPort
            Set oBoundingPort = oPrevBoundingPort
            
            Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
            
            If Not oPrevBoundingPort.SectionID = eBoundingSubPort Then
            
                Set oCurrentBoundingPort = GetLateralSubPortBeforeTrim(oSD_Webcut.Bounding, eBoundingSubPort)
                Set oBoundingPort = pHelper.GetEquivalentLastPort(oCurrentBoundingPort)
            
                Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysConn.Port2, oBoundingPort
            End If
            
            If Not oPrevBoundedPort.SectionID = eBoundedEdgeSubPort Then
                Set oStructProfilePart = oSD_Webcut.Bounded
                Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                oStructEndCutUtil.GetLatePortForFeatureSegment oSD_Webcut.object, eBoundedEdgeSubPort, oCurrentBoundedPort
                
                Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysConn.Port1, oCurrentBoundedPort
            
            End If
            
            '**********************************
            'Get PC Filter
            oCM.GetPCSelection oMemberDesc, strACAttribute, strPCSelection, strNewFilter
            
            'Update PC
            UpdatePCWithNewFilterProgID pObject, strNewFilter
            
        End If
    End If
    
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

