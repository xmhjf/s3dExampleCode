VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeTrimDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  File        : FlangeTrimDef.cls
'
'  Description :
'
'  Author      : Alligators
'
' 12/May/15   GH    CR-260982 Added Property descriptions to all the Physical connection Item members to handle smart item change.
'
'*********************************************************************************************
Const m_sClassName As String = "FlangeTrimDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

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
Public Sub ItemAggregator(oAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    oAD.UserTypeClsid = CA_FLANGECUT       '
    oAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    oAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    oAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Dim oPDs As IJDPropertyDescriptions
    Set oPDs = oAD
    oPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    oPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
        
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
        
    ' Creates a Physical Connection with the bounding (mapped) top port
    lDispId = 1
    sDispId = "Std_EndCut_FacePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) web right top port
    lDispId = 2
    sDispId = "Std_EndCut_TopEdgePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) top flange right top port
    lDispId = 3
    sDispId = "Std_EndCut_BottomEdgePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right port
    lDispId = 4
    sDispId = "Std_EndCut_TopEdgeInsidePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right bottom port
    lDispId = 5
    sDispId = "Std_EndCut_BottomEdgeInsidePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
            
    ' Creates a Physical Connection with the bounding (mapped) web right port
    lDispId = 6
    sDispId = "Std_EndCut_TopPC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right top
    lDispId = 7
    sDispId = "Std_EndCut_BottomPC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right port
    lDispId = 8
    sDispId = "Std_EndCut_TopEdgeOutsidePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right bottom port
    lDispId = 9
    sDispId = "Std_EndCut_BottomEdgeOutsidePC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right bottom port
    lDispId = 10
    sDispId = "Std_EndCut_FaceTopPC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) bottom port
    lDispId = 11
    sDispId = "Std_EndCut_FaceBottomPC_LT_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top port
    lDispId = 12
    sDispId = "Std_EndCut_FacePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) web right top port
    lDispId = 13
    sDispId = "Std_EndCut_TopEdgePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) top flange right top port
    lDispId = 14
    sDispId = "Std_EndCut_BottomEdgePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right port
    lDispId = 15
    sDispId = "Std_EndCut_TopEdgeInsidePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right bottom port
    lDispId = 16
    sDispId = "Std_EndCut_BottomEdgeInsidePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
            
    ' Creates a Physical Connection with the bounding (mapped) web right port
    lDispId = 17
    sDispId = "Std_EndCut_TopPC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right top
    lDispId = 18
    sDispId = "Std_EndCut_BottomPC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right port
    lDispId = 19
    sDispId = "Std_EndCut_TopEdgeOutsidePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right bottom port
    lDispId = 20
    sDispId = "Std_EndCut_BottomEdgeOutsidePC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right bottom port
    lDispId = 21
    sDispId = "Std_EndCut_FaceTopPC_RB_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    ' Creates a Physical Connection with the bounding (mapped) bottom port
    lDispId = 22
    sDispId = "Std_EndCut_FaceBottomPC_RB_" & Trim(Str(lDispId))
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

'******************************************************************************************
' Method:
' SetFlangeCuttingBehaviour
' *******************************************************************************************
Public Sub SetFlangeCuttingBehaviour(oPD As IJDPropertyDescription, pObject As Object)
    
    Const METHOD = "SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    ' ----------------------------------------------
    ' Exit if the behavior attribute does not exists
    ' ----------------------------------------------
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = oPD.definition
    
   
    Dim IsProExists As Boolean
    IsProExists = Has_Attribute(oPD.CAO, "CuttingBehavior")
    
    If Not IsProExists Then
        Exit Sub
    End If
        
    ' -------------------------------
    ' Determine if there will be a PC
    ' -------------------------------
    Dim bIsNeeded As Boolean
    Dim i As Integer
    Dim oCM As New CustomMethods
    Dim oMemberDescription As IJDMemberDescription
    
    bIsNeeded = True
    
    For i = 1 To 22
        
        Set oMemberDescription = oMemberDescriptions.ItemByDispid(i)
        oMemberDescription.CAO = oPD.CAO
    
        oCM.CM_CenterCutPCNeeded oMemberDescription, bIsNeeded
        
        If bIsNeeded Then
            Exit For
        End If
    
    Next

    ' ----------------------------------------------------------------------------------
    ' Get proper flange cutting behaviour based on existence of PC and flange cut inputs
    ' ----------------------------------------------------------------------------------
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
  
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


Public Sub CMNeedtoComputePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "FlangeTrimDef::CMNeedtoComputePC"
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
            Dim oSD_Flangecut As New StructDetailObjects.WebCut
            Set oSD_Flangecut.object = oWCObj
            
            Dim oStructProfilePart As IJStructProfilePart
            Dim oStructEndCutUtil As IJStructEndCutUtil
            
            Dim pHelper As New StructDetailObjects.Helper
            Dim oBoundingPort As IJPort
            Set oBoundingPort = oPrevBoundingPort
            
            If Not oPrevBoundingPort.SectionID = eBoundingSubPort Then
                Set oCurrentBoundingPort = GetLateralSubPortBeforeTrim(oSD_Flangecut.Bounding, eBoundingSubPort)
                Set oBoundingPort = pHelper.GetEquivalentLastPort(oCurrentBoundingPort)
            End If
            
            If Not oPrevBoundedPort.SectionID = eBoundedEdgeSubPort Then
                Set oStructProfilePart = oSD_Flangecut.Bounded
                Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                oStructEndCutUtil.GetLatePortForFeatureSegment oSD_Flangecut.object, eBoundedEdgeSubPort, oCurrentBoundedPort
            End If
            
            If Not (oPrevBoundingPort.SectionID = eBoundingSubPort) Or Not (oPrevBoundedPort.SectionID = eBoundedEdgeSubPort) Then
                Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
                Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
            
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysConn.Port1, oCurrentBoundedPort
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysConn.Port2, oBoundingPort
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


