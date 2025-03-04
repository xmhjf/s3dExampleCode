VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenTopAndBtmEdgeWebDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenTopAndBtmEdgeWebDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    18/Oct/2011 - pnalugol - To address 201305: conditionals to create Top/Btm flange cuts modified.
'    03/August/2015 - MDT TR-274273 Change SM orientation from vertical to horizontal- deletes PCs under GenMemACs
'*********************************************************************************************
Const m_sClassName As String = "GenTopAndBtmEdgeWebDef"
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
    
    Dim sError As String
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    sError = "Adding Items"
    
    'Items for the Physical Connections to Each Bounding Edge
    Set oMemDesc = pMDs.AddMember("TopPC", 1, "CM_ConstructTopPC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("TopEdgePC", 2, "CM_ConstructTopEdgePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopEdgePCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("TopInsidePC", 3, "CM_ConstructTopInsidePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopInsidePCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("FacePC", 4, "CM_ConstructFacePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFacePCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomInsidePC", 5, "CM_ConstructBottomInsidePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomInsidePCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomEdgePC", 6, "CM_ConstructBottomEdgePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomEdgePCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomPC", 7, "CM_ConstructBottomPC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CMNeedToComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Items for the Top and Bottom Web and Flange Cuts
    Set oMemDesc = pMDs.AddMember("TopWebCut", 8, "CM_ConstructTopWebCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopWebCutNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomWebCut", 9, "CM_ConstructBottomWebCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomWebCutNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("TopFlangeCut", 10, "CM_ConstructTopFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "SetTopFlangeInputs"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomFlangeCut", 11, "CM_ConstructBottomFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "SetBtmFlangeInputs"
    Set oMemDesc = Nothing

    'Items for the Corner Features
    Set oMemDesc = pMDs.AddMember("TopEdgeOutsideCorner", 12, "CM_ConstructTopEdgeCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopEdgeOutsideCornerNeeded"
    Set oMemDesc = Nothing

    Set oMemDesc = pMDs.AddMember("TopEdgeInsideCorner", 13, "CM_ConstructTopEdgeInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopEdgeInsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("FaceTopInsideCorner", 14, "CM_ConstructFaceTopInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFaceTopInsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("FaceBottomInsideCorner", 15, "CM_ConstructFaceBottomInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFaceBottomInsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomEdgeInsideCorner", 16, "CM_ConstructBottomEdgeInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomEdgeInsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomEdgeOutsideCorner", 17, "CM_ConstructBottomEdgeCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomEdgeOutsideCornerNeeded"
    Set oMemDesc = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
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

'*********************************************************************************************
' Method      : SetTopFlangeInputs
' Description : Set the top cut as the input, if it exists.  Otherwise,
'               set the center cut as the input
'
'*********************************************************************************************
Public Sub SetTopFlangeInputs(oMD As IJDMemberDescription)

    Const METHOD = "GenTopEdgeDef::SetTopFlangeInputs"
    
    On Error GoTo ErrorHandler
    
    ' ----------------------------
    ' Update the flange cut inputs
    ' ----------------------------
    Dim oDefCM As New EndCutDefCM
    oDefCM.UpdateFlangeCutInputByDispID 8, oMD.CAO, oMD.object

    ' ---------------------
    ' Copy question answers
    ' ---------------------
    oDefCM.CM_CopyQuestionsToTopFlangeCut oMD
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : SetBtmFlangeInputs
' Description : Set the bottom cut as the input, if it exists.  Otherwise,
'               set the center cut as the input
'
'*********************************************************************************************
Public Sub SetBtmFlangeInputs(oMD As IJDMemberDescription)

    Const METHOD = "GenBtmEdgeDef::SetBtmFlangeInputs"
    
    On Error GoTo ErrorHandler
    
    ' ----------------------------
    ' Update the flange cut inputs
    ' ----------------------------
    Dim oDefCM As New EndCutDefCM
    oDefCM.UpdateFlangeCutInputByDispID 9, oMD.CAO, oMD.object

    ' ---------------------
    ' Copy question answers
    ' ---------------------
    oDefCM.CM_CopyQuestionsToBottomFlangeCut oMD
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub


'*********************************************************************************************
' Method      : CM_IsTopFlangeCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sTopFlangeCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateTopFlangeCut", sTopFlangeCutNeeded

    If sTopFlangeCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sTopFlangeCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    If bIsNeeded Then
        Dim oMemberObjects As IJDMemberObjects
        Set oMemberObjects = pMemberDescription.CAO

        If Not oMemberObjects.ItemByDispid(8) Is Nothing Then
            'Top Web Cut exist
             bIsNeeded = False
             Exit Sub
        Else
            bIsNeeded = True
        End If
    End If

    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub



'*********************************************************************************************
' Method      : CM_IsBottomFlangeCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sBottomFlangeCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateBottomFlangeCut", sBottomFlangeCutNeeded

    If sBottomFlangeCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sBottomFlangeCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
        
    If bIsNeeded Then
        Dim oMemberObjects As IJDMemberObjects
        Set oMemberObjects = pMemberDescription.CAO

        If Not oMemberObjects.ItemByDispid(9) Is Nothing Then
            'btm Web Cut exist
             bIsNeeded = False
             Exit Sub
        Else
            bIsNeeded = True
        End If
    End If

    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub
'******************************************************************************************
' Method:
' CMNeedtoComputePC
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************
Public Sub CMNeedtoComputePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "CMNeedToComputePC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
          
    Dim oEndCutDefCM As New EndCutDefCM
    oEndCutDefCM.CM_NeedToComputePenPC pPropertyDescriptions, pObject, True, True
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
