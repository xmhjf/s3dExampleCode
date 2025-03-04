VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenBUBtmEdgeFlangeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenBUBtmEdgeFlangeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "GenBUBtmEdgeFlangeDef"
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
    
    pAD.UserTypeClsid = CA_FLANGECUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pAD
    pPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    pPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "CM_SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' List all the Item members that this SmartClass can create
'*********************************************************************************************
' Method      : ItemMembers
' Description :
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim oMemDesc As IJDMemberDescription
    
    sError = "Adding Items"
    
    'Items for the Physical Connections to Each Bounding Edge
    Set oMemDesc = pMDs.AddMember("FacePC", 1, "CM_ConstructFacePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFacePCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("InsidePC", 2, "CM_ConstructBottomInsidePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsInsidePCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("EdgePC", 3, "CM_ConstructBottomEdgePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsEdgePCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("OutsidePC", 4, "CM_ConstructBottomOutsidePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsOutsidePCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("FaceOutsidePC", 5, "CM_ConstructBottomFacePC", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFaceOutsidePCNeeded"
    Set oMemDesc = Nothing
    
    'Items for the Top and Bottom Flange Cuts
    Set oMemDesc = pMDs.AddMember("TopFlangeCut", 6, "CM_ConstructTopFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsFromFlangeCenterCutToTopCut"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("BottomFlangeCut", 7, "CM_ConstructBottomFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsFromFlangeCenterCutToBottomCut"
    Set oMemDesc = Nothing

    'Items for the Corner Features
    Set oMemDesc = pMDs.AddMember("FaceInsideCorner", 8, "CM_ConstructFaceBottomInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFaceInsideCornerNeeded"
    Set oMemDesc = Nothing

    Set oMemDesc = pMDs.AddMember("EdgeInsideCorner", 9, "CM_ConstructBottomEdgeInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsEdgeInsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("EdgeOutsideCorner", 10, "CM_ConstructBottomEdgeOutsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsEdgeOutsideCornerNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = pMDs.AddMember("FaceOutsideCorner", 11, "CM_ConstructFaceBottomOutsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsFaceOutsideCornerNeeded"
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


'******************************************************************************************
' Method:
' CM_SetFlangeCuttingBehaviour
'
' *******************************************************************************************
Public Sub CM_SetFlangeCuttingBehaviour(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CM_SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = pPropertyDescriptions.definition
    
    Dim oMemberDescription As IJDMemberDescription
    Dim oEndCutDefCM As New EndCutDefCM
    Dim iMember As Integer
    Dim bIsNeeded As Boolean
    bIsNeeded = True
    
    Dim IsProExists As Boolean
    
    'Check if the attribute is exists on flange cut
    IsProExists = Has_Attribute(pPropertyDescriptions.CAO, "CuttingBehavior")
    
    'Exit if attribute is not bulkloaded
    If Not IsProExists Then Exit Sub
        
    'Looping for only 1 to 5 items as they are only the PC item members
    'Call each Item member conditional and check weather it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    For iMember = 1 To 5
        
        'Get each Member Description
        Set oMemberDescription = oMemberDescriptions.ItemByDispid(iMember)
        oMemberDescription.CAO = pPropertyDescriptions.CAO
    
        Select Case iMember
        
            Case 1
                oEndCutDefCM.CM_IsFacePCNeeded oMemberDescription, bIsNeeded
            
            Case 2
                oEndCutDefCM.CM_IsInsidePCNeeded oMemberDescription, bIsNeeded
                
            Case 3
                oEndCutDefCM.CM_IsEdgePCNeeded oMemberDescription, bIsNeeded
            
            Case 4
                oEndCutDefCM.CM_IsOutsidePCNeeded oMemberDescription, bIsNeeded
            
            Case 5
                oEndCutDefCM.CM_IsFaceOutsidePCNeeded oMemberDescription, bIsNeeded
                
        End Select
        
        'If any of the PC conditional is True, then Exit
        If bIsNeeded Then Exit For
    
    Next
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub



