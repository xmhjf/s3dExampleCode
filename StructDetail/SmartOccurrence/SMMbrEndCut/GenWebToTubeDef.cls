VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebToTubeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenWebToTubeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************
Const m_sClassName As String = "GenWebToTubeDef"
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
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
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
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    'Physical Connection for the Web Cut
    iDispId = 1
    sDispId = "GenericWeb_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPhysConns", CUSTOMERID & "MbrEndCut.MbrAxisWebDef")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oMemDesc = Nothing
    
    'Top web Cut, valid when web penetrated
    iDispId = 2
    'Items for the Top and Bottom Web and Flange Cuts
    Set oMemDesc = pMDs.AddMember("TopWebCut", iDispId, "CM_ConstructTopWebCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopWebCutNeeded"
    Set oMemDesc = Nothing
    
    'Bottom web Cut, valid when web penetrated
    iDispId = 3
    Set oMemDesc = pMDs.AddMember("BottomWebCut", iDispId, "CM_ConstructBottomWebCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomWebCutNeeded"
    Set oMemDesc = Nothing
    
    'Top Flange Cut
    iDispId = 4
    sDispId = "GenericTopFlange_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructTopFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oMemDesc = Nothing
        
    'Bottom Flange Cut
    iDispId = 5
    sDispId = "GenericBottomFlange_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructBottomFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
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
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CM_IsPhysConnsNeeded
' Description :
'*********************************************************************************************
Public Sub CM_IsPhysConnsNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim sError As String

    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectTube", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
 
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
        
        If Not oMemberObjects.ItemByDispid(2) Is Nothing Then
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
        
        If Not oMemberObjects.ItemByDispid(3) Is Nothing Then
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
