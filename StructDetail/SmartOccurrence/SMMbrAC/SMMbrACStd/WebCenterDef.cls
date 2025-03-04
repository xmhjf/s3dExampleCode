VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCenterDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrACStd
'  File        : WebCenterDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'*********************************************************************************************
Const m_sClassName As String = "WebCenterDef"
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
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right top port
    lDispId = 2
    sDispId = "Std_EndCut_FaceTopPC__" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right top port
    lDispId = 3
    sDispId = "Std_EndCut_TopEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right port
    lDispId = 4
    sDispId = "Std_EndCut_TopEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) top flange right bottom port
    lDispId = 5
    sDispId = "Std_EndCut_TopEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right port
    lDispId = 6
    sDispId = "Std_EndCut_FacePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right top
    lDispId = 7
    sDispId = "Std_EndCut_BottomEdgeInsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right port
    lDispId = 8
    sDispId = "Std_EndCut_BottomEdgePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom flange right bottom port
    lDispId = 9
    sDispId = "Std_EndCut_BottomEdgeOutsidePC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) web right bottom port
    lDispId = 10
    sDispId = "Std_EndCut_FaceBottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
    Set oMemDesc = Nothing
    
    ' Creates a Physical Connection with the bounding (mapped) bottom port
    lDispId = 11
    sDispId = "Std_EndCut_BottomPC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructCenterCutPC", CUSTOMERID & "MbrACStd.CustomMethods")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrACStd.CustomMethods", "CM_CenterCutPCNeeded"
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


