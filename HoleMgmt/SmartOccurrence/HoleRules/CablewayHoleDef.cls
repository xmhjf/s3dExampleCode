VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CablewayHoleDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurrence\HoleRules
'
' File: CablewayHoleDef.cls
'
' Author: Hole Mgmt Team
'
' Abstract: definition for cableway hole trace custom assembly
'
' Description: This is the definition for the cableway hole trace custom assembly.
'              This class will be triggered if the outfitting for the hole trace
'              is a single cableway.
'******************************************************************************

Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgId
Const m_CAProgid As String = "HoleRules.CablewayHoleDef"
Const m_FamilyProgid As String = ""

'Error management
Private sERROR As String
Private sMETHOD As String
Private Const MODULE = "M:\SharedContent\Src\HoleMgmt\SmartOccurrence\HoleRules\CablewayHoleDef.cls"

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    'the hole trace is the only input
    pIH.SetInput INPUT_HOLETRACE
  
    Exit Sub

ErrorHandler:
    pIH.ReportError "ItemInputs"
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
  
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler

    sMETHOD = "ItemMembers"
    
'''    Dim pMemDesc As IJDMemberDescription
'''    Dim pPDs As IJDPropertyDescriptions
'''
'''    'Create MemberDesc and add conditional to check if the hole is approved before cutting plate
'''    sERROR = "Constructing hole feature"
'''    Set pMemDesc = pMDs.AddMember("HoleFeature", HOLE_FEATURE, "CMConstruct_HoleFeature", LIBRARY_SOURCE_ID)
'''    pMemDesc.SetCMConditional LIBRARY_SOURCE_ID, "CM_IsApproved"
'''    Set pMemDesc = Nothing
'''
'''    'Create MemberDesc and add conditional to check if the hole requires edge reinforcement
'''    sERROR = "Constructing edge reinforcement"
'''    Set pMemDesc = pMDs.AddMember("EdgeReinforcement", EDGE_REINFORCEMENT, "CMConstruct_EdgeReinforcement", LIBRARY_SOURCE_ID)
'''    pMemDesc.SetCMConditional LIBRARY_SOURCE_ID, "CM_IsCorrectFitting"
'''    Set pMemDesc = Nothing
'''
'''    'Create MemberDesc and add conditional to check if the hole requires catalog fitting
'''    sERROR = "Constructing catalog fitting"
'''    Set pMemDesc = pMDs.AddMember("CatalogFitting", CATALOG_FITTING, "CMConstruct_CatalogFitting", LIBRARY_SOURCE_ID)
'''    pMemDesc.SetCMConditional LIBRARY_SOURCE_ID, "CM_IsCorrectFitting"
'''    ' Set the part properly
'''    Set pPDs = pMemDesc
'''    sERROR = "Setting catalog fitting property"
'''    pPDs.AddProperty "BestFitPart", 1, IID_IJPartOcc, "CMUpdate_CWayPartOcc", LIBRARY_SOURCE_ID
'''    Set pMemDesc = Nothing
'''    Set pPDs = Nothing
'''
'''    'Create MemberDesc and add conditional to check if the hole requires profile fitting
'''    sERROR = "Constructing profile fitting"
'''    Set pMemDesc = pMDs.AddMember("ProfileFitting", PROFILE_FITTING, "CMConstruct_ProfileFitting", LIBRARY_SOURCE_ID)
'''    pMemDesc.SetCMConditional LIBRARY_SOURCE_ID, "CM_IsCorrectFitting"
'''    Set pMemDesc = Nothing
'''
'''    'Create MemberDesc and add conditional to check if the hole requires physical connection
'''    sERROR = "Constructing physical connection"
'''    Set pMemDesc = pMDs.AddMember("PhysicalConn", PHYSICAL_CONNECTION, "CMConstruct_PhysicalConn", LIBRARY_SOURCE_ID)
'''    pMemDesc.SetCMConditional LIBRARY_SOURCE_ID, "CM_IsCorrectFitting"
'''    Set pMemDesc = Nothing
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_CAProgid
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler

    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
  
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
  
    ' Set definition progId and codebase
    pDefinition.ProgId = m_CAProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
    
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
    Exit Function
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJDUserSymbolServices_InstanciateDefinition").Number
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
    On Error GoTo ErrorHandler
    
    sMETHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    sERROR = ""
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation
    
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
  
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    
    ' To fix the TR#206747, we made this change. This is a temporary fix only. There is some issue with SymbolDefinition code as per TR# 206916.
    ' After fix of TR#206916, we can provide better solution and will make further changes here.
    
    On Error Resume Next
    Set pAD = pDefinition
    
    If pAD Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler
        
    ItemAggregator pAD
   
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    ItemMembers pMDs

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function
 
