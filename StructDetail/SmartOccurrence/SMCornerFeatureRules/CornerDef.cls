VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : CornerDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'   Item definition: basic template to define an item
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature

Const m_ItemProgid As String = CUSTOMERID + "CornerFeatRules.CornerDef"
Const m_ItemName As String = CUSTOMERID + "CornerFeatRules.CornerDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\CornerDef.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description :
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    pAD.SetCMFinalConstruct CUSTOMERID + "CornerFeatRules.CornerDef", "CMFinalConstructCF"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
 
End Sub



'*********************************************************************************************
' Method      : ItemMembers
' Description :
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    ' TODO - Add member description below
    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("CornerFeatureFET1", 1, "CMConstructFET1", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMGetCornerTreatment"
    Set pMemDesc = Nothing
   
    Set pMemDesc = pMDs.AddMember("CornerFeatureCollar", 2, "CMConstructCornerCollar", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMGetCornerCollar"
    Set pMemDesc = Nothing
 
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
 
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
'*********************************************************************************************
' Method      : CMFinalConstructCF
' Description :
'
'*********************************************************************************************
Public Sub CMFinalConstructCF(ByVal pAggregatorDescription As IJDAggregatorDescription)


  Dim oCornerFeature As IJSDOCornerFeature
  
  Set oCornerFeature = New StructDetailObjectsex.CornerFeature
  Set oCornerFeature.object = pAggregatorDescription.CAO
  
  oCornerFeature.AddFeatureToCut
    
End Sub
'*********************************************************************************************
' Method      : CMConstructFET1
' Description :
'
'*********************************************************************************************
Public Sub CMConstructFET1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    On Error GoTo ErrorHandler
    Set pObject = CMConstructFETForFeature(pMemberDescription, pResourceManager)
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMConstructFET1").Number
 
End Sub
'*********************************************************************************************
' Method      : CMGetCornerTreatment
' Description :
'
'*********************************************************************************************
Public Sub CMGetCornerTreatment(ByRef pMD As IJDMemberDescription, _
                          ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
    bIsNeeded = False
    
    Dim sApplyTreatment As String
'    sApplyTreatment = GetAnswer_ApplyTreatment(pMD, SMARTTYPE_CORNERFEATURE)
    
    GetSelectorAnswer pMD.CAO, gsApplyTreatment, sApplyTreatment
    
    If LCase(sApplyTreatment) = LCase("No") Then
        bIsNeeded = False
    ElseIf LCase(sApplyTreatment) = LCase("Yes") Then
        bIsNeeded = True
    End If
    
    Dim oCF As IJSDOCornerFeature
  
    Set oCF = New StructDetailObjectsex.CornerFeature
    Set oCF.object = pMD.CAO
    
    
    UpdateACsForEdgePortsOfCF oCF.object
    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMGetCornerTreatment").Number
 
End Sub

'*********************************************************************************************
' Method      : CMConstructCornerCollar
' Description :
'
'*********************************************************************************************
Public Sub CMConstructCornerCollar(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
On Error GoTo ErrorHandler
    
    Set pObject = CreateCollarAtCorner(pMemberDescription, pResourceManager)
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMConstructCornerCollar").Number
 
End Sub
'*********************************************************************************************
' Method      : CMGetCornerCollar
' Description :
'
'*********************************************************************************************
Public Sub CMGetCornerCollar(ByRef pMD As IJDMemberDescription, _
                          ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
    bIsNeeded = False
    
    'As Bulkload is optional for End Users while migration, thereby need
    'to check whether particular functionality exist and then procedd further
    'otherwise exit right away.
    If Not IsCornerCollarFunctionalityBulkloaded() Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sSmartItem As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Set oSmartOccurrence = pMD.CAO
    Set oSmartItem = oSmartOccurrence.SmartItemObject
    sSmartItem = oSmartItem.Name

    If InStr(LCase(sSmartItem), LCase("WithCollar")) > 0 Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMGetCornerCollar").Number
 
End Sub


' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
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
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : ReportError
' Description :
'
'*********************************************************************************************
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

