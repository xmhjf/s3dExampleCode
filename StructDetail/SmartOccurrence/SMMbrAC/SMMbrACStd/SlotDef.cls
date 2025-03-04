VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  File        : SlotDef.cls
'
'  Description : This creates Slot Physical Connection
'
'  Author      : Alligators
'
'  History     :
'
'  15/12/2015  mkonduri TR-276387 Added a method GetEquivalentLastPort() to obtain port after
'                                  ladder rung extends.
'
'*********************************************************************************************
Implements IJDUserSymbolServices

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotDef"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sStdACProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Private m_oPhysConn As Object

'Error management
Private sError As String
Private sMETHOD As String
'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_BOUNDING
  pIH.SetInput INPUT_BOUNDED
  
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
  
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructSlot"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'*********************************************************************************************
Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    Dim lDispId As Long
    Dim sDispId As String
    Dim oMemDesc As IJDMemberDescription
    lDispId = 1
    sDispId = "Slot_PC" & Trim(Str(lDispId))
    Set oMemDesc = oMDs.AddMember(sDispId, lDispId, "CMConstructSlotPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oMemDesc = Nothing
    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub
'******************************************************************************************
' Method:
' CMConstructSlotPC
' *******************************************************************************************
Public Sub CMConstructSlotPC(ByVal oMemberDescription As IJDMemberDescription, _
                             ByVal oResourceManager As IUnknown, _
                             ByRef oObject As Object)
                             
    Const sMETHOD = "CMConstructSlotPC"
    Dim sError As String
    Dim m_oPhysConn As Object
    Dim oACObject As IJAppConnection
    Dim oSlot As New StructDetailObjects.Slot
    Dim oPenetratingMemberPart As New StructDetailObjects.MemberPart
    Dim oPenetratedMemberPart As New StructDetailObjectsex.SDOMemberPart
    Dim oPenetratingPortForPC As IJPort
    Dim oPenetratingPortForPCBeforeTrim As IJPort
    Dim oPenetratedPortForPC As IJPort
    Dim oSD_Helper As New StructDetailObjects.Helper
    
    AssemblyConnection_SmartItemName oMemberDescription.CAO, , oACObject
    Set oSlot.object = oMemberDescription.CAO
    
    Dim oACDef As New ACDef
    Dim strSmartClass As String
    strSmartClass = oACDef.GetStandardACAttribute(oACObject, "SlotPC")
    sError = "Setting Slot Inputs"
    
    ' Passing Global port for construction of Physical conection
    Set oPenetratingMemberPart.object = oSlot.Penetrating
    Set oPenetratingPortForPCBeforeTrim = oPenetratingMemberPart.BasePortBeforeTrim(BPT_Lateral)
    Set oPenetratingPortForPC = oSD_Helper.GetEquivalentLastPort(oPenetratingPortForPCBeforeTrim)
    
    
    Set oPenetratedMemberPart.object = oSlot.Penetrated
    Set oPenetratedPortForPC = oPenetratedMemberPart.CutoutPort(oMemberDescription.CAO)
    
    Dim pSystemParent As IJSystemChild
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = oMemberDescription.CAO
      
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    oPhysicalConnection.Create oResourceManager, _
                                 oPenetratedPortForPC, _
                                 oPenetratingPortForPC, _
                                 strSmartClass, _
                                 pSystemParent, ConnectionStandard
    
    sError = "Setting Physical Connection to private variable"
    Set m_oPhysConn = oPhysicalConnection.object
    
    Set oObject = m_oPhysConn
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructSlot(ByVal pAggregatorDescription As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    sMETHOD = "CMFinalConstructSlot"
    
    ' Create Slot and add slotcut
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = pAggregatorDescription.CAO
    oSlot.AddSlotToCut
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructSlot").Number
End Sub

Public Sub CM_IsPCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = "CM_IsPCNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Setting PC conditional"
    bIsNeeded = False
    Dim oACObject As IJAppConnection
    AssemblyConnection_SmartItemName pMemberDescription.CAO, , oACObject
    Dim oACDef As New ACDef
    Dim strSmartClass As String
    strSmartClass = oACDef.GetStandardACAttribute(oACObject, "SlotPC")
    If Not strSmartClass = "" Then
        bIsNeeded = True
    End If
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_sRuleProgID
End Function
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

  Dim pDefinition As IJDSymbolDefinition
  Dim pCAFactory As New CAFactory
  
  Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_sRuleProgID
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
  sError = ""
  
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
  pIH.InitAs m_sFamilyProgID
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
  
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function


