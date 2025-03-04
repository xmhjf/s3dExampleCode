VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQB.cls
'   ProgID:         HS_Hilti_MQParts.MQB
'   Author:         JRK
'   Creation Date:  Friday, November 16, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQB"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQB"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices

'********************************************************************
' Function: IJDUserSymbolServices_EditOccurence
' This method is called each time the user wants to edit the symbol
'   occurrence (Modify its parameter values or its representation for
'   instance). Prefer to register a custom command to fulfill this
'   functionality, into the symbol definition.
' Currently obsolete (as of Oct 99).
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

'********************************************************************
' Sub: IJDUserSymbolServices_InitializeSymbolDefinition
' This procedure is called to set up the Inputs and outputs of the Symbol
'
' Parameters: pSymbolDefinition - IMSSymbolEntities.IJDSymbolDefinition
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrHandler

    'Inputs
    Dim Inputs As tInputs
    AddInput Inputs, "ChannelW", "Double"
    AddInput Inputs, "ChannelH", "Double"
    AddInput Inputs, "T", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "CPTo", "Double"
    AddInput Inputs, "ChannelSize", "String"
    AddInput Inputs, "ChannelType", "String"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Channel"
    AddOutput Outputs, "Structure"
    AddOutput Outputs, "Body"
    AddOutput Outputs, "AddOutputs_Done"

    SetupInputsAndOutputsEx pSymbolDefinition, Inputs, Outputs, m_progID

    Exit Sub
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'********************************************************************
' Function: IJDUserSymbolServices_GetDefinitionName
' Return the Definition Name
'
' Parameters: definitionParamaters - Variant
' Returns: String
'********************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_progID
End Function

'********************************************************************
' Function: IJDUserSymbolServices_InstanciateDefinition
' Instantiates a persistent symbol definition object and initializes
'   it for the first time.
'
' Parameters: CodeBase - String, defParamaters - Variant, ActiveConnection - Object
' Returns: Object
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
    Set m_SymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition m_SymbolDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = m_SymbolDefinition
    m_SymbolDefinition.ProgId = m_progID
    m_SymbolDefinition.CodeBase = CodeBase
    m_SymbolDefinition.name = m_SymbolDefinition.ProgId
End Function

'********************************************************************
' Sub: IJDUserSymbolServices_InvokeRepresentation
' It handles the recompute of  the symbol graphics outputs given the
'   specified representation and set of parameters.
'
' Parameters: sblOcc - Object, repName - String, outputcoll - Object, arrayOfInputs()
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    Set oCodeListData = sblOcc
    If StrComp(repName, "Symbolic") = 0 Then
        Call Physical(arrayOfInputs, sblOcc)
    End If
    Set oCodeListData = Nothing
End Sub

'********************************************************************
' Function: Physical
' Handles the recompute of the Physical representation
'
' Parameters: arrayOfInputs()
' Returns: Nothing
'********************************************************************
Private Function Physical(ByRef arrayOfInputs(), ByVal sblOcc As Object)
    Const METHOD = "Physical"
    On Error GoTo ErrHandler

    Dim dChannelW As Double
    Dim dChannelH As Double
    Dim dT As Double
    Dim dW As Double
    Dim dCPTo As Double
    Dim dD As Double

    dChannelW = arrayOfInputs(2)
    dChannelH = arrayOfInputs(3)
    dT = arrayOfInputs(4)
    dW = arrayOfInputs(5)
    dCPTo = arrayOfInputs(6)
    dD = 0.04

    AddPort "Channel", 0, 0, 0, _
            0, 0, 1, _
            1, 0, 0, m_outputColl, "Channel", arrayOfInputs(1), "Channel"
    AddPort "Structure", 0, 0, -dChannelH, _
            0, 0, 1, _
            1, 0, 0, m_outputColl, "Structure", arrayOfInputs(1), "Structure"

    AddPolygon 12, "EXTRUDED", 0, "POLY"
    AddPoint -dD / 2#, -dChannelW / 2#, 0, m_outputColl, "", "POLYPT0"
    AddPoint -dD / 2#, -dChannelW / 2#, -dChannelH, m_outputColl, "", "POLYPT1"
    AddPoint -dD / 2#, -dW / 2#, -dChannelH, m_outputColl, "", "POLYPT2"
    AddPoint -dD / 2#, -dW / 2#, -dChannelH + dT, m_outputColl, "", "POLYPT3"
    AddPoint -dD / 2#, -dChannelW / 2# - dT, -dChannelH + dT, m_outputColl, "", "POLYPT4"
    AddPoint -dD / 2#, -dChannelW / 2# - dT, dT, m_outputColl, "", "POLYPT5"
    AddPoint -dD / 2#, dChannelW / 2# + dT, dT, m_outputColl, "", "POLYPT6"
    AddPoint -dD / 2#, dChannelW / 2# + dT, -dChannelH + dT, m_outputColl, "", "POLYPT7"
    AddPoint -dD / 2#, dW / 2#, -dChannelH + dT, m_outputColl, "", "POLYPT8"
    AddPoint -dD / 2#, dW / 2#, -dChannelH, m_outputColl, "", "POLYPT9"
    AddPoint -dD / 2#, dChannelW / 2#, -dChannelH, m_outputColl, "", "POLYPT10"
    AddPoint -dD / 2#, dChannelW / 2#, 0, m_outputColl, "", "POLYPT11"
    AddExtrusion dD, 0, 0, 1, m_outputColl, "Body"

    Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    ' <Determine BOM description here>
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim sItemNo As String
    Dim sConnectorBOM As String
    Dim dWeight As Double
    Dim sInterfaceName As String

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    If InStr(UCase(sPartNumber), "-R") <> 0 Then
        sConnectorBOM = " + 3, 304012, Channel connector MQN-R"
    Else
        If InStr(UCase(sPartNumber), "-F") <> 0 Then
            sConnectorBOM = " + 3, 387779, Channel Connector MQN-HDG plus"
        Else
            sConnectorBOM = " + 3, 369623, Channel connector MQN"
        End If
    End If

    bstrBOMDesc = Hilti_BuildBom(pSupportComp) & sConnectorBOM

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub



