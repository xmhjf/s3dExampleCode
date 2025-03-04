VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQT_41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQT_41.cls
'   ProgID:         HS_Hilti_MQParts.MQT_41
'   Author:         JRK
'   Creation Date:  Wednesday, November 21, 2007

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
Private Const MODULE = "MQT_41"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQT_41"

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
    Const iNumInputs As Integer = 9
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 8
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "ChannelSize"
    sDataType(2) = "Long"

    sInputName(3) = "AppCase"
    sDataType(3) = "String"

    sInputName(4) = "B"
    sDataType(4) = "Double"

    sInputName(5) = "D"
    sDataType(5) = "Double"

    sInputName(6) = "H"
    sDataType(6) = "Double"

    sInputName(7) = "T"
    sDataType(7) = "Double"

    sInputName(8) = "ItemNo"
    sDataType(8) = "String"

    sInputName(9) = "Group"
    sDataType(9) = "String"

    sOutputName(1) = "Channel"
    sOutputName(2) = "Bottom"
    sOutputName(3) = "RightBottom"
    sOutputName(4) = "LeftBottom"
    sOutputName(5) = "RightTop"
    sOutputName(6) = "LeftTop"
    sOutputName(7) = "Top"
    sOutputName(8) = "Rod"

    SetupInputsAndOutputs pSymbolDefinition, iNumInputs, iNumOutputs, sInputName(), sDataType(), sOutputName(), m_progID


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

    Dim lChannelSize As Long
    Dim dFlangeT As Double
    Dim sAppCase As String
    Dim dB As Double
    Dim dD As Double
    Dim dH As Double
    Dim dT As Double
    Dim dChannelH As Double
    Dim dBoltL As Double
    Dim sChannelSize As String
    Dim sChannelRecord As String

    lChannelSize = arrayOfInputs(2)
    sAppCase = arrayOfInputs(3)
    dB = arrayOfInputs(4)
    dD = arrayOfInputs(5)
    dH = arrayOfInputs(6)
    dT = arrayOfInputs(7)
    dH = 0.02

    Dim vInterfaceID As Variant
    Dim oAttribMetaData As IJDAttributeMetaData
    Dim oAttribInfo As IJDAttributeInfo
    Dim oCLMD As METADATALib.IJDCodeListMetaData
    Dim strTable As String
    Set oCLMD = sblOcc

    Set oAttribMetaData = sblOcc
    vInterfaceID = oAttribMetaData.iID("IJOAMQChSizePk")
    Set oAttribInfo = oAttribMetaData.AttributeInfo(vInterfaceID, "ChannelSize")
    strTable = oAttribInfo.CodeListTableName
    sChannelSize = Trim(oCLMD.ShortStringValue(strTable, lChannelSize))

    sChannelRecord = ReadParametricData("JUAMQT_Aux", "where ChannelSize = '" + sChannelSize + "'")
    dChannelH = GetNParam(sChannelRecord, "ChannelH")

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    Set oSupportComp = Nothing

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dFlangeT = 0.01
    Else
        Hilti_InitializeMyHH oICH
       
        Set oICH = Nothing
    End If

    Dim sYN As String
    sYN = Trim(GetSParam(sChannelRecord, "One_YNMQT41"))
    If UCase(sYN) = "N" Then
        PF_EventHandler "Selected channel size is not applicable with selected beam clamp size.", Err, MODULE, METHOD, True
    End If

    AddPort "Channel", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Channel", arrayOfInputs(1), "Channel"

    AddBox 0.052, 0.0685, 0.005, Loc(-0.026, -0.03425, -0.005), m_outputColl, "Bottom", "Bottom"
    AddBox 0.052, 0.005, 0.041, Loc(-0.026, -0.03425, 0), m_outputColl, "RightBottom", "RightBottom"
    AddBox 0.052, 0.005, 0.041, Loc(-0.026, 0.03425 - 0.005, 0), m_outputColl, "LeftBottom", "LeftBottom"
    AddBox 0.025, 0.005, 0.0486, Loc(-0.026, -0.03425, 0.046 - 0.005), m_outputColl, "RightTop", "RightTop"
    AddBox 0.025, 0.005, 0.0486, Loc(-0.026, 0.03425 - 0.005, 0.046 - 0.005), m_outputColl, "LeftTop", "LeftTop"
    AddBox 0.048, 0.0685, 0.01, Loc(-0.026, -0.03425, 0.0946 - 0.005), m_outputColl, "Top", "Top"
    AddCylinder 0.05, 0.006, Loc(0.01, 0, 0.066 - 0.005), m_outputColl, "Rod", "Rod"
    
    Set oCLMD = Nothing
    Set oAttribMetaData = Nothing
    Set oAttribInfo = Nothing
    
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
    Dim dWeight As Double

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    dWeight = Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASHilti_MQT_41_GALV", "PartNumber", "JDPart", "'" & sPartNumber & "'")

    bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + 1, 369623, Channel Connector MQN"
    Set oPartOcc = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
