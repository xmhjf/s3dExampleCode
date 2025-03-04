VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQT.cls
'   ProgID:         HS_Hilti_MQParts.MQT
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
Private Const MODULE = "MQT"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQT"

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
    Const iNumInputs As Integer = 12
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 10
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "ChannelSize"
    sDataType(2) = "Long"

    sInputName(3) = "NumClamps"
    sDataType(3) = "Long"

    sInputName(4) = "Gap"
    sDataType(4) = "Double"

    sInputName(5) = "AppCase"
    sDataType(5) = "String"

    sInputName(6) = "B"
    sDataType(6) = "Double"

    sInputName(7) = "D"
    sDataType(7) = "Double"

    sInputName(8) = "H"
    sDataType(8) = "Double"

    sInputName(9) = "T"
    sDataType(9) = "Double"

    sInputName(10) = "ItemNo"
    sDataType(10) = "String"

    sInputName(11) = "Group"
    sDataType(11) = "String"

    sInputName(12) = "FlangeThick"
    sDataType(12) = "Double"

    sOutputName(1) = "Channel1"
    sOutputName(2) = "Channel2"
    sOutputName(3) = "Plate"
    sOutputName(4) = "Rod1"
    sOutputName(5) = "Rod2"
    sOutputName(6) = "Rod3"
    sOutputName(7) = "Plate2"
    sOutputName(8) = "Rod4"
    sOutputName(9) = "Rod5"
    sOutputName(10) = "Rod6"

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
    Dim lNumClamps As Long
    Dim dGap As Double
    Dim sAppCase As String
    Dim dB As Double
    Dim dD As Double
    Dim dH As Double
    Dim dT As Double
    Dim dChannelH As Double
    Dim dBoltL As Double
    Dim sChannelSize As String
    Dim sChannelRecord As String
    Dim dFlangeT As Double

    lChannelSize = arrayOfInputs(2)
    lNumClamps = arrayOfInputs(3)
    dGap = arrayOfInputs(4)
    sAppCase = arrayOfInputs(5)
    dB = arrayOfInputs(6)
    dD = arrayOfInputs(7)
    dH = arrayOfInputs(8)
    dT = arrayOfInputs(9)
    dBoltL = 0.041 + (0.6 * dFlangeT)
    If lNumClamps = 1 Then
        dGap = 0
    End If

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    Set oSupportComp = Nothing

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dFlangeT = 0.01
    Else
        Hilti_InitializeMyHH oICH
        ' Changed to use hshlpr instead of hs_hgrhlpr

        'Steel Dimensions
        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dFlangeT = SupportStruct.dFlangeThickness
        Else
            dFlangeT = arrayOfInputs(12)
        End If
        
        Set oStruct = Nothing
        
        Set oICH = Nothing
    End If

    If dFlangeT > 0.023 Then
        PF_EventHandler "Flange Thickness exceeds maximum allowable.", Err, MODULE, METHOD, True
    End If

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

    sChannelRecord = ReadParametricData("JUAMQT_Aux", "where ChannelSize =  '" + sChannelSize + "'")
    dChannelH = GetNParam(sChannelRecord, "ChannelH")

    AddPort "Channel1", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Channel1", arrayOfInputs(1), "Channel1"
    AddPort "Channel2", 0, 0, -dChannelH - dFlangeT, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Channel2", arrayOfInputs(1), "Channel2"

    AddBox dH, dB, 0.005, Loc(-dH / 2 - dGap / 2, -dB / 2, 0), m_outputColl, "Plate", "Plate"
    AddCylinder dH, dD / 2, Loc(-dGap / 2, 0.041 / 2 + dD / 2, -dChannelH - dD - dFlangeT), m_outputColl, "Rod1", "Rod1"
    AddCylinder dH, dD / 2, Loc(-dGap / 2, -0.041 / 2 - dD / 2, -dChannelH - dD - dFlangeT), m_outputColl, "Rod2", "Rod2"
    AddCylinder 0.041, dD / 2, "ROTX(270)*" + Loc(-dGap / 2, -0.041 / 2, -dChannelH - dD / 2 - dFlangeT), m_outputColl, "Rod3", "Rod3"

    If lNumClamps = 2 Then
        AddBox dH, dB, 0.005, Loc(-dH / 2 + dGap / 2, -dB / 2, 0), m_outputColl, "Plate2", "Plate2"
        AddCylinder dH, dD / 2, Loc(dGap / 2, 0.041 / 2 + dD / 2, -dChannelH - dD - dFlangeT), m_outputColl, "Rod4", "Rod4"
        AddCylinder dH, dD / 2, Loc(dGap / 2, -0.041 / 2 - dD / 2, -dChannelH - dD - dFlangeT), m_outputColl, "Rod5", "Rod5"
        AddCylinder 0.041, dD / 2, "ROTX(270)*" + Loc(dGap / 2, -0.041 / 2, -dChannelH - dD / 2 - dFlangeT), m_outputColl, "Rod6", "Rod6"
    End If
    
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
    Dim sConnectorBOM As String
    Dim dWeight As Double
    Dim lNumClamps As Long

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    lNumClamps = GetAttributeFromObject(pSupportComp, "NumClamps")


    If lNumClamps = 1 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    Else
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1, " & Replace(Hilti_BuildBom(pSupportComp), "Hilti, ", "")
    End If
    
    Set oPartOcc = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
