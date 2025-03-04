VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQT_M12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQT_M12.cls
'   ProgID:         HS_Hilti_MQParts.MQT_M12
'   Author:         JRK
'   Creation Date:  Wednesday, November 21, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP         Addition to S3D
'   13-07-2010        YLP         D-12966 Application error when selecting blank channel size for “Hilti MQT-M12” part.
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQT_M12"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQT_M12"

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
    Const iNumInputs As Integer = 10
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 4
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "ChannelSize"
    sDataType(2) = "Long"

    sInputName(3) = "AppCase"
    sDataType(3) = "Long"

    sInputName(4) = "ConnSide"
    sDataType(4) = "Long"

    sInputName(5) = "B"
    sDataType(5) = "Double"

    sInputName(6) = "D"
    sDataType(6) = "Double"

    sInputName(7) = "T"
    sDataType(7) = "Double"

    sInputName(8) = "ItemNo"
    sDataType(8) = "String"

    sInputName(9) = "Group"
    sDataType(9) = "String"

    sInputName(10) = "FlangeThick"
    sDataType(10) = "Double"

    sOutputName(1) = "Channel"
    sOutputName(2) = "Box1"
    sOutputName(3) = "Box2"
    sOutputName(4) = "Left"

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
    Dim lAppCase As String
    Dim lConnSide As Long
    Dim dB As Double
    Dim dD As Double
    Dim dT As Double
    Dim dH As Double
    Dim dChannelH As Double
    Dim dBoltL As Double
    Dim sChannelSize As String
    Dim sChannelRecord As String

    lChannelSize = arrayOfInputs(2)
    lAppCase = arrayOfInputs(3)
    lConnSide = arrayOfInputs(4)
    dB = arrayOfInputs(5)
    dD = arrayOfInputs(6)
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

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    
    If lChannelSize = -1 Then
        lChannelSize = 1
        sChannelSize = Trim(oCLMD.ShortStringValue(strTable, lChannelSize))
        PF_EventHandler "Channel Size should not be Empty. Re-setting the value to " & sChannelSize, Err, MODULE, METHOD, True
    End If
    
    sChannelSize = Trim(oCLMD.ShortStringValue(strTable, lChannelSize))
    
    sChannelRecord = ReadParametricData("JUAMQT_Aux", "where ChannelSize = '" + sChannelSize + "'")
    dChannelH = GetNParam(sChannelRecord, "ChannelH")

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dFlangeT = 0.01
    Else
        dFlangeT = arrayOfInputs(10)
    End If

    If dFlangeT > 0.036 Then
        PF_EventHandler "Flange Thickness exceeds maximum allowable.", Err, MODULE, METHOD, True
    End If

    Dim sYN As String
    If lAppCase = -1 Then
        lAppCase = 1
        PF_EventHandler "Application Case should not be Null. Re-setting to Case 1", Err, MODULE, METHOD, True
    End If
    
    sYN = Trim(GetSParam(sChannelRecord, Trim(Str(lAppCase)) & "_YNMQTM12"))
    If UCase(sYN) = "N" Then
        PF_EventHandler "Selected channel size is not applicable with selected beam clamp size.", Err, MODULE, METHOD, True
    End If

    If lAppCase = 1 Then
        dBoltL = 0.06 + dFlangeT + dChannelH
    ElseIf lAppCase = 2 Then
        dBoltL = 0.041 + (0.6 * dFlangeT)
    End If

    AddPort "Channel", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Channel", arrayOfInputs(1), "Channel"

    If lAppCase = 1 Then
        AddBox dT, dB, dH, Loc(-dT / 2#, -dB / 2# - dD / 2#, 0), m_outputColl, "Box1", "Box1"
        AddBox dT, dB, dH, Loc(-dT / 2#, -dB / 2# - dD / 2#, -dChannelH - dH), m_outputColl, "Box2", "Box2"
        AddCylinder dBoltL, dD / 2#, Loc(0, -dD / 2#, -dChannelH / 2# - dBoltL / 2#), m_outputColl, "Left", "Left"
    Else
        AddBox dT, dB, dH, Loc(-dT / 2#, -dB / 2# - dD / 2#, 0), m_outputColl, "Box1", "Box1"
        AddCylinder dBoltL, dD / 2#, Loc(0, -dD / 2#, 0.041 / 2# + dD / 2# - dBoltL), m_outputColl, "Left", "Left"
    End If
    
    Set oSupportComp = Nothing
    Set oICH = Nothing

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
    oPartOcc.GetPart oPart      ' The associated catalog part

    Dim sItemNo As String
    Dim dWeight As Double
    Dim lAppCase As Long
    Dim dBoltL As Double
    Dim dFlangeT As Double
    Dim dChannelH As Double
    Dim lChannelSize As Long
    Dim sChannelSize As String
    Dim sChannelRecord As String
    Dim sChannelRecord2 As String
    Dim sChannelConnItemNo As String
    Dim sChannelConnSize As String

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)
    lAppCase = GetAttributeFromObject(pSupportComp, "AppCase")

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    Set oSupportComp = pSupportComp
    oSupportComp.GetOccAssembly oICH

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dFlangeT = 0.01
    Else
        dFlangeT = GetAttributeFromObject(pSupportComp, "FlangeThick")
    End If
        
    lChannelSize = GetAttributeFromObject(pSupportComp, "ChannelSize")
      
    Dim vInterfaceID As Variant
    Dim oAttribMetaData As IJDAttributeMetaData
    Dim oAttribInfo As IJDAttributeInfo
    Dim oCLMD As METADATALib.IJDCodeListMetaData
    Dim strTable As String
    Set oCLMD = pSupportComp

    Set oAttribMetaData = pSupportComp
    vInterfaceID = oAttribMetaData.iID("IJOAMQChSizePk")
    Set oAttribInfo = oAttribMetaData.AttributeInfo(vInterfaceID, "ChannelSize")
    strTable = oAttribInfo.CodeListTableName
    
    If lChannelSize = -1 Then
        lChannelSize = 1
        oICH.SetSymbolInputByName pSupportComp, "ChannelSize", lChannelSize
        sChannelSize = Trim(oCLMD.ShortStringValue(strTable, lChannelSize))
    End If
    sChannelSize = Trim(oCLMD.ShortStringValue(strTable, lChannelSize))

    sChannelRecord = ReadParametricData("JUAMQT_Aux", "where ChannelSize = '" + sChannelSize + "'")
    dChannelH = GetNParam(sChannelRecord, "ChannelH")
    sChannelRecord2 = ReadParametricData("JUAMQ_ConnTypes", "where ChannelConn = 'MQM-M12'")
    sChannelConnItemNo = Trim(GetSParam(sChannelRecord2, "ChannelConnItemGalv"))
    sChannelConnSize = Trim(GetSParam(sChannelRecord2, "ChannelConnGalv"))

    If lAppCase = -1 Then
        lAppCase = 1
        oICH.SetSymbolInputByName pSupportComp, "AppCase", lAppCase
    End If
    
    If lAppCase = 1 Then
        dBoltL = 0.06 + dFlangeT + dChannelH
    ElseIf lAppCase = 2 Then
        dBoltL = 0.041 + (0.6 * dFlangeT)
    End If

    If sChannelConnSize = "-" Then
        PF_EventHandler "Selected connection type is not available for selected material.", Err, MODULE, METHOD, True
    End If

    sItemNo = GetAttributeFromObject(oPart, "ItemNo")

    If lAppCase = 1 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    ElseIf lAppCase = 2 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1, " & sChannelConnItemNo & " " & sChannelConnSize & " + 1, M12x" + Str(dBoltL * 1000)
    End If
    
    Set oICH = Nothing
    Set oSupportComp = Nothing
    Set oCLMD = Nothing
    Set oAttribMetaData = Nothing
    Set oAttribInfo = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
