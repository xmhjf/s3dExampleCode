VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQP.cls
'   ProgID:         HS_Hilti_MQParts.MQP
'   Author:         JRK
'   Creation Date:  Wednesday, November 14, 2007

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
Private Const MODULE = "MQP"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQP"

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
    AddInput Inputs, "ConnType", "Long"
    AddInput Inputs, "ChannelW", "Double"
    AddInput Inputs, "T", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "H", "Double"
    AddInput Inputs, "ChannelD", "Double"
    AddInput Inputs, "NumMQN", "Long"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "BaseStyle", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Structure1"
    AddOutput Outputs, "Structure2"
    AddOutput Outputs, "Body"
    AddOutput Outputs, "Box"
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
    Dim dT As Double
    Dim dW As Double
    Dim dChannelD As Double
    Dim dD As Double
    Dim dH As Double
    Dim lNumMQN As Long
    Dim sGroup As String
    Dim dT2 As Double
    Dim dOtherW As Double
    Dim dOtherH As Double
    Dim sBaseStyle As String

    dChannelW = arrayOfInputs(3)
    dT = arrayOfInputs(4)
    dW = arrayOfInputs(5)
    dD = arrayOfInputs(6)
    dH = arrayOfInputs(7)
    dChannelD = arrayOfInputs(8)
    lNumMQN = arrayOfInputs(9)
    sGroup = Trim(arrayOfInputs(11))
    sBaseStyle = Trim(arrayOfInputs(12))

    If sBaseStyle = "2" Then
        dT2 = 0.004
        AddPort "Structure1", 0, 0, 0, _
                -1, 0, 0, _
                0, 0, -1, m_outputColl, "Structure1", arrayOfInputs(1), "Structure1"
        AddPort "Structure2", 0, 0, dT, _
                -1, 0, 0, _
                0, 0, -1, m_outputColl, "Structure2", arrayOfInputs(1), "Structure2"
        AddBox dW, dD, dT, Loc(-dW / 2#, -dD / 2#, 0), m_outputColl, "Body", "Body"
        AddBox dChannelW + dT2 * 2#, dChannelD + dT2 * 2#, dH, Loc(-dChannelW / 2# - dT2, -dChannelD / 2# - dT2, dT), m_outputColl, "Box", "Box"
    Else
        AddPort "Structure1", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "Structure1", arrayOfInputs(1), "Structure1"
        dW = 0.0413
        dH = 0.061
        dT = 0.004
        dOtherW = 0.167
        dOtherH = 0.112
        AddPolygon 6, "EXTRUDED", 0, "POLY"
        AddPoint 0, -dW / 2#, 0, m_outputColl, "", "POLYPT0"
        AddPoint 0, -dW / 2#, dH, m_outputColl, "", "POLYPT1"
        AddPoint dT, -dW / 2#, dH, m_outputColl, "", "POLYPT2"
        AddPoint dT, -dW / 2#, dT, m_outputColl, "", "POLYPT3"
        AddPoint dH, -dW / 2#, dT, m_outputColl, "", "POLYPT4"
        AddPoint dH, -dW / 2#, 0, m_outputColl, "", "POLYPT5"
        AddExtrusion 0, dW, 0, 1, m_outputColl, "Body"
    End If

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

    Dim lConnType As Long
    Dim sConnType As String
    Dim sItemNo As String
    Dim lMaterial As Long
    Dim sRecord1 As String
    Dim sQTY As String
    Dim sConnTypeID As String
    Dim sAnchorSize As String
    Dim sBOMAdd As String
    Dim sChannelConn As String
    Dim sSteelConnSize As String
    Dim sConcreteConn As String
    Dim sSteelSize As String
    Dim sSteelConn As String
    Dim sAnchorItemNo As String
    Dim sChannelConnSize As String
    Dim sChannelConnItemNo As String
    Dim sSteelConnItemNo As String
    Dim iStart As Integer
    Dim sConnectorBOM As String
    Dim dWeight As Double
    Dim sGroup As String
    Dim sNumMQN As String

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    lConnType = GetAttributeFromObject(pSupportComp, "ConnType")
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sGroup = UCase(GetAttributeFromObject(oPart, "Group"))
    sNumMQN = Trim(GetAttributeFromObject(oPart, "NumMQN"))

    Dim vInterfaceID As Variant
    Dim oAttribMetaData As IJDAttributeMetaData
    Dim oAttribInfo As IJDAttributeInfo
    Dim oCLMD As METADATALib.IJDCodeListMetaData
    Dim strTable As String
    Set oCLMD = pSupportComp
    Set oAttribMetaData = pSupportComp

    If sGroup = "CHANNEL BASE" Then
        sQTY = "2"
    Else
        sQTY = "1"
    End If

    If InStr(UCase(sPartNumber), "-R") <> 0 Then
        sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeSS", "ConnType", lConnType)
        vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeSS")
        sConnectorBOM = ", 304012, Channel connector MQN-R"
    Else
        If InStr(UCase(sPartNumber), "-F") <> 0 Then
            sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeHDG2", "ConnType", lConnType)
            vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeHDG2")
            sConnectorBOM = ", 387779, Channel Connector MQN-HDG plus"
        Else
            sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeGalv", "ConnType", lConnType)
            vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeGalv")
            sConnectorBOM = ", 369623, Channel connector MQN"
        End If
    End If

    Set oAttribInfo = oAttribMetaData.AttributeInfo(vInterfaceID, "ConnType")
    strTable = oAttribInfo.CodeListTableName
    sConnTypeID = oCLMD.ShortStringValue(strTable, lConnType)
    sChannelConn = ""
    sConcreteConn = ""
    sSteelConn = ""
    sAnchorSize = ""
    sAnchorItemNo = ""
    sChannelConnSize = ""
    sChannelConnItemNo = ""
    sSteelConnSize = ""
    sSteelConnItemNo = ""

    If InStr(UCase(sConnType), "CONCRETE") <> 0 Then
        iStart = Len(sConnType) - 9
        'new---------------------------- Concrete /Anchor
        sConcreteConn = Right(sConnType, iStart)
        'old----------------------------
        sAnchorSize = sConcreteConn
        sAnchorItemNo = sConnTypeID
    End If

    If InStr(UCase(sConnType), "CHANNEL") <> 0 Then
        iStart = Len(sConnType) - 8
        'new---------------------------- Channel
        sChannelConn = Right(sConnType, iStart)
        'old----------------------------
        sChannelConnSize = sChannelConn
        sChannelConnItemNo = sConnTypeID
    End If

    If InStr(UCase(sConnType), "STEEL") <> 0 Then
        iStart = Len(sConnType) - 6
        'new---------------------------- Steel
        sSteelConn = Right(sConnType, iStart)
        'old----------------------------
        sSteelConnSize = sSteelConn
        sSteelConnItemNo = sConnTypeID
    End If

    If sConnType = "None" Then
        sBOMAdd = ""
    Else
        If InStr(UCase(sConnType), "CHANNEL") <> 0 Then

            If InStr(UCase(sChannelConnSize), "MQM M12") <> 0 Then
                sBOMAdd = sQTY + ", 216458, Hexagon Screw M12X25"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 F") <> 0 Then
                sBOMAdd = sQTY + ", 304766, Hexagon Nut M12-F"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 R") <> 0 Then
                sBOMAdd = sQTY + ", 87634, Hexagon Screw M12X25 A4-70"
            End If

        Else

            If InStr(UCase(sConnType), "CONCRETE") <> 0 Then

                If InStr(UCase(sConcreteConn), "M12X95/25") <> 0 Then
                    sBOMAdd = sQTY + ", 311369, HVU-TZ M12"
                Else
                    If InStr(UCase(sConcreteConn), "M12X75/15") <> 0 Then
                        sBOMAdd = sQTY + ", 334539, HIT-HY 150/330/2"
                    Else
                        If InStr(UCase(sConcreteConn), "M12X110/28") <> 0 Then
                            sBOMAdd = sQTY + ", 310625, HVU-M12x110"
                        Else
                            sBOMAdd = ""
                        End If
                    End If
                End If

            Else
                sBOMAdd = sQTY + ", 58042, Washer A 10.5/20 A4 " + sQTY + ", 52461, Hexagon Nut M10 A4-70"
            End If
        End If
    End If

    If UCase(sGroup) = "CHANNEL BASE" Then
        If InStr(UCase(sConnType), "NONE") <> 0 Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1" & sConnectorBOM
        Else
            If InStr(UCase(sConnType), "CHANNEL") <> 0 Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sChannelConn & " + " & sBOMAdd & " + " & sNumMQN & sConnectorBOM
            Else
                If InStr(UCase(sConnType), "CONCRETE") <> 0 Then
                    bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sAnchorSize & " + " & sBOMAdd & " + " & sNumMQN & sConnectorBOM
                Else
                    bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sSteelConnSize & " + " & sBOMAdd & " + " & sNumMQN & sConnectorBOM
                End If
            End If
        End If
    Else
        If InStr(UCase(sConnType), "NONE") <> 0 Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1" & sConnectorBOM
        Else
            If InStr(UCase(sConnType), "CHANNEL") <> 0 Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sChannelConn & " + " & sBOMAdd & " + " & sQTY & sConnectorBOM
            Else
                If InStr(UCase(sConnType), "CONCRETE") <> 0 Then
                    bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sAnchorSize & " + " & sBOMAdd & " + " & sQTY & sConnectorBOM
                Else
                    bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sSteelConnSize & " + " & sBOMAdd & " + " & sQTY & sConnectorBOM
                End If
            End If
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub



