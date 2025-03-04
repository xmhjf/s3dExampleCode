VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQK_S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQK_S.cls
'   ProgID:         HS_Hilti_MQParts.MQK_S
'   Author:         JRK
'   Creation Date:  Saturday, October 13, 2007

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
Private Const MODULE = "MQK_S"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQK_S"

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
    AddInput Inputs, "L", "Double"
    AddInput Inputs, "H", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "A", "Double"
    AddInput Inputs, "LLength", "Double"
    AddInput Inputs, "PlateH", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Structure1"
    AddOutput Outputs, "Structure2"
    AddOutput Outputs, "Structure3"
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

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    Set oSupportComp = Nothing

    Dim dMatTh As Double
    Dim lConnType As Long
    Dim dL As Double
    Dim dH As Double
    Dim dW As Double
    Dim dA As Double
    Dim dLLength As Double
    Dim dPlateH As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim dDiff As Double
    Dim dHeight As Double

    dMatTh = 0.003
    lConnType = arrayOfInputs(2)
    dL = arrayOfInputs(3)
    dH = arrayOfInputs(4)
    dW = arrayOfInputs(5)
    dA = arrayOfInputs(6)
    dLLength = arrayOfInputs(7)
    dPlateH = arrayOfInputs(8)
    sItemNo = arrayOfInputs(9)
    sGroup = arrayOfInputs(10)
    dDiff = dMatTh / Sqr(2) - (0.003 - (dMatTh / Sqr(2)))
    dHeight = Sqr(((dA) * (dA)) - ((dL - dLLength) * (dL - dLLength)))

    AddPort "Structure1", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure1", arrayOfInputs(1), "Structure1"

    AddPort "Structure2", 0, -dW / 2, -dW / 2, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure2", arrayOfInputs(1), "Structure2"
    AddPort "Structure3", -dL + dLLength / 2, 0, dHeight + dPlateH / 2, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure3", arrayOfInputs(1), "Structure3"

    AddPolygon 8, "EXTRUDED", 0, "BODY"
    AddPoint dLLength / 2, -dW / 2, 0, m_outputColl, "", "BODYPT0"
    AddPoint -dLLength / 2, -dW / 2, 0, m_outputColl, "", "BODYPT1"
    AddPoint -dL + dLLength / 2, -dW / 2, dHeight, m_outputColl, "", "BODYPT2"
    AddPoint -dL + dLLength / 2, -dW / 2, dHeight + dPlateH, m_outputColl, "", "BODYPT3"
    AddPoint -dL + dLLength / 2 + dMatTh, -dW / 2, dHeight + dPlateH, m_outputColl, "", "BODYPT4"
    AddPoint -dL + dLLength / 2 + dMatTh, -dW / 2, dHeight + dDiff, m_outputColl, "", "BODYPT5"
    AddPoint -dLLength / 2 + dDiff, -dW / 2, dMatTh, m_outputColl, "", "BODYPT6"
    AddPoint dLLength / 2, -dW / 2, dMatTh, m_outputColl, "", "BODYPT7"
    AddExtrusion 0, dW, 0, 1, m_outputColl, "Body"

    AddBox dLLength, dMatTh, dW + dMatTh, Loc(-dLLength / 2, -dW / 2# - dMatTh, -dW), m_outputColl, "Box", "Box"

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

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    lConnType = GetAttributeFromObject(pSupportComp, "ConnType")
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")

    Dim vInterfaceID As Variant
    Dim oAttribMetaData As IJDAttributeMetaData
    Dim oAttribInfo As IJDAttributeInfo
    Dim oCLMD As METADATALib.IJDCodeListMetaData
    Dim strTable As String
    Set oCLMD = pSupportComp
    Set oAttribMetaData = pSupportComp

    If InStr(UCase(sPartNumber), "-R") <> 0 Then
        sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeSS", "ConnType", lConnType)
        vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeSS")
        sConnectorBOM = " + 1, 304012, Channel connector MQN-R"
    Else
        If InStr(UCase(sPartNumber), "-F") <> 0 Then
            sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeHDG", "ConnType", lConnType)
            vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeHDG")
            sConnectorBOM = " + 1, 387779, Channel Connector MQN-HDG plus"
        Else
            sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeGalv", "ConnType", lConnType)
            vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeGalv")
            sConnectorBOM = " + 1, 369623, Channel connector MQN"
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
        sQTY = ""
        sBOMAdd = ""
    Else
        sQTY = " + 1, "

        If InStr(UCase(sConnType), "CHANNEL") <> 0 Then

            If InStr(UCase(sChannelConnSize), "MQM M12") <> 0 Then
                sBOMAdd = sQTY + "216458, Hexagon Screw M12X25"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 F") <> 0 Then
                sBOMAdd = sQTY + "304766, Hexagon Nut M12-F"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 R") <> 0 Then
                sBOMAdd = sQTY + "87634, Hexagon Screw M12X25 A4-70"
            End If

        Else

            If InStr(UCase(sConnType), "CONCRETE") <> 0 Then

                If InStr(UCase(sConcreteConn), "M12X95/25") <> 0 Then
                    sBOMAdd = sQTY + "311369, HVU-TZ M12"
                Else
                    If InStr(UCase(sConcreteConn), "M12X75/15") <> 0 Then
                        sBOMAdd = sQTY + "334539, HIT-HY 150/330/2"
                    Else
                        If InStr(UCase(sConcreteConn), "M12X110/28") <> 0 Then
                            sBOMAdd = sQTY + "310625, HVU-M12x110"
                        Else
                            sBOMAdd = ""
                        End If
                    End If
                End If

            Else
                sBOMAdd = sQTY + "58042, Washer A 10.5/20 A4 " + sQTY + "52461, Hexagon Nut M10 A4-70"
            End If
        End If
    End If

    If InStr(UCase(sConnType), "CONCRETE") <> 0 Then
        bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBOM & sQTY & sConnTypeID & " " & sAnchorSize & sBOMAdd
    Else
        If InStr(UCase(sConnType), "CHANNEL") <> 0 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBOM & sQTY & sConnTypeID & " " & sChannelConn & sBOMAdd
        Else
            If InStr(UCase(sConnType), "STEEL") <> 0 Then
                bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBOM & sQTY & sConnTypeID & " " & sSteelConnSize & sBOMAdd
            End If
        End If
    End If

    If UCase(Trim(sConnType)) = "NONE" Then
        bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBOM
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
