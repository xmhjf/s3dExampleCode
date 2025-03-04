VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQP_U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQP_U.cls
'   ProgID:         HS_Hilti_MQParts.MQP_U
'   Author:         JRK
'   Creation Date:  Friday, November 16, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP         Addition to S3D
'   02/09/11          Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14/09/11          VSP         TR-CP-193697  Incorrect BOM Description and Warnings were observed in part placement
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQP_U"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQP_U"

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
    AddInput Inputs, "BaseL", "Double"
    AddInput Inputs, "BaseH", "Double"
    AddInput Inputs, "BaseW", "Double"
    AddInput Inputs, "TopL", "Double"
    AddInput Inputs, "TopW", "Double"
    AddInput Inputs, "TopH", "Double"
    AddInput Inputs, "T", "Double"
    AddInput Inputs, "Angle", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Structure"
    AddOutput Outputs, "Channel"
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

    Dim dAngle As Double
    Dim dBaseL As Double
    Dim dBaseH As Double
    Dim dBaseW As Double
    Dim dTopL As Double
    Dim dTopW As Double
    Dim dTopH As Double
    Dim dCalc1 As Double
    Dim dCalc2 As Double
    Dim dT As Double

    dBaseL = arrayOfInputs(3)
    dBaseH = arrayOfInputs(4)
    dBaseW = arrayOfInputs(5)
    dTopL = arrayOfInputs(6)
    dTopW = arrayOfInputs(7)
    dTopH = arrayOfInputs(8)
    dT = arrayOfInputs(9)
    dAngle = arrayOfInputs(10)
    dAngle = Deg(dAngle)

    'Add the connection points
    If dAngle = 30 Or dAngle = 60 Then
        dCalc1 = dBaseH / 2 + (Sin(dAngle) * dTopH)
        dCalc2 = (CosDeg(dAngle) * dTopH)
    Else
        If dAngle = 90 Then
            dCalc1 = dBaseH / 2 + dTopH
            dCalc2 = 0
        Else
            dCalc1 = dBaseH / 2
            dCalc2 = dTopH
        End If
    End If

    If dAngle > 90.001 Or dAngle < -0.001 Then
        PF_EventHandler "Pivot angle must be between 0 and 90 degrees.", Err, MODULE, METHOD, True
    End If

    AddPort "Structure", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"
    AddPort "Channel", 0, 0, dBaseH / 2, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Channel", arrayOfInputs(1), "Channel"

    AddPolygon 8, "EXTRUDED", 0, "POLY"
    AddPoint -dBaseL / 2#, -dBaseW / 2#, 0, m_outputColl, "", "POLYPT0"
    AddPoint -dBaseL / 2#, dBaseW / 2#, 0, m_outputColl, "", "POLYPT1"
    AddPoint -dBaseL / 2#, dBaseW / 2#, dBaseH, m_outputColl, "", "POLYPT2"
    AddPoint -dBaseL / 2#, dBaseW / 2# - dT, dBaseH, m_outputColl, "", "POLYPT3"
    AddPoint -dBaseL / 2#, dBaseW / 2# - dT, dT, m_outputColl, "", "POLYPT4"
    AddPoint -dBaseL / 2#, -dBaseW / 2# + dT, dT, m_outputColl, "", "POLYPT5"
    AddPoint -dBaseL / 2#, -dBaseW / 2# + dT, dBaseH, m_outputColl, "", "POLYPT6"
    AddPoint -dBaseL / 2#, -dBaseW / 2#, dBaseH, m_outputColl, "", "POLYPT7"
    AddExtrusion dBaseL, 0, 0, 1, m_outputColl, "Body"

    Dim dPHI As Double
    Dim dDZ1 As Double
    Dim dDY1 As Double
    Dim dR As Double

    dPHI = 90 - dAngle
    dR = dTopL / 2

    dDY1 = dR * CosDeg(dPHI)
    dDZ1 = dR * SinDeg(dPHI)

    Dim dXH As Double
    Dim dYH As Double
    Dim dZH As Double
    Dim dBaseZ As Double

    dXH = dTopW / 2
    dYH = dTopL / 2
    dBaseZ = dBaseH / 2

    AddBox dTopW, dTopL, dTopH, "ROTX(" + Str(dPHI) + ")*" + Loc(-dXH, -dYH - (dDY1 - dYH), dBaseZ - dDZ1) + "*ROTZ(90)", m_outputColl, "Box", "Box"

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

    If InStr(UCase(sPartNumber), "MQP-U M12/12-R") <> 0 Then
        sConnectorBOM = " + 2, 52462, Hexagon nut M12 A4-70 + 2, 58041, Flat washer A 13/24 A4"
    ElseIf InStr(UCase(sPartNumber), "MQP-U M16/16-R") <> 0 Then
        sConnectorBOM = " + 2, 52463, Hexagon nut M16 A4-70 + 2, 387989, Flat washer A 17/30 A4"
    ElseIf InStr(UCase(sPartNumber), "MQP-U M12-F") <> 0 Then
        sConnectorBOM = " + 2, 304766, Hexagon nut M12-F + 2, 304771, Flat washer A 13/24-F"
    ElseIf InStr(UCase(sPartNumber), "MQP-U M16-F") <> 0 Then
        sConnectorBOM = " + 2, 304767, Hexagon nut M16-F + 2, 304772, Flat washer A 17/30-F"
    ElseIf InStr(UCase(sPartNumber), "MQP-U M12") <> 0 Then
        sConnectorBOM = " + 2, 216467, Hexagon nut M12 zinc plated + 2, 282852, Flat washer A 13/24 zinc plated"
    Else
        sConnectorBOM = " + 2, 216468, Hexagon nut M16 zinc plated + 2, 282853, Flat washer A 17/30 zinc plated"
    End If

    sConnType = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeGalv", "ConnType", lConnType)
    vInterfaceID = oAttribMetaData.iID("IJOAMQConnTypeGalv")

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

    sQTY = "1"

    If sConnType = "None" Then
        sBOMAdd = ""
    Else
        If InStr(UCase(sConnType), "CHANNEL") <> 0 Then

            If InStr(UCase(sChannelConnSize), "MQM M12") <> 0 Then
                sBOMAdd = " + " + sQTY + ", 216458, Hexagon Screw M12X25"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 F") <> 0 Then
                sBOMAdd = " + " + sQTY + ", 304766, Hexagon Nut M12-F"
            End If
            If InStr(UCase(sChannelConnSize), "MQM M12 R") <> 0 Then
                sBOMAdd = " + " + sQTY + ", 87634, Hexagon Screw M12X25 A4-70"
            End If

        Else

            If InStr(UCase(sConnType), "CONCRETE") <> 0 Then

                If InStr(UCase(sConcreteConn), "M12X95/25") <> 0 Then
                    sBOMAdd = " + " + sQTY + ", 311369, HVU-TZ M12"
                Else
                    If InStr(UCase(sConcreteConn), "M12X75/15") <> 0 Then
                        sBOMAdd = " + " + sQTY + ", 334539, HIT-HY 150/330/2"
                    Else
                        If InStr(UCase(sConcreteConn), "M12X110/28") <> 0 Then
                            sBOMAdd = " + " + sQTY + ", 310625, HVU-M12x110"
                        Else
                            sBOMAdd = ""
                        End If
                    End If
                End If

            Else
                sBOMAdd = " + " + sQTY + ", 58042, Washer A 10.5/20 A4 + " + sQTY + ", 52461, Hexagon Nut M10 A4-70"
            End If
        End If
    End If

    If InStr(UCase(sConnType), "NONE") <> 0 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & sConnectorBOM

    Else
        If InStr(UCase(sConnType), "CHANNEL") <> 0 Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sChannelConn & sBOMAdd & sConnectorBOM
        Else
            If InStr(UCase(sConnType), "CONCRETE") <> 0 Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sAnchorSize & sBOMAdd & sConnectorBOM
            Else
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + " & sQTY & ", " & sConnTypeID & " " & sSteelConnSize & sBOMAdd & sConnectorBOM
            End If
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
