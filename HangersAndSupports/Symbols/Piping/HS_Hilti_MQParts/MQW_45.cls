VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQW_45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQW_45.cls
'   ProgID:         HS_Hilti_MQParts.MQW_45
'   Author:         JOGN
'   Creation Date:  Wednesday November 14, 2007

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
Private Const MODULE = "MQW_45"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQW_45"

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
    Const iNumInputs As Integer = 8
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 3
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "AngleLength"
    sDataType(2) = "Double"

    sInputName(3) = "AngleHeight"
    sDataType(3) = "Double"

    sInputName(4) = "AngleWidth"
    sDataType(4) = "Double"

    sInputName(5) = "AngleThick"
    sDataType(5) = "Double"

    sInputName(6) = "NumPush"
    sDataType(6) = "Long"

    sInputName(7) = "ItemNo"
    sDataType(7) = "String"

    sInputName(8) = "Group"
    sDataType(8) = "String"

    sOutputName(1) = "VertStructure"
    sOutputName(2) = "Body"
    sOutputName(3) = "Body2"

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

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Dim sPartNumber As String
    Set oPartOcc = sblOcc

    Dim dAngleLength As Double
    Dim dAngleHeight As Double
    Dim dAngleWidth As Double
    Dim dAngleThick As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim lNumPush As Long
    Dim dCalc1 As Double
    Dim dCalc2 As Double
    Dim dCalc3 As Double
    Dim dCalc4 As Double
    Dim dCalc5 As Double
    Dim dCalc6 As Double
    Dim dCalc7 As Double
    Dim dCalc8 As Double
    Dim dCalc9 As Double
    Dim dCalc10 As Double

    dAngleLength = arrayOfInputs(2)
    dAngleHeight = arrayOfInputs(3)
    dAngleWidth = arrayOfInputs(4)
    dAngleThick = arrayOfInputs(5)
    lNumPush = arrayOfInputs(6)
    sItemNo = arrayOfInputs(7)
    sGroup = arrayOfInputs(8)

    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    Set oSupportComp = Nothing

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        sPartNumber = ""
    Else
        oPartOcc.GetPart oPart
        sPartNumber = oPart.PartNumber
    End If
    Set oSupportComp = Nothing

    If InStr(sPartNumber, "8/45") > 0 Then
        AddPort "VertStructure", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "VertStructure", arrayOfInputs(1), "VertStructure"

        dCalc1 = CosDeg(45) * (dAngleWidth + dAngleThick)
        dCalc2 = SinDeg(45) * (dAngleWidth + dAngleThick)
        dCalc3 = SinDeg(45) * dAngleThick
        dCalc4 = CosDeg(45) * dAngleThick
        dCalc5 = SinDeg(45) * dAngleLength

        AddPolygon 6, "EXTRUDED", 0, "POLY1"
        AddPoint 0, dAngleWidth / 2, 0, m_outputColl, "", "PollyPT11"
        AddPoint 0, dAngleWidth / 2, -dAngleWidth, m_outputColl, "", "PollyPT12"
        AddPoint 0, dAngleWidth / 2 + dAngleThick, -dAngleWidth, m_outputColl, "", "PollyPT13"
        AddPoint 0, dAngleWidth / 2 + dAngleThick, dAngleThick, m_outputColl, "", "PollyPT14"
        AddPoint 0, -dAngleWidth / 2, dAngleThick, m_outputColl, "", "PollyPT15"
        AddPoint 0, -dAngleWidth / 2, 0, m_outputColl, "", "PollyPT16"
        AddExtrusion dAngleLength, 0, 0, 1, m_outputColl, "Body"

        AddPolygon 6, "EXTRUDED", 0, "POLY2"
        AddPoint 0, -dAngleWidth / 2, dAngleThick, m_outputColl, "", "PollyPT21"
        AddPoint 0, dAngleWidth / 2 + dAngleThick, dAngleThick, m_outputColl, "", "PollyPT22"
        AddPoint -dCalc1, dAngleWidth / 2 + dAngleThick, dAngleThick - dCalc2, m_outputColl, "", "PollyPT23"
        AddPoint -dCalc1, dAngleWidth / 2, dAngleThick - dCalc2, m_outputColl, "", "PollyPT24"
        AddPoint -dCalc3, dAngleWidth / 2, dAngleThick - dCalc4, m_outputColl, "", "PollyPT25"
        AddPoint -dCalc3, -dAngleWidth / 2, dAngleThick - dCalc4, m_outputColl, "", "PollyPT26"
        AddExtrusion -dCalc5, 0, dCalc5, 1, m_outputColl, "Body2"
    Else
        If InStr(sPartNumber, "3/135") > 0 Then
            AddPort "VertStructure", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "VertStructure", arrayOfInputs(1), "VertStructure"
                
        dCalc1 = CosDeg(45) * (dAngleWidth + dAngleThick)
        dCalc2 = SinDeg(45) * (dAngleWidth + dAngleThick)
        dCalc3 = SinDeg(45) * dAngleThick
        dCalc4 = CosDeg(45) * dAngleThick
        dCalc5 = SinDeg(45) * dAngleHeight
        
        dCalc7 = CosDeg(45) * (dAngleThick)
        dCalc8 = CosDeg(45) * (dAngleHeight + dAngleThick)
        dCalc9 = CosDeg(45) * (dAngleHeight)


        AddPolygon 6, "EXTRUDED", 0, "POLY1"
        AddPoint 0, dAngleWidth / 2, 0, m_outputColl, "", "PollyPT11"
        AddPoint 0, dAngleWidth / 2, -dAngleWidth, m_outputColl, "", "PollyPT12"
        AddPoint 0, dAngleWidth / 2 + dAngleThick, -dAngleWidth, m_outputColl, "", "PollyPT13"
        AddPoint 0, dAngleWidth / 2 + dAngleThick, dAngleThick, m_outputColl, "", "PollyPT14"
        AddPoint 0, -dAngleWidth / 2, dAngleThick, m_outputColl, "", "PollyPT15"
        AddPoint 0, -dAngleWidth / 2, 0, m_outputColl, "", "PollyPT16"
        AddExtrusion dAngleLength, 0, 0, 1, m_outputColl, "Body"
        
        
        AddPolygon 4, "EXTRUDED", 0, "POLY2"
        AddPoint -dCalc7, dAngleWidth / 2 + dAngleThick, (dAngleThick - dCalc7), m_outputColl, "", "PollyPT21"
        AddPoint -dCalc8, dAngleWidth / 2 + dAngleThick, (dCalc9 + (dAngleThick - dCalc7)), m_outputColl, "", "PollyPT22"
        AddPoint -dCalc9, dAngleWidth / 2 + dAngleThick, (dCalc8 + (dAngleThick - dCalc7)), m_outputColl, "", "PollyPT23"
        AddPoint -0, dAngleWidth / 2 + dAngleThick, dAngleThick, m_outputColl, "", "PollyPT24"
        AddExtrusion 0, -dAngleWidth - dAngleThick, 0, 1, m_outputColl, "Body2"
   
    Else

        AddPort "VertStructure", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "VertStructure", arrayOfInputs(1), "VertStructure"

        If InStr(sPartNumber, "3/45") > 0 Then
            dCalc1 = Sqr(((dAngleThick ^ 2) / 2))
            dCalc2 = SinDeg(45) * dAngleLength
            dCalc3 = CosDeg(45) * dAngleLength
            dCalc4 = TanDeg(22.5) * dAngleThick

            AddPolygon 6, "EXTRUDED", 0, "POLY"
            AddPoint 0, -dAngleWidth / 2#, 0, m_outputColl, "", "POLYPT0"
            AddPoint dAngleHeight, -dAngleWidth / 2#, 0, m_outputColl, "", "POLYPT1"
            AddPoint dAngleHeight, -dAngleWidth / 2#, dAngleThick, m_outputColl, "", "POLYPT2"
            AddPoint dCalc4, -dAngleWidth / 2#, dAngleThick, m_outputColl, "", "POLYPT3"
            AddPoint -dCalc3 + dCalc1, -dAngleWidth / 2#, dCalc2 + dCalc1, m_outputColl, "", "POLYPT4"
            AddPoint -dCalc3, -dAngleWidth / 2, dCalc2, m_outputColl, "", "POLYPT5"
            AddExtrusion 0, dAngleWidth, 0, 1, m_outputColl, "Body"
        Else
            dCalc1 = TanDeg(22.5) * dAngleThick
            dCalc2 = CosDeg(45) * dAngleLength
            dCalc3 = SinDeg(45) * dAngleLength
            dCalc4 = Sqr(((dAngleThick ^ 2) / 2))

            AddPolygon 8, "EXTRUDED", 0, "POLY"
            AddPoint 0, -dAngleWidth / 2#, 0, m_outputColl, "", "POLYPT0"
            AddPoint dAngleLength, -dAngleWidth / 2#, 0, m_outputColl, "", "POLYPT1"
            AddPoint dAngleLength, -dAngleWidth / 2#, dAngleThick, m_outputColl, "", "POLYPT2"
            AddPoint 0, -dAngleWidth / 2#, dAngleThick, m_outputColl, "", "POLYPT3"
            AddPoint dCalc2, -dAngleWidth / 2#, dCalc3 + dAngleThick, m_outputColl, "", "POLYPT4"
            AddPoint dCalc2 - dCalc4, -dAngleWidth / 2#, dCalc3 + dCalc4 + dAngleThick, m_outputColl, "", "POLYPT5"
            AddPoint -dCalc4, -dAngleWidth / 2#, dCalc4 + dAngleThick, m_outputColl, "", "POLYPT6"
            AddPoint 0, -dAngleWidth / 2#, dAngleThick, m_outputColl, "", "POLYPT7"
            AddExtrusion 0, dAngleWidth, 0, 1, m_outputColl, "Body"
        End If

    End If
    
    End If

    Set oPartOcc = Nothing

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
    Dim sPartNumber As String
    Dim dNumPush As Double
    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    dNumPush = Val(GetAttributeFromObject(oPart, "NumPush"))

    If InStr(UCase(sPartNumber), "8/45") <> 0 Then
        If Right(Trim(sPartNumber), 1) = "R" Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 304012, Channel connector MQN-R"
        Else
            If Right(Trim(sPartNumber), 1) = "F" Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 387779, Channel Connector MQN-HDG plus"
            Else
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 369623, Channel connector MQN"
            End If
        End If
    Else
        If Right(Trim(sPartNumber), 1) = "R" Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 3, 304012, Channel connector MQN-R"
        Else
            If Right(Trim(sPartNumber), 1) = "F" Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 3, 387779, Channel Connector MQN-HDG plus"
            Else
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 3, 369623, Channel connector MQN"
            End If
        End If
    End If

    Set oPartOcc = Nothing
    Set oPart = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
