VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQV_3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQV_2D.cls
'   ProgID:         HS_Hilti_MQParts.MQV_2D
'   Author:         JRK
'   Creation Date:  Monday, November 19, 2007

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
Private Const MODULE = "MQV_3D"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQV_3D"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices

Dim sPartNumber As String

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
    Const iNumInputs As Integer = 7
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 6
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "L"
    sDataType(2) = "Double"

    sInputName(3) = "H"
    sDataType(3) = "Double"

    sInputName(4) = "W"
    sDataType(4) = "Double"

    sInputName(5) = "T"
    sDataType(5) = "Double"

    sInputName(6) = "ItemNo"
    sDataType(6) = "String"

    sInputName(7) = "Group"
    sDataType(7) = "String"

    sOutputName(1) = "Channel1"
    sOutputName(2) = "Channel2"
    sOutputName(3) = "Channel3"
    sOutputName(4) = "Body"
    sOutputName(5) = "LSectionLeft"
    sOutputName(6) = "LSectionFront"

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

    Dim dW As Double
    Dim dT As Double
    Dim dH As Double
    Dim dL As Double

    dL = arrayOfInputs(2)
    dH = arrayOfInputs(3)
    dW = arrayOfInputs(4)
    dT = arrayOfInputs(5)

    AddPort "Channel1", 0, 0, 0, _
            0, -1, 0, _
            0, 0, -1, m_outputColl, "Channel1", arrayOfInputs(1), "Channel1"

    AddPort "Channel2", -dW / 2, 0, dW / 2, _
            0, 1, 0, _
            -1, 0, 0, m_outputColl, "Channel2", arrayOfInputs(1), "Channel2"

    AddPort "Channel3", 0, -dW / 2, dW / 2, _
            1, 0, 0, _
            0, -1, 0, m_outputColl, "Channel3", arrayOfInputs(1), "Channel3"

    AddPolygon 6, "EXTRUDED", 0, "POLY"
    AddLine dW / 2, -dW / 2 - dT, -dH - dT, -dW / 2 - dT, -dW / 2 - dT, -dH - dT, m_outputColl, "", "Poly1"
    AddLine -dW / 2 - dT, -dW / 2 - dT, -dH - dT, -dW / 2 - dT, dW / 2, -dH - dT, m_outputColl, "", "Poly2"
    AddLine -dW / 2 - dT, dW / 2, -dH - dT, -dW / 2, dW / 2, -dH - dT, m_outputColl, "", "Poly3"
    AddLine -dW / 2, dW / 2, -dH - dT, -dW / 2, -dW / 2, -dH - dT, m_outputColl, "", "Poly4"
    AddLine -dW / 2, -dW / 2, -dH - dT, dW / 2, -dW / 2, -dH - dT, m_outputColl, "", "Poly5"
    AddLine dW / 2, -dW / 2, -dH - dT, dW / 2, -dW / 2 - dT, -dH - dT, m_outputColl, "", "Poly6"
    AddExtrusion 0, 0, dH, 1, m_outputColl, "Body"

    AddPolygon 6, "EXTRUDED", 0, "LSectionLeft"
    AddLine -dL - dW / 2 - dT, dW / 2, -dT, -dL - dW / 2 - dT, -dW / 2 - dT, -dT, m_outputColl, "", "LSectionLeft1"
    AddLine -dL - dW / 2 - dT, -dW / 2 - dT, -dT, -dL - dW / 2 - dT, -dW / 2 - dT, dW, m_outputColl, "", "LSectionLeft2"
    AddLine -dL - dW / 2 - dT, -dW / 2 - dT, dW, -dL - dW / 2 - dT, -dW / 2, dW, m_outputColl, "", "LSectionLeft3"
    AddLine -dL - dW / 2 - dT, -dW / 2, dW, -dL - dW / 2 - dT, -dW / 2, 0, m_outputColl, "", "LSectionLeft4"
    AddLine -dL - dW / 2 - dT, -dW / 2, 0, -dL - dW / 2 - dT, dW / 2, 0, m_outputColl, "", "LSectionLeft5"
    AddLine -dL - dW / 2 - dT, dW / 2, 0, -dL - dW / 2 - dT, dW / 2, -dT, m_outputColl, "", "LSectionLeft6"
    AddExtrusion dL, 0, 0, 1, m_outputColl, "LSectionLeft"

    AddPolygon 6, "EXTRUDED", 0, "LSectionFront"
    AddLine dW / 2, -dL - dW / 2 - dT, -dT, -dW / 2 - dT, -dL - dW / 2 - dT, -dT, m_outputColl, "", "LSectionFront"
    AddLine -dW / 2 - dT, -dL - dW / 2 - dT, -dT, -dW / 2 - dT, -dL - dW / 2 - dT, dW, m_outputColl, "", "LSectionFront2"
    AddLine -dW / 2 - dT, -dL - dW / 2 - dT, dW, -dW / 2, -dL - dW / 2 - dT, dW, m_outputColl, "", "LSectionFront3"
    AddLine -dW / 2, -dL - dW / 2 - dT, dW, -dW / 2, -dL - dW / 2 - dT, 0, m_outputColl, "", "LSectionFront4"
    AddLine -dW / 2, -dL - dW / 2 - dT, 0, dW / 2, -dL - dW / 2 - dT, 0, m_outputColl, "", "LSectionFront5"
    AddLine dW / 2, -dL - dW / 2 - dT, 0, dW / 2, -dL - dW / 2 - dT, -dT, m_outputColl, "", "LSectionFront6"
    AddExtrusion 0, dL, 0, 1, m_outputColl, "LSectionFront"

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
    Dim dWeight As Double

    Dim sItemNo As Long
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sPartNumber = oPart.PartNumber

    If Right(Trim(sPartNumber), 1) = "R" Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 304012, Channel connector MQN-R"
    ElseIf Right(Trim(sPartNumber), 1) = "F" Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 387779, Channel Connector MQN-HDG plus"
    Else
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 369623, Channel connector MQN"
    End If
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub






