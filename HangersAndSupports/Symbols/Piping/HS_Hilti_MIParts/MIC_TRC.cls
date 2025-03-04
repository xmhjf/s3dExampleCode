VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIC_TRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Anvil_FIG95.cls
'   ProgID:         HS_Hilti_MIParts.MIC_TRC
'   Author:         JOGN
'   Creation Date: October 10, 2007
'
'   Description:
'       Crossbeam connector
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP          Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MIC_TRC"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MIC_TRC"

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
    Const iNumInputs As Integer = 13
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 4
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "PlateThick"
    sDataType(2) = "Double"

    sInputName(3) = "ItemNo"
    sDataType(3) = "String"

    sInputName(4) = "Finish"
    sDataType(4) = "Long"

    sInputName(5) = "Group"
    sDataType(5) = "String"

    sInputName(6) = "RodDia"
    sDataType(6) = "Double"

    sInputName(7) = "Height"
    sDataType(7) = "Double"

    sInputName(8) = "HeightOffset"
    sDataType(8) = "Double"

    sInputName(9) = "PartWidth"
    sDataType(9) = "Double"

    sInputName(10) = "WidthOffset"
    sDataType(10) = "Double"

    sInputName(11) = "PartDepth"
    sDataType(11) = "Double"

    sInputName(12) = "RodOffset"
    sDataType(12) = "Double"

    sInputName(13) = "PortOffset"
    sDataType(13) = "Double"

    'Outputs
    sOutputName(1) = "Rod Hole"
    sOutputName(2) = "Structure"
    sOutputName(3) = "Body"
    sOutputName(4) = "Rod"

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
        Call Physical(arrayOfInputs)
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
Private Function Physical(ByRef arrayOfInputs())
    Const METHOD = "Physical"
    On Error GoTo ErrHandler

    Dim dPlateThick As Double
    Dim sItemNo As String
    Dim lFinish As Long
    Dim sGroup As String
    Dim dRodDia As Double
    Dim dHeight As Double
    Dim dHeightOffset As Double
    Dim dWidth As Double
    Dim dWidthOffset As Double
    Dim dRodOffset As Double
    Dim dDepth As Double
    Dim dPortOffset As Double

    dPlateThick = arrayOfInputs(2)
    sItemNo = arrayOfInputs(3)
    lFinish = arrayOfInputs(4)
    sGroup = arrayOfInputs(5)
    dRodDia = arrayOfInputs(6)
    dHeight = arrayOfInputs(7)
    dHeightOffset = arrayOfInputs(8)
    dWidth = arrayOfInputs(9)
    dWidthOffset = arrayOfInputs(10)
    dDepth = arrayOfInputs(11)
    dRodOffset = arrayOfInputs(12)
    dPortOffset = arrayOfInputs(13)

    AddPort "Rod Hole", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Rod Hole", arrayOfInputs(1), "Rod Hole"

    AddPort "Structure", 0, 0, dHeight - dRodOffset - dPortOffset, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"

    AddCylinder dPlateThick * 4 + dWidth, dRodDia / 2, Loc(0, 0, -dWidth / 2 - (dPlateThick * 2)) + " * ROTX(-90) ", m_outputColl, "Rod"

    AddPolygon 12, "EXTRUDED", "", "POLY"
    AddPoint -dDepth / 2, dWidth / 2, -dRodOffset, m_outputColl, "", "POLYPT0"
    AddPoint -dDepth / 2, dWidth / 2 + dPlateThick, -dRodOffset, m_outputColl, "", "POLYPT1"
    AddPoint -dDepth / 2, dWidth / 2 + dPlateThick, dHeight - dHeightOffset - dRodOffset, m_outputColl, "", "POLYPT2"
    AddPoint -dDepth / 2, dWidth / 2 - dWidthOffset, dHeight + dPlateThick - dRodOffset, m_outputColl, "", "POLYPT3"
    AddPoint -dDepth / 2, -dWidth / 2 + dWidthOffset, dHeight + dPlateThick - dRodOffset, m_outputColl, "", "POLYPT4"
    AddPoint -dDepth / 2, -dWidth / 2 - dPlateThick, dHeight - dHeightOffset - dRodOffset, m_outputColl, "", "POLYPT5"
    AddPoint -dDepth / 2, -dWidth / 2 - dPlateThick, -dRodOffset, m_outputColl, "", "POLYPT6"
    AddPoint -dDepth / 2, -dWidth / 2, -dRodOffset, m_outputColl, "", "POLYPT7"
    AddPoint -dDepth / 2, -dWidth / 2, dHeight - dHeightOffset - dRodOffset, m_outputColl, "", "POLYPT8"
    AddPoint -dDepth / 2, -dWidth / 2 + dWidthOffset, dHeight - dRodOffset, m_outputColl, "", "POLYPT9"
    AddPoint -dDepth / 2, dWidth / 2 - dWidthOffset, dHeight - dRodOffset, m_outputColl, "", "POLYPT10"
    AddPoint -dDepth / 2, dWidth / 2, dHeight - dHeightOffset - dRodOffset, m_outputColl, "", "POLYPT11"
    AddExtrusion dDepth, 0, 0, 1, m_outputColl, "Body", "Body"


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

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub








