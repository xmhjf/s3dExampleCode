VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIC_PA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Anvil_FIG95.cls
'   ProgID:         HS_Hilti_MIParts.MIC_PA
'   Author:         JOGN
'   Creation Date: October 9, 2007
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
Private Const MODULE = "MIC_PA"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MIC_PA"

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
    Const iNumOutputs As Integer = 9
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "PlateThick"
    sDataType(2) = "Double"

    sInputName(3) = "BoltDia"
    sDataType(3) = "Double"

    sInputName(4) = "PlateDepth"
    sDataType(4) = "Double"

    sInputName(5) = "PlateWidth"
    sDataType(5) = "Double"

    sInputName(6) = "GirderH"
    sDataType(6) = "Double"

    sInputName(7) = "ItemNo"
    sDataType(7) = "String"

    sInputName(8) = "Finish"
    sDataType(8) = "Long"

    sInputName(9) = "Group"
    sDataType(9) = "String"

    'Outputs
    sOutputName(1) = "StructureInt"
    sOutputName(2) = "StructureExt"
    sOutputName(3) = "Plate"
    sOutputName(4) = "Clamp1Right"
    sOutputName(5) = "Clamp1Left"
    sOutputName(6) = "Clamp1Bottom"
    sOutputName(7) = "Clamp2Right"
    sOutputName(8) = "Clamp2Left"
    sOutputName(9) = "Clamp2Bottom"

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
    Dim dBoltDia As Double
    Dim dPlateDepth As Double
    Dim dPlateWidth As Double
    Dim dGirderH As Double
    Dim sItemNo As String
    Dim lFinish As Long
    Dim sGroup As String
    Dim dVertBoltLength As Double
    Dim dHorzBoltLength As Double

    dPlateThick = arrayOfInputs(2)
    dBoltDia = arrayOfInputs(3)
    dPlateDepth = arrayOfInputs(4)
    dPlateWidth = arrayOfInputs(5)
    dGirderH = arrayOfInputs(6)
    sItemNo = arrayOfInputs(7)
    lFinish = arrayOfInputs(8)
    sGroup = arrayOfInputs(9)

    'Extra Connector Data
    Dim dP_W As Double
    Dim dP_W1 As Double
    Dim dP_T As Double
    Dim dP_H As Double
    Dim dP_D As Double
    Dim dB_DIA As Double
    Dim dB_CC As Double
    Dim dB_EXTRA As Double
    Dim dB_H As Double
    Dim dB_H2 As Double
    Dim dB_2H As Double

    dVertBoltLength = 2 * dPlateThick + dGirderH + dBoltDia
    dHorzBoltLength = 0.09 + dBoltDia

    AddPort "StructureInt", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureInt", arrayOfInputs(1), "StructureInt"

    AddPort "StructureExt", 0, 0, dGirderH / 2 + dPlateThick, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureExt", arrayOfInputs(1), "StructureExt"

    AddBox dPlateDepth, dPlateWidth, dPlateThick, Loc(-dPlateDepth / 2, -dPlateWidth / 2, dGirderH / 2), m_outputColl, "Plate"

    AddCylinder dVertBoltLength, dBoltDia / 2, "ROTZ(270)*" + Loc(-dPlateDepth / 2 + dBoltDia, -dHorzBoltLength / 2, -dGirderH / 2 - dBoltDia), m_outputColl, "Clamp1Right"
    AddCylinder dVertBoltLength, dBoltDia / 2, "ROTZ(270)*" + Loc(-dPlateDepth / 2 + dBoltDia, dHorzBoltLength / 2, -dGirderH / 2 - dBoltDia), m_outputColl, "Clamp1Left"
    AddCylinder dHorzBoltLength + dBoltDia, dBoltDia / 2, "ROTX(-90)*" + Loc(-dPlateDepth / 2 + dBoltDia, -dHorzBoltLength / 2 - dBoltDia / 2, -dVertBoltLength / 2 + dPlateThick), m_outputColl, "Clamp1Bottom"

    AddCylinder dVertBoltLength, dBoltDia / 2, "ROTZ(270)*" + Loc(dPlateDepth / 2 - dBoltDia, -dHorzBoltLength / 2, -dGirderH / 2 - dBoltDia), m_outputColl, "Clamp2Right"
    AddCylinder dVertBoltLength, dBoltDia / 2, "ROTZ(270)*" + Loc(dPlateDepth / 2 - dBoltDia, dHorzBoltLength / 2, -dGirderH / 2 - dBoltDia), m_outputColl, "Clamp2Left"
    AddCylinder dHorzBoltLength + dBoltDia, dBoltDia / 2, "ROTX(-90)*" + Loc(dPlateDepth / 2 - dBoltDia, -dHorzBoltLength / 2 - dBoltDia / 2, -dVertBoltLength / 2 + dPlateThick), m_outputColl, "Clamp2Bottom"

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








