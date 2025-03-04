VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIC_T"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Anvil_FIG95.cls
'   ProgID:         HS_Hilti_MIParts.MIC_T
'   Author:         JOGN
'   Creation Date: October 3, 2007
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
Private Const MODULE = "MIC_T"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MIC_T"

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
    Const iNumOutputs As Integer = 7
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

    sInputName(6) = "TopDepth"
    sDataType(6) = "Double"

    sInputName(7) = "TopWidth"
    sDataType(7) = "Double"

    sInputName(8) = "BottomDepth"
    sDataType(8) = "Double"

    sInputName(9) = "BottomWidth"
    sDataType(9) = "Double"

    sInputName(10) = "GirderWidth"
    sDataType(10) = "Long"

    sInputName(11) = "GirderHeight"
    sDataType(11) = "Long"

    sInputName(12) = "PortOffset"
    sDataType(12) = "Double"

    sInputName(13) = "CP_Offset"
    sDataType(13) = "Double"

    'Outputs
    sOutputName(1) = "StructureHorz"
    sOutputName(2) = "StructureVert"
    sOutputName(3) = "StructureCase"
    sOutputName(4) = "TopRight"
    sOutputName(5) = "BottomRight"
    sOutputName(6) = "TopLeft"
    sOutputName(7) = "BottomLeft"


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
    Dim dTopDepth As Double
    Dim dTopWidth As Double
    Dim dBottomDepth As Double
    Dim dBottomWidth As Double
    Dim lGirderWidth As Long
    Dim lGirderHeight As Long
    Dim dPortOffset As Double
    Dim dGirderWidth As Double
    Dim dGirderHeight As Double
    Dim dGap As Double
    Dim dCP_Offset As Double
    Dim sErrMsg As String

    dPlateThick = arrayOfInputs(2)
    sItemNo = arrayOfInputs(3)
    lFinish = arrayOfInputs(4)
    sGroup = arrayOfInputs(5)
    dTopDepth = arrayOfInputs(6)
    dTopWidth = arrayOfInputs(7)
    dBottomDepth = arrayOfInputs(8)
    dBottomWidth = arrayOfInputs(9)
    lGirderWidth = arrayOfInputs(10)
    lGirderHeight = arrayOfInputs(11)
    dPortOffset = arrayOfInputs(12)
    dCP_Offset = arrayOfInputs(13)

    If lGirderHeight = lGirderWidth Then
        If lGirderHeight = 2 Then
            sErrMsg = "Girder Height and Girder Width can not both equal 120, please select different height or width."
            PF_EventHandler sErrMsg, Err, MODULE, METHOD, True

            dPortOffset = 0.045
            dGap = 0.09
        Else
            dPortOffset = 0.045
            dGap = 0.09
        End If
    Else
        If lGirderHeight = 1 Then
            dPortOffset = 0.045
            dGap = 0.12
        Else
            dPortOffset = 0.06
            dGap = 0.09
        End If
    End If

    AddPort "StructureHorz", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureHorz", arrayOfInputs(1), "StructureHorz"

    AddPort "StructureVert", 0, 0, -dPortOffset - dCP_Offset, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureVert", arrayOfInputs(1), "StructureVert"

    AddPort "StructureCase", 0, 0, -dPortOffset, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureCase", arrayOfInputs(1), "StructureCase"

    AddBox dTopDepth, dPlateThick, dTopWidth, Loc(-dTopDepth / 2, dGap / 2, -dTopWidth / 2), m_outputColl, "TopRight"
    AddBox dTopDepth, dPlateThick, dTopWidth, Loc(-dTopDepth / 2, -dGap / 2 - dPlateThick, -dTopWidth / 2), m_outputColl, "TopLeft"
    AddBox dBottomDepth, dPlateThick, dBottomWidth, Loc(-dBottomDepth / 2, dGap / 2, -dTopWidth / 2 - dBottomWidth), m_outputColl, "BottomRight"
    AddBox dBottomDepth, dPlateThick, dBottomWidth, Loc(-dBottomDepth / 2, -dGap / 2 - dPlateThick, -dTopWidth / 2 - dBottomWidth), m_outputColl, "BottomLeft"

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
    Dim sPartNumber As String

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub




