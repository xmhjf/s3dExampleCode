VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQA_Inch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQA_Inch.cls
'   ProgID:         HS_Hilti_MQParts.MQA_Inch
'   Author:         SN
'   Creation Date:  Friday, November 02, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQA_Inch"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQA_Inch"

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
    Const iNumInputs As Integer = 11
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 5
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "Box1W"
    sDataType(2) = "Double"

    sInputName(3) = "Box1D"
    sDataType(3) = "Double"

    sInputName(4) = "Box1H"
    sDataType(4) = "Double"

    sInputName(5) = "Box2W"
    sDataType(5) = "Double"

    sInputName(6) = "Box2D"
    sDataType(6) = "Double"

    sInputName(7) = "Box2H"
    sDataType(7) = "Double"

    sInputName(8) = "RodL"
    sDataType(8) = "Double"

    sInputName(9) = "RodDia"
    sDataType(9) = "Double"

    sInputName(10) = "ItemNo"
    sDataType(10) = "String"

    sInputName(11) = "Group"
    sDataType(11) = "String"

    sOutputName(1) = "Structure1"
    sOutputName(2) = "Structure2"
    sOutputName(3) = "Box1"
    sOutputName(4) = "Box2"
    sOutputName(5) = "Rod"

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

    Dim dBox1W As Double
    Dim dBox1D As Double
    Dim dBox1H As Double
    Dim dBox2W As Double
    Dim dBox2D As Double
    Dim dBox2H As Double
    Dim dRodL As Double
    Dim dRodDia As Double
    Dim sItemNo As String
    Dim sGroup As String

    dBox1W = arrayOfInputs(2)
    dBox1D = arrayOfInputs(3)
    dBox1H = arrayOfInputs(4)
    dBox2W = arrayOfInputs(5)
    dBox2D = arrayOfInputs(6)
    dBox2H = arrayOfInputs(7)
    dRodL = arrayOfInputs(8)
    dRodDia = arrayOfInputs(9)
    sItemNo = arrayOfInputs(10)
    sGroup = arrayOfInputs(11)

    AddPort "Structure1", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure1", arrayOfInputs(1), "Structure1"

    AddPort "Structure2", 0, 0, dBox1H, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Structure2", arrayOfInputs(1), "Structure2"

    AddBox dBox1W, dBox1D, dBox1H, Loc(-dBox1W / 2#, -dBox1D / 2#, 0), m_outputColl, "Box1", "Box1"
    AddBox dBox2W, dBox2D, dBox2H, Loc(-dBox2W / 2#, -dBox2D / 2#, -dBox2H), m_outputColl, "Box2", "Box2"
    AddCylinder dRodL, dRodDia / 2, Loc(0, 0, dBox1H), m_outputColl, "Rod"

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

    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Set oPartOcc = Nothing
    Set oPart = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
