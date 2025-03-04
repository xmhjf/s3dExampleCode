VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RodConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQZ_L.cls
'   ProgID:         HS_Hilti_MQParts.MQZ_L
'   Author:         JRK
'   Creation Date:  Monday, November 19, 2007

'   Description:
'       Pre-assembled 90
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "RodConn"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.RodConn"

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

    ' Inputs
    Const iNumInputs As Integer = 7
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    ' Outputs
    Const iNumOutputs As Integer = 3
    Dim sOutputName(iNumOutputs) As String

    ' Inputs
    sInputName(2) = "Diameter1"
    sDataType(2) = "Double"

    sInputName(3) = "Thickness1"
    sDataType(3) = "Double"

    sInputName(4) = "PortOffset"
    sDataType(4) = "Double"

    sInputName(5) = "PortOffset1"
    sDataType(5) = "Double"

    sInputName(6) = "ItemNo"
    sDataType(6) = "String"

    sInputName(7) = "Group"
    sDataType(7) = "String"

    sOutputName(1) = "Rod"
    sOutputName(2) = "Rod2"
    sOutputName(3) = "Body"

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

    Dim dDiameter1 As Double
    Dim dThickness1 As Double
    Dim dPortOffset As Double
    Dim dPortOffset1 As Double
    Dim sItemNo As String
    Dim sGroup As String

    dDiameter1 = arrayOfInputs(2)
    dThickness1 = arrayOfInputs(3)
    dPortOffset = arrayOfInputs(4)
    dPortOffset1 = arrayOfInputs(5)
    sItemNo = arrayOfInputs(6)
    sGroup = arrayOfInputs(7)

    AddPort "Rod", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Rod", arrayOfInputs(1), "Rod"

    AddPort "Rod2", 0, 0, -dPortOffset + dPortOffset1, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Rod2", arrayOfInputs(1), "Rod2"

    AddCylinder dThickness1, dDiameter1 / 2, Loc(0, 0, -dPortOffset), m_outputColl, "Body", "Body"

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

    If Right(Trim(sPartNumber), 1) = "R" Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    Else
        If Right(Trim(sPartNumber), 1) = "F" Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp)
        Else
            bstrBOMDesc = Hilti_BuildBom(pSupportComp)
        End If
    End If

    Set oPartOcc = Nothing
    Set oPart = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub









