VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Roll_Connecter"
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
Private Const MODULE = "Roll_Connecter"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.Roll_Connecter"

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
    Const iNumInputs As Integer = 15
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    ' Outputs
    Const iNumOutputs As Integer = 5
    Dim sOutputName(iNumOutputs) As String

    ' Inputs
    sInputName(2) = "BaseL"
    sDataType(2) = "Double"

    sInputName(3) = "BaseW"
    sDataType(3) = "Double"

    sInputName(4) = "BaseThick"
    sDataType(4) = "Double"

    sInputName(5) = "TopL"
    sDataType(5) = "Double"

    sInputName(6) = "TopW"
    sDataType(6) = "Double"

    sInputName(7) = "TopThick"
    sDataType(7) = "Double"

    sInputName(8) = "RodL"
    sDataType(8) = "Double"

    sInputName(9) = "Rdia"
    sDataType(9) = "Double"

    sInputName(10) = "RodInset"
    sDataType(10) = "Double"

    sInputName(11) = "RodOffset"
    sDataType(11) = "Double"

    sInputName(12) = "MidWidth"
    sDataType(12) = "Double"

    sInputName(13) = "DistanceBetween"
    sDataType(13) = "Double"

    sInputName(14) = "ItemNo"
    sDataType(14) = "String"

    sInputName(15) = "Group"
    sDataType(15) = "String"

    sOutputName(1) = "Base"
    sOutputName(2) = "Rod"
    sOutputName(3) = "Bottom"
    sOutputName(4) = "Top"
    sOutputName(5) = "RodShape"

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

    Dim dBaseL As Double
    Dim dBaseW As Double
    Dim dBaseThick As Double
    Dim dTopL As Double
    Dim dTopW As Double
    Dim dTopThick As Double
    Dim dRodL As Double
    Dim dRodDia As Double
    Dim dRodInset As Double
    Dim dRodOffset As Double
    Dim dMidWidth As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim dDistanceBetween As Double

    dBaseL = arrayOfInputs(2)
    dBaseW = arrayOfInputs(3)
    dBaseThick = arrayOfInputs(4)
    dTopL = arrayOfInputs(5)
    dTopW = arrayOfInputs(6)
    dTopThick = arrayOfInputs(7)
    dRodL = arrayOfInputs(8)
    dRodDia = arrayOfInputs(9)
    dRodInset = arrayOfInputs(10)
    dRodOffset = arrayOfInputs(11)
    dMidWidth = arrayOfInputs(12)
    dDistanceBetween = arrayOfInputs(13)
    sItemNo = arrayOfInputs(14)
    sGroup = arrayOfInputs(15)

    AddPort "Base", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Base", arrayOfInputs(1), "Base"

    AddPort "Rod", dRodOffset, 0, dBaseThick + dTopThick - dRodInset + dRodL, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Rod", arrayOfInputs(1), "Rod"


    AddPolygon 6, "EXTRUDED", 0, "Bottom"
    AddLine dBaseL / 2, dMidWidth / 2, 0, 0, dBaseW / 2, 0, m_outputColl, "", "Bottom1"
    AddLine 0, dBaseW / 2, 0, -dBaseL / 2, dMidWidth / 2, 0, m_outputColl, "", "Bottom2"
    AddLine -dBaseL / 2, dMidWidth / 2, 0, -dBaseL / 2, -dMidWidth / 2, 0, m_outputColl, "", "Bottom3"
    AddLine -dBaseL / 2, -dMidWidth / 2, 0, 0, -dBaseW / 2, 0, m_outputColl, "", "Bottom4"
    AddLine 0, -dBaseW / 2, 0, dBaseL / 2, -dMidWidth / 2, 0, m_outputColl, "", "Bottom5"
    AddLine dBaseL / 2, -dMidWidth / 2, 0, dBaseL / 2, dMidWidth / 2, 0, m_outputColl, "", "Bottom6"
    AddExtrusion 0, 0, dBaseThick, 1, m_outputColl, "Bottom"

    AddBox dTopL, dTopW, dTopThick, Loc(-dTopL / 2, -dTopW / 2, dBaseThick), m_outputColl, "Top", "Top"

    AddCylinder dRodL, dRodDia / 2, Loc(dRodOffset, 0, dBaseThick + dTopThick - dRodInset), m_outputColl, "RodShape", "RodShape"

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







