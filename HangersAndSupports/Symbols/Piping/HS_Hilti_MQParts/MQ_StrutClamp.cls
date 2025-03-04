VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQ_StrutClamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQ_StrutClamp.cls
'   ProgID:         HS_Hilti_MQParts.MQ_StrutClamp
'   Author:         JRK
'   Creation Date:  Wednesday, November 7, 2007

'   Description:
'       Strut Clamp
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQ_StrutClamp"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQ_StrutClamp"

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
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "H", "Double"
    AddInput Inputs, "B", "Double"
    AddInput Inputs, "Depth", "Double"
    AddInput Inputs, "S", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Body1"
    AddOutput Outputs, "BoxLeft"
    AddOutput Outputs, "BoxRight"
    AddOutput Outputs, "BoltHead"
    AddOutput Outputs, "BoltThread"
    AddOutput Outputs, "Nut"
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

    Dim dD As Double
    Dim dH As Double
    Dim dB As Double
    Dim dDepth As Double
    Dim dS As Double
    Dim dBHS As Double
    Dim dSpacer As Double
    Dim dFlangeH As Double
    Dim dBoltK As Double
    Dim dBoltE As Double
    Dim dBoltH As Double
    Dim dBoltD As Double

    dD = arrayOfInputs(2)
    dH = arrayOfInputs(3)
    dB = arrayOfInputs(4)
    dDepth = arrayOfInputs(5)
    dS = arrayOfInputs(6)
    dBHS = dB / 2 - dS

    AddPort "Route", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"

    AddComposite 8, "EXTRUDED", 0, "POLY"
    AddLine -dDepth / 2#, -dBHS - dS, -dD / 2, -dDepth / 2, -dBHS - dS, 0, m_outputColl, "", "OuterLeft"
    AddArc 90, 270, dBHS + dS, "ROTY(90)*" & Loc(-dDepth / 2, 0, 0), m_outputColl, "", "OuterArc"
    AddLine -dDepth / 2#, dBHS + dS, 0, -dDepth / 2#, dBHS + dS, -dD / 2#, m_outputColl, "", "OuterRight"
    AddLine -dDepth / 2#, dBHS + dS, -dD / 2, -dDepth / 2#, dBHS, -dD / 2#, m_outputColl, "", "BotRight"
    AddLine -dDepth / 2#, dBHS, -dD / 2, -dDepth / 2#, dBHS, 0, m_outputColl, "", "InnerRight"
    AddArc 90, 270, dBHS, "ROTY(90)*" & Loc(-dDepth / 2, 0, 0), m_outputColl, "", "InnerArc"
    AddLine -dDepth / 2#, -dBHS, 0, -dDepth / 2#, -dBHS, -dD / 2#, m_outputColl, "", "InnerLeft"
    AddLine -dDepth / 2#, -dBHS, -dD / 2, -dDepth / 2#, -dBHS - dS, -dD / 2#, m_outputColl, "", "BotLeft"
    AddExtrusion dDepth, 0, 0, 1, m_outputColl, "Body1", "Body1"

    dSpacer = 0.005
    dFlangeH = dH - 0.014 - dB

    AddBox dDepth, dS, dFlangeH + dS, Loc(-dDepth / 2#, -dSpacer / 2# - dS, dB / 2# - dS), m_outputColl, "BoxLeft", "BoxLeft"
    AddBox dDepth, dS, dFlangeH + dS, Loc(-dDepth / 2#, dSpacer / 2#, dB / 2# - dS), m_outputColl, "BoxRight", "BoxRight"

    dBoltK = 0.0053
    dBoltE = 0.01438
    dBoltH = 0.02
    dBoltD = 0.008

    AddCylinder dBoltK, dBoltE / 2#, "ROTX(90)*" + Loc(0, dSpacer / 2# + dS + dBoltK, dB / 2# + dFlangeH / 2#), m_outputColl, "BoltHead", "BoltHead"
    AddCylinder dBoltH, dBoltD / 2#, "ROTX(90)*" + Loc(0, dBoltH / 2, dB / 2# + dFlangeH / 2#), m_outputColl, "BoltThread", "BoltThread"
    AddCylinder dBoltK, dBoltE / 2#, "ROTX(90)*" + Loc(0, -dSpacer / 2# - dS, dB / 2# + dFlangeH / 2#), m_outputColl, "Nut", "Nut"

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

    bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "")

    Set oPartOcc = Nothing
    Set oPart = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub




