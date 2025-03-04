VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_PS2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MI_PS2_1.cls
'   ProgID:         HS_Hilti_MIParts.MI_PS2_1
'   Author:         JRK
'   Creation Date:  Friday, September 28, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP          Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MI_PS2_1"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_PS2_1"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Dim sPartNumber As String

Implements IJHgrSymbolWCGServices
Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices

Private Sub IJHgrSymbolWCGServices_EvaluateLocalWCG(ByVal pSupportComp As Object, Weight As Double, CogX As Double, CogY As Double, CogZ As Double)
    Const METHOD = " IJHgrSymbolWCGServices_EvaluateLocalWCG"
    On Error GoTo ErrorHandler

    Dim dWeight As Double
    Dim lShoeH As Double

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp
    oPartOcc.GetPart oPart

    lShoeH = GetAttributeFromObject(oPart, "ShoeH")
    If lShoeH = 1 Then
        dWeight = GetAttributeFromObject(oPart, "Weight85")
    ElseIf lShoeH = 2 Then
        dWeight = GetAttributeFromObject(oPart, "Weight140")
    ElseIf lShoeH = 3 Then
        dWeight = GetAttributeFromObject(oPart, "Weight84")
    ElseIf lShoeH = 4 Then
        dWeight = GetAttributeFromObject(oPart, "Weight107")
    ElseIf lShoeH = 5 Then
        dWeight = GetAttributeFromObject(oPart, "Weight142")
    End If

    Weight = dWeight

    CogZ = 0
    CogX = 0
    CogY = 0

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

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
    AddInput Inputs, "ShoeH", "Long"
    AddInput Inputs, "Connex", "Long"
    AddInput Inputs, "Girder", "Long"
    AddInput Inputs, "FillerStrip", "Long"
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "B", "Double"
    AddInput Inputs, "S", "Double"
    AddInput Inputs, "F", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "H0", "Double"
    AddInput Inputs, "H1", "Double"
    AddInput Inputs, "L", "Double"
    AddInput Inputs, "L1_85", "Double"
    AddInput Inputs, "L1_140", "Double"
    AddInput Inputs, "ClScrew", "String"
    AddInput Inputs, "BoltDia", "Double"
    AddInput Inputs, "Weight85", "Double"
    AddInput Inputs, "Weight140", "Double"
    AddInput Inputs, "ItemID", "String"
    AddInput Inputs, "ItemNo85", "String"
    AddInput Inputs, "ItemNo140", "String"
    AddInput Inputs, "OrderNo", "String"
    AddInput Inputs, "OrderNo45", "String"
    AddInput Inputs, "ClampRot", "Long"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Girder"
    AddOutput Outputs, "Clamp1Body"
    AddOutput Outputs, "Clamp2Body"
    AddOutput Outputs, "Clamp1Top1"
    AddOutput Outputs, "Clamp1Top2"
    AddOutput Outputs, "Clamp1Bot1"
    AddOutput Outputs, "Clamp1Bot2"
    AddOutput Outputs, "Clamp1TOP_BOLT"
    AddOutput Outputs, "Clamp1BOT_BOLT"
    AddOutput Outputs, "Clamp2Top1"
    AddOutput Outputs, "Clamp2Top2"
    AddOutput Outputs, "Clamp2Bot1"
    AddOutput Outputs, "Clamp2Bot2"
    AddOutput Outputs, "Clamp2TOP_BOLT"
    AddOutput Outputs, "Clamp2BOT_BOLT"
    AddOutput Outputs, "VertPlate"
    AddOutput Outputs, "HortPlate"
    AddOutput Outputs, "Optional1"
    AddOutput Outputs, "Optional2"
    AddOutput Outputs, "Optional3"
    AddOutput Outputs, "Optional4"
    AddOutput Outputs, "Optional5"
    AddOutput Outputs, "Optional6"
    AddOutput Outputs, "Optional7"
    AddOutput Outputs, "Optional8"
    AddOutput Outputs, "Optional9"
    AddOutput Outputs, "Optional10"
    AddOutput Outputs, "Optional11"
    AddOutput Outputs, "Optional12"
    AddOutput Outputs, "Optional13"
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

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH
    Set oSupportComp = Nothing

    Dim lShoeH As Long
    Dim dShoeH As Double
    Dim lConnex As Long
    Dim sConnex As String
    Dim lGirder As Long
    Dim lFillerStrip As Long
    Dim dD As Double
    Dim dB As Double
    Dim dS As Double
    Dim dF As Double
    Dim dW As Double
    Dim dH0 As Double
    Dim dH1 As Double
    Dim dL As Double
    Dim dL1 As Double
    Dim dL1_85 As Double
    Dim dL1_140 As Double
    Dim sClScrew As String
    Dim dBoltDia As Double
    Dim dWeight85 As Double
    Dim dWeight140 As Double
    Dim sItemID As String
    Dim sItemNo85 As String
    Dim sItemNo140 As String
    Dim sOrderNo As String
    Dim sOrderNo45 As String
    Dim dInsulat As Double
    Dim dPipeDia As Double
    Dim dSpacer As Double
    Dim lClampRot As Long
    Dim sAngle As String
    Dim dSPlateTh As Double
    Dim dSPlateL As Double
    Dim dSPlateW As Double

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dInsulat = 0.0254
    Else
        Hilti_InitializeMyHH oICH
        GetPipeDiameterAndInsulatTh 1, dPipeDia, dInsulat
        Hilti_DestroyMyHH oICH
        Set oICH = Nothing
    End If

    lShoeH = arrayOfInputs(2)
    lConnex = arrayOfInputs(3)
    lGirder = arrayOfInputs(4)
    lFillerStrip = arrayOfInputs(5)
    dD = arrayOfInputs(6)
    dB = arrayOfInputs(7)
    dS = arrayOfInputs(8)
    dF = arrayOfInputs(9)
    dW = arrayOfInputs(10)
    dH0 = arrayOfInputs(11)
    dH1 = arrayOfInputs(12)
    dL = arrayOfInputs(13)
    dL1_85 = arrayOfInputs(14)
    dL1_140 = arrayOfInputs(15)
    sClScrew = arrayOfInputs(16)
    dBoltDia = arrayOfInputs(17)
    dWeight85 = arrayOfInputs(18)
    dWeight140 = arrayOfInputs(19)
    sItemID = arrayOfInputs(20)
    sItemNo85 = arrayOfInputs(21)
    sItemNo140 = arrayOfInputs(22)
    sOrderNo = arrayOfInputs(23)
    sOrderNo45 = arrayOfInputs(24)
    lClampRot = arrayOfInputs(25)
    dSpacer = 0.002
    dPipeDia = dD
    dSPlateTh = 0

    If Not oCodeListData Is Nothing Then
        sConnex = Trim(oCodeListData.ShortStringValue("MICode_Connection", lConnex))
    Else
        sConnex = "Sliding"
    End If
    
    dShoeH = (oCodeListData.LongStringValue("MICode_ShoeH", lShoeH)) / 1000
    
    If lShoeH = 1 Or lShoeH = 3 Or lShoeH = 4 Then
        dL1 = dL1_85
    ElseIf lShoeH = 2 Or lShoeH = 5 Then
        dL1 = dL1_140
    End If
    dH1 = dShoeH - dS
    dH0 = dShoeH

    If lClampRot = 1 Then
        sAngle = "0"
    Else
        sAngle = "45"
    End If

    AddPort "Route", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"
    AddPort "Girder", 0, 0, -dPipeDia / 2 - dShoeH, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Girder", arrayOfInputs(1), "Girder"

    'Clamp cylinders
    AddCylinder dB, dPipeDia / 2 + dS, "ROTX(" + Str(-90) + ") * " + Loc(0, -dL / 2, 0) + " * ROTZ(270)", m_outputColl, "Clamp1Body", "Clamp1Body"
    AddCylinder dB, dPipeDia / 2 + dS, "ROTX(" + Str(-90) + ") * " + Loc(0, dL / 2 - dB, 0) + " * ROTZ(270)", m_outputColl, "Clamp2Body", "Clamp2Body"
    'Add clamp boxes
    AddBox dB, dB, dS, Loc(-dL / 2, dPipeDia / 2, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1Top1", "Clamp1Top1"
    AddBox dB, dB, dS, Loc(-dL / 2, -dPipeDia / 2 - dB, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1Top2", "Clamp1Top2"
    AddBox dB, dB, dS, Loc(-dL / 2, dPipeDia / 2, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1Bot1", "Clamp1Bot1"
    AddBox dB, dB, dS, Loc(-dL / 2, -dPipeDia / 2 - dB, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1Bot2", "Clamp1Bot2"
    'Add clamp bolts
    AddCylinder dSpacer + dS * 2 + dBoltDia * 2, dBoltDia / 2, Loc(-dL / 2 + dB / 2, -dPipeDia / 2 - dB / 2, -dSpacer / 2 - dS - dBoltDia) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1TOP_BOLT", "Clamp1TOP_BOLT"
    AddCylinder dSpacer + dS * 2 + dBoltDia * 2, dBoltDia / 2, Loc(-dL / 2 + dB / 2, dPipeDia / 2 + dB / 2, -dSpacer / 2 - dS - dBoltDia) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp1BOT_BOLT", "Clamp1BOT_BOLT"
    'Add clamp boxes
    AddBox dB, dB, dS, Loc(dL / 2 - dB, dPipeDia / 2, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2Top1", "Clamp2Top1"
    AddBox dB, dB, dS, Loc(dL / 2 - dB, -dPipeDia / 2 - dB, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2Top2", "Clamp2Top2"
    AddBox dB, dB, dS, Loc(dL / 2 - dB, dPipeDia / 2, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2Bot1", "Clamp2Bot1"
    AddBox dB, dB, dS, Loc(dL / 2 - dB, -dPipeDia / 2 - dB, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2Bot2", "Clamp2Bot2"
    'Add clamp bolts
    AddCylinder dSpacer + dS * 2 + dBoltDia * 2, dBoltDia / 2, Loc(dL / 2 - dB / 2, -dPipeDia / 2 - dB / 2, -dSpacer / 2 - dS - dBoltDia) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2TOP_BOLT", "Clamp2TOP_BOLT"
    AddCylinder dSpacer + dS * 2 + dBoltDia * 2, dBoltDia / 2, Loc(dL / 2 - dB / 2, dPipeDia / 2 + dB / 2, -dSpacer / 2 - dS - dBoltDia) + " *ROTX(" + sAngle + ")", m_outputColl, "Clamp2BOT_BOLT", "Clamp2BOT_BOLT"
    'Add shoe boxes
    AddBox dL, dF, dH1 - dF, Loc(-dL / 2, -dF / 2, -dPipeDia / 2 - dS - dH1 + dF), m_outputColl, "VertPlate", "VertPlate"
    AddBox dL, dW, dF, Loc(-dL / 2, -dW / 2, -dPipeDia / 2 - dS - dH1), m_outputColl, "HortPlate", "HortPlate"

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
    Dim dGirder As Double

    If lGirder = 1 Then
        dP_W = 0.045
        dP_W1 = 0.03
        dP_T = 0.006
        dP_H = 0.018
        dP_D = 0.13
        dB_DIA = 0.012
        dB_CC = 0.102
        dB_EXTRA = 0.005
        dB_2H = 0.126
        dGirder = 0.09
    Else
        dP_W = 0.045
        dP_W1 = 0.03
        dP_T = 0.006
        dP_H = 0.018
        dP_D = 0.13
        dB_DIA = 0.012
        dB_CC = 0.102
        dB_EXTRA = 0.005
        dB_2H = 0.156
        dGirder = 0.12
    End If

    dB_H = dB_DIA + dGirder + dP_H + dB_EXTRA
    dB_H2 = dB_CC + dB_DIA

    If sConnex = "Guided w/ Low Friction Plate" Or sConnex = "Sliding w/ Low Friction Plate" Then
        dSPlateTh = 0.005
        dSPlateL = 0.075
        dSPlateW = dGirder - 0.02
    End If

    If sConnex = "Guided" Or sConnex = "Guided w/ Low Friction Plate" Then
        AddBox dP_D, dP_W, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T), m_outputColl, "Optional1", "Optional1"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T * 2), m_outputColl, "Optional2", "Optional2"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T * 3), m_outputColl, "Optional3", "Optional3"
        AddCylinder dB_H, dB_DIA / 2, "ROTZ(270)*" + Loc(-dB_CC / 2, -dW / 2 - dP_W1 / 2, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional4", "Optional4"
        AddCylinder dB_H, dB_DIA / 2, "ROTZ(270)*" + Loc(dB_CC / 2, -dW / 2 - dP_W1 / 2, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional5", "Optional5"
        AddCylinder dB_H2, dB_DIA / 2, "ROTY(270)*" + Loc(dB_CC / 2 + dB_DIA / 2, -dW / 2 - dP_W1 / 2, -dPipeDia / 2 - dH0 - dGirder - dB_DIA / 2), m_outputColl, "Optional6", "Optional6"

        AddBox dP_D, dP_W, dP_T, Loc(-dP_D / 2#, dW / 2# - (dP_W - dP_W1) + 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T), m_outputColl, "Optional7", "Optional7"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, dW / 2 + 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T * 2), m_outputColl, "Optional8", "Optional8"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, dW / 2 + 0.002, -dPipeDia / 2# - dH0 + dP_H - dP_T * 3), m_outputColl, "Optional9", "Optional9"
        AddCylinder dB_H, dB_DIA / 2, "ROTZ(270)*" + Loc(-dB_CC / 2, dW / 2# + dP_W1 / 2#, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional10", "Optional10"
        AddCylinder dB_H, dB_DIA / 2, "ROTZ(270)*" + Loc(dB_CC / 2, dW / 2# + dP_W1 / 2#, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional11", "Optional11"
        AddCylinder dB_H2, dB_DIA / 2, "ROTY(270)*" + Loc(dB_CC / 2 + dB_DIA / 2, dW / 2# + dP_W1 / 2#, -dPipeDia / 2 - dH0 - dGirder - dB_DIA / 2), m_outputColl, "Optional12", "Optional12"
    End If

    If sConnex = "Fixed" Then
        AddCylinder dB_2H, dB_DIA / 2, "ROTZ(270)*" + Loc(-dB_CC / 2, -dW / 2 + dP_W1 / 2, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional1", "Optional1"
        AddCylinder dB_2H, dB_DIA / 2, "ROTZ(270)*" + Loc(dB_CC / 2, -dW / 2 + dP_W1 / 2, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional2", "Optional2"
        AddCylinder dB_H2, dB_DIA / 2, "ROTY(270)*" + Loc(dB_CC / 2 + dB_DIA / 2, -dW / 2 + dP_W1 / 2, -dPipeDia / 2 - dH0 - dGirder - dB_DIA / 2), m_outputColl, "Optional3", "Optional3"
        AddCylinder dB_2H, dB_DIA / 2, "ROTZ(270)*" + Loc(-dB_CC / 2, dW / 2# - dP_W1 / 2#, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional4", "Optional4"
        AddCylinder dB_2H, dB_DIA / 2, "ROTZ(270)*" + Loc(dB_CC / 2, dW / 2# - dP_W1 / 2#, -dPipeDia / 2 - dH0 + dP_H + dB_EXTRA - dB_H), m_outputColl, "Optional5", "Optional5"
        AddCylinder dB_H2, dB_DIA / 2, "ROTY(270)*" + Loc(dB_CC / 2 + dB_DIA / 2, dW / 2# - dP_W1 / 2#, -dPipeDia / 2 - dH0 - dGirder - dB_DIA / 2), m_outputColl, "Optional6", "Optional6"
    Else
        If dSPlateTh <> 0 Then
            AddBox dSPlateW, dSPlateL, dSPlateTh, Loc(-dSPlateW / 2#, -dSPlateL / 2, -dPipeDia / 2# - dH0 - dSPlateTh), m_outputColl, "Optional13", "Optional13"
        End If
    End If

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
    oPartOcc.GetPart oPart      ' The associated catalog part

    Dim sConnectorBom As String
    Dim sConnector2Bom As String
    Dim lGirder As Long
    Dim lConnex As Long
    Dim lClampRot As Long
    Dim sConnex As String
    Dim lShoeH As Long
    Dim sShoeH As String
    Dim sItemNo As String
    Dim lFillerStrip As Long
    Dim sOrderNo As String
    Dim dFillerStripLength As Double
    Dim dPi As Double
    Dim dPipeDia As Double
    Dim sFillerStripDisp As String
    Dim dWeight As Double

    lShoeH = GetAttributeFromObject(oPart, "ShoeH")
    lGirder = GetAttributeFromObject(pSupportComp, "Girder")
    lConnex = GetAttributeFromObject(pSupportComp, "Connex")
    lFillerStrip = GetAttributeFromObject(pSupportComp, "FillerStrip")
    lClampRot = GetAttributeFromObject(pSupportComp, "ClampRot")
    dPi = 3.14159265358979

    If lShoeH = 1 Then
        sShoeH = "85"
    ElseIf lShoeH = 2 Then
        sShoeH = "140"
    ElseIf lShoeH = 3 Then
        sShoeH = "84"
    ElseIf lShoeH = 4 Then
        sShoeH = "107"
    Else
        sShoeH = "142"
    End If

    sItemNo = GetAttributeFromObject(oPart, "ItemNo" & sShoeH)
    dWeight = GetAttributeFromObject(oPart, "Weight" & sShoeH)
    dPipeDia = GetAttributeFromObject(oPart, "D")


    If lClampRot = 1 Then
        sOrderNo = GetAttributeFromObject(oPart, "OrderNo")
    Else
        sOrderNo = GetAttributeFromObject(oPart, "OrderNo45")
    End If

    If lFillerStrip = 1 Then
        dFillerStripLength = dPi * 2 * (dPipeDia / 2)
        sFillerStripDisp = ConvertValueToLongStringValue(pSupportComp, "IJUAMI_D", "D", dFillerStripLength)
    End If

    If lGirder = 1 Then
        sConnectorBom = " + 1, 304838, Connector Pipe Shoe MIC-PS90"
        sConnector2Bom = " + 1, 304840, Guide, MIA-B090-M12"
    Else
        sConnectorBom = " + 1, 304839, Guide MIC-PS120"
        sConnector2Bom = " + 1, 304841, Guide MIA-B0120-M12"
    End If

    If lConnex = 1 Then    '"Sliding" Then
        If lFillerStrip = 1 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + Filler Strip " & sFillerStripDisp
        Else
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "")
        End If
    End If

    If lConnex = 2 Then    '"Sliding w/ Low Friction Plate" Then
        If lFillerStrip = 1 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + Filler Strip " & sFillerStripDisp & " + 1, 304842, Slide Plate MIC-PG"
        Else
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + 1, 304842, Slide Plate MIC-PG"
        End If
    End If

    If lConnex = 3 Then    '"Guided" Then
        If lFillerStrip = 1 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + Filler Strip " & sFillerStripDisp & sConnectorBom
        Else
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBom
        End If
    End If

    If lConnex = 4 Then    '"Guided w/ Low Friction Plate" Then
        If lFillerStrip = 1 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + Filler Strip " & sFillerStripDisp & sConnectorBom & " + 1, 304842, Slide Plate MIC-PG"
        Else
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnectorBom & " + 1, 304842, Slide Plate MIC-PG"
        End If
    End If

    If lConnex = 5 Then    '"Fixed" Then
        If lFillerStrip = 1 Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " + Filler Strip " & sFillerStripDisp & sConnector2Bom
        Else
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & sConnector2Bom
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub


