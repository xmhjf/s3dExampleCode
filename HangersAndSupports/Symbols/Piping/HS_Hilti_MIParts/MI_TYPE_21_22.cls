VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_TYPE_21_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MI_TYPE_21_22.cls
'   ProgID:         HS_Hilti_MIParts.MI_TYPE_21_22
'   Author:         JRK
'   Creation Date:  Friday, November 2, 2007

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
Private Const MODULE = "MI_TYPE_21_22"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_TYPE_21_22"

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

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp
    oPartOcc.GetPart oPart
    Dim sBearingType As String
    Dim sPartNumber As String
    sPartNumber = oPart.PartNumber
    Dim sType As String
    sType = Mid(sPartNumber, 5, 2)
    dWeight = MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASMI_TYPE_" & Trim(sType), "PartNumber", "JDPart", "'" & sPartNumber & "'")
    sBearingType = Trim(GetAttributeFromObject(oPart, "BearingType"))

    If sBearingType = "LVL" Then
        dWeight = dWeight - 1.7
    End If
    If sBearingType = "FVN" Then
        dWeight = dWeight + 1#
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
    AddInput Inputs, "Length", "Double"
    AddInput Inputs, "Connex", "Long"
    AddInput Inputs, "Girder", "Long"
    AddInput Inputs, "HDefault", "Long"
    AddInput Inputs, "BC", "Double"
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "B", "Double"
    AddInput Inputs, "C", "Double"
    AddInput Inputs, "S", "Double"
    AddInput Inputs, "F", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "H0", "Double"
    AddInput Inputs, "L", "Double"
    AddInput Inputs, "LFVN", "Double"
    AddInput Inputs, "L1", "Double"
    AddInput Inputs, "Clamp", "String"
    AddInput Inputs, "ItemID", "String"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "OrderNo", "String"
    AddInput Inputs, "BearingType", "String"
    AddInput Inputs, "SlidingPlate", "Long"
    AddInput Inputs, "ClampRot", "Long"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Girder"
    AddOutput Outputs, "PipeClampBody1"
    AddOutput Outputs, "RightTop1"
    AddOutput Outputs, "RightBottom1"
    AddOutput Outputs, "RightBolt1"
    AddOutput Outputs, "LeftTop1"
    AddOutput Outputs, "LeftBottom1"
    AddOutput Outputs, "LeftBolt1"
    AddOutput Outputs, "Shoe"
    AddOutput Outputs, "UprightLeftClampConn"
    AddOutput Outputs, "UprightRightShoeConn"
    AddOutput Outputs, "FrontUpRightBolt"
    AddOutput Outputs, "BackUpRightBolt"
    AddOutput Outputs, "LanglochRightBack"
    AddOutput Outputs, "LanglochLeftBack"
    AddOutput Outputs, "LanglochRightFront"
    AddOutput Outputs, "LanglochLeftFront"
    AddOutput Outputs, "RightFrontShoeBolt"
    AddOutput Outputs, "LeftFrontShoeBolt"
    AddOutput Outputs, "RightBackShoeBolt"
    AddOutput Outputs, "LeftBackShoeBolt"
    AddOutput Outputs, "RightFrontFixBox"
    AddOutput Outputs, "RightBackFixBox"
    AddOutput Outputs, "RightFixBolt"
    AddOutput Outputs, "LeftFrontFixBox"
    AddOutput Outputs, "LeftBackFixBox"
    AddOutput Outputs, "LeftFixBolt"
    AddOutput Outputs, "ClampPart1"
    AddOutput Outputs, "ClampPart2"
    AddOutput Outputs, "ClampPart3"
    AddOutput Outputs, "ClampPart4"
    AddOutput Outputs, "ClampPart5"
    AddOutput Outputs, "ClampPart6"
    AddOutput Outputs, "ClampPart7"
    AddOutput Outputs, "ClampPart8"
    AddOutput Outputs, "ClampPart9"
    AddOutput Outputs, "ClampPart10"
    AddOutput Outputs, "ClampPart11"
    AddOutput Outputs, "ClampPart12"
    AddOutput Outputs, "PipeClampBody2"
    AddOutput Outputs, "RightTop2"
    AddOutput Outputs, "RightBottom2"
    AddOutput Outputs, "RightBolt2"
    AddOutput Outputs, "LeftTop2"
    AddOutput Outputs, "LeftBottom2"
    AddOutput Outputs, "LeftBolt2"
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

    Dim dLength As Double
    Dim lConnex As Long
    Dim sConnex As String
    Dim lGirder As Long
    Dim lHDefault As Long
    Dim dBC As Double
    Dim dD As Double
    Dim dB As Double
    Dim dC As Double
    Dim dS As Double
    Dim dF As Double
    Dim dW As Double
    Dim dH As Double
    Dim dL As Double
    Dim dLFVN As Double
    Dim dL1 As Double
    Dim sClamp As String
    Dim dHDefault As Double
    Dim sBearingType As String
    Dim dH1 As Double
    Dim dBoltDia As Double
    Dim dSpacer As Double
    Dim dBoltH As Double
    Dim dPipeDia As Double
    Dim lSlidingPlate As Long
    Dim lClampRot As Long
    Dim dInsulat As Double
    Dim dKorrKreis As Double
    Dim dKorrBox As Double

    dLength = arrayOfInputs(2)
    lConnex = arrayOfInputs(3)
    lGirder = arrayOfInputs(4)
    lHDefault = arrayOfInputs(5)
    dBC = arrayOfInputs(6)
    dD = arrayOfInputs(7)
    dB = arrayOfInputs(8)
    dC = arrayOfInputs(9)
    dS = arrayOfInputs(10)
    dF = arrayOfInputs(11)
    dW = arrayOfInputs(12)
    dH = arrayOfInputs(13)
    dL = arrayOfInputs(14)
    dLFVN = arrayOfInputs(15)
    dL1 = arrayOfInputs(16)
    sClamp = arrayOfInputs(17)
    sBearingType = Trim(arrayOfInputs(21))
    lSlidingPlate = arrayOfInputs(22)
    lClampRot = arrayOfInputs(23)
    dPipeDia = dD

    If (sBearingType = "LVL" Or sBearingType = "LVG") And lSlidingPlate = 2 Then
        dF = dF / 2#
    End If

    If sBearingType = "FVN" Then
        dL = dLFVN
        dF = dF / 2#
    End If


    If lHDefault = 1 Then      'take calculated H based on LENGTH
        dH = dLength - dD / 2#
    Else
        dHDefault = Val(ConvertValueToLongStringValue(sblOcc, "IJOAMI_HDefault", "HDefault", lHDefault)) / 1000
        dH = dHDefault
    End If

    If sBearingType = "LVL" And lConnex = 1 Then
        PF_EventHandler "MI-Girder connection not supported for " & sBearingType, Err, MODULE, METHOD, True
    End If

    If dH > 0.215 Then
        PF_EventHandler "Shoe Height must be between 90 and 215 mm", Err, MODULE, METHOD, True
    End If

    dH1 = dH - dS

    AddPort "Route", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Route", arrayOfInputs(1), "Route"
    AddPort "Girder", 0, 0, dH + dD / 2, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Girder", arrayOfInputs(1), "Girder"

    dBoltDia = 0.01
    dBoltDia = Val(Right(Trim(sClamp), 2)) / 1000
    dSpacer = 0.3 * dS
    dBoltH = dSpacer + 2 * dS + 1.2 * dBoltDia
    Dim sAngle As String
    If lClampRot = 1 Then
        sAngle = "0"
    Else
        sAngle = "45"
    End If

    'Clamp Body
    AddCylinder dC, dPipeDia / 2 + dS, "ROTX(" + Str(-90) + ") * " + Loc(0, -dL / 2, 0) + " * ROTZ(270)", m_outputColl, "PipeClampBody1", "PipeClampBody1"
    AddCylinder dC, dPipeDia / 2 + dS, "ROTX(" + Str(-90) + ") * " + Loc(0, dL / 2 - dC, 0) + " * ROTZ(270)", m_outputColl, "PipeClampBody2", "PipeClampBody2"
    'Add clamp boxes
    AddBox dC, dBC, dS, Loc(-dL / 2, dPipeDia / 2, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "RightTop1", "RightTop1"
    AddBox dC, dBC, dS, Loc(-dL / 2, -dPipeDia / 2 - dBC, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftTop1", "LeftTop1"
    AddBox dC, dBC, dS, Loc(-dL / 2, dPipeDia / 2, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "RightBottom1", "RightBottom1"
    AddBox dC, dBC, dS, Loc(-dL / 2, -dPipeDia / 2 - dBC, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftBottom1", "LeftBottom1"
    AddBox dC, dBC, dS, Loc(dL / 2 - dC, dPipeDia / 2, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "RightTop2", "RightTop2"
    AddBox dC, dBC, dS, Loc(dL / 2 - dC, -dPipeDia / 2 - dBC, dSpacer) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftTop2", "LeftTop2"
    AddBox dC, dBC, dS, Loc(dL / 2 - dC, dPipeDia / 2, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "RightBottom2", "RightBottom2"
    AddBox dC, dBC, dS, Loc(dL / 2 - dC, -dPipeDia / 2 - dBC, -dSpacer - dS) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftBottom2", "LeftBottom2"
    'Add clamp bolts
    AddCylinder dBoltH, dBoltDia / 2, Loc(-dL / 2 + dC / 2, -dPipeDia / 2 - dBC / 2, -dBoltH / 2) + " *ROTX(" + sAngle + ")", m_outputColl, "RightBolt1", "RightBolt1"
    AddCylinder dBoltH, dBoltDia / 2, Loc(-dL / 2 + dC / 2, dPipeDia / 2 + dBC / 2, -dBoltH / 2) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftBolt1", "LeftBolt1"
    AddCylinder dBoltH, dBoltDia / 2, Loc(dL / 2 - dC / 2, -dPipeDia / 2 - dBC / 2, -dBoltH / 2) + " *ROTX(" + sAngle + ")", m_outputColl, "RightBolt2", "RightBolt2"
    AddCylinder dBoltH, dBoltDia / 2, Loc(dL / 2 - dC / 2, dPipeDia / 2 + dBC / 2, -dBoltH / 2) + " *ROTX(" + sAngle + ")", m_outputColl, "LeftBolt2", "LeftBolt2"

    'Add shoe boxes
    AddBox dL, dW, dF, Loc(-dL / 2, -dW / 2, dPipeDia / 2 + dS + dH1 - dF), m_outputColl, "Shoe", "Shoe"
    AddBox dL, dS, dH1 - dF - (dH1 / 7), Loc(-dL / 2, 0, dPipeDia / 2 + dS), m_outputColl, "UprightLeftClampConn", "UprightLeftClampConn"
    AddBox dL, dS, dH1 - dF - (dH1 / 7), Loc(-dL / 2, -dS, dPipeDia / 2 + dS + dH1 / 7), m_outputColl, "UprightRightShoeConn", "UprightRightShoeConn"

    'Adjustment bolts of upright boxes
    Dim dBoltZD As Double
    Dim dBoltZH As Double
    dBoltZD = 0.01
    dBoltZH = 4 * dS
    AddCylinder dBoltZH, dBoltZD / 2, "ROTX(270)*" & Loc(-dL / 4, -dBoltZH / 2, dD / 2 + dH1 / 2), m_outputColl, "FrontUpRightBolt", "FrontUpRightBolt"
    AddCylinder dBoltZH, dBoltZD / 2, "ROTX(270)*" & Loc(dL / 4, -dBoltZH / 2, dD / 2 + dH1 / 2), m_outputColl, "BackUpRightBolt", "BackUpRightBolt"

    'Adjustment holes within upright boxes
    AddBox dBoltZD * 4, dS, dBoltZD, "ROTY(30)*" + Loc(-0.01982 - dL / 4, -dS, 0.00567 + dD / 2 + dH1 / 2), m_outputColl, "LanglochRightBack", "LanglochRightBack"
    AddBox dBoltZD * 4, dS, dBoltZD, "ROTY(-30)*" + Loc(-0.01482 + dL / 4, -dS, -0.01433 + dD / 2 + dH1 / 2), m_outputColl, "LanglochLeftBack", "LanglochLeftBack"
    AddBox dBoltZD * 4, dS, dBoltZD, "ROTY(-30)*" + Loc(-0.01482 - dL / 4, 0, -0.01433 + dD / 2 + dH1 / 2), m_outputColl, "LanglochRightFront", "LanglochRightFront"
    AddBox dBoltZD * 4, dS, dBoltZD, "ROTY(30)*" + Loc(-0.01982 + dL / 4, 0, 0.00567 + dD / 2 + dH1 / 2), m_outputColl, "LanglochLeftFront", "LanglochLeftFront"

    ' The four shoe bolts for bearing type lvl and lvg
    If (sBearingType = "LVL" Or sBearingType = "LVG") And lSlidingPlate = 1 Then
        AddCylinder 1.5 * dF, 0.003, Loc(0.01 - dL / 2, -0.02 + dW / 2, dD / 2 + dS + dH1 - 1.25 * dF), m_outputColl, "RightFrontShoeBolt", "RightFrontShoeBolt"
        AddCylinder 1.5 * dF, 0.003, Loc(0.01 - dL / 2, 0.02 - dW / 2, dD / 2 + dS + dH1 - 1.25 * dF), m_outputColl, "LeftFrontShoeBolt", "LeftFrontShoeBolt"
        AddCylinder 1.5 * dF, 0.003, Loc(-0.01 + dL / 2, -0.02 + dW / 2, dD / 2 + dS + dH1 - 1.25 * dF), m_outputColl, "RightBackShoeBolt", "RightBackShoeBolt"
        AddCylinder 1.5 * dF, 0.003, Loc(-0.01 + dL / 2, 0.02 - dW / 2, dD / 2 + dS + dH1 - 1.25 * dF), m_outputColl, "LeftBackShoeBolt", "LeftBackShoeBolt"
    End If

    'Shoe fix boxes for LVG and FVN and fixbolts
    Dim dXFix As Double
    Dim dYFix As Double
    Dim dZFix As Double
    If (sBearingType = "LVG" Or sBearingType = "FVN") And lConnex = 2 Then
        dXFix = 2# * dF
        dYFix = dL / 5#
        dZFix = 2.5 * dF

        AddBox dYFix, dXFix, dZFix, Loc(dYFix / 2# - dL / 2#, -0.25 * dXFix + dW / 2#, dD / 2# + dS + dH1 - 1.5 * dF), m_outputColl, "RightFrontFixBox", "RightFrontFixBox"
        AddBox dYFix, dXFix, dZFix, Loc(-1.5 * dYFix + dL / 2#, -0.25 * dXFix + dW / 2#, dD / 2# + dS + dH1 - 1.5 * dF), m_outputColl, "RightBackFixBox", "RightBackFixBox"
        AddCylinder dL - dYFix / 2#, 0.003, "ROTY(270)*" + Loc(dL / 2# - dYFix / 4#, 0.25 * dXFix + dW / 2#, dD / 2# + dS + dH1 - dF), m_outputColl, "RightFixBolt", "RightFixBolt"

        AddBox dYFix, dXFix, dZFix, Loc(dYFix / 2# - dL / 2#, -0.75 * dXFix - dW / 2#, dD / 2# + dS + dH1 - 1.5 * dF), m_outputColl, "LeftFrontFixBox", "LeftFrontFixBox"
        AddBox dYFix, dXFix, dZFix, Loc(-1.5 * dYFix + dL / 2#, -0.75 * dXFix - dW / 2#, dD / 2# + dS + dH1 - 1.5 * dF), m_outputColl, "LeftBackFixBox", "LeftBackFixBox"
        AddCylinder dL - dYFix / 2#, 0.003, "ROTY(270)*" + Loc(dL / 2# - dYFix / 4#, -0.25 * dXFix - dW / 2#, dD / 2# + dS + dH1 - dF), m_outputColl, "LeftFixBolt", "LeftFixBolt"
    End If

    If lConnex = 1 And (sBearingType = "LVG" Or sBearingType = "FVN") Then

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

        dP_W = 0.045
        dP_W1 = 0.03
        dP_T = 0.006
        dP_H = 0.018
        dP_D = 0.13
        dB_DIA = 0.012
        dB_CC = 0.102
        dB_EXTRA = 0.005

        If lGirder = 1 Then
            dB_2H = 0.126
            dGirder = 0.09
        Else
            dB_2H = 0.156
            dGirder = 0.12
        End If

        dB_H = dB_DIA + dGirder + dP_H + dB_EXTRA
        dB_H2 = dB_CC + dB_DIA

        AddBox dP_D, dP_W, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, dD / 2# + dH - dP_H), m_outputColl, "ClampPart1", "ClampPart1"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, dD / 2# + dH - dP_H + dP_T), m_outputColl, "ClampPart2", "ClampPart2"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, -dW / 2# - dP_W1 - 0.002, dD / 2# + dH - dP_H + 2 * dP_T), m_outputColl, "ClampPart3", "ClampPart3"

        AddCylinder dB_H, dB_DIA / 2#, "ROTZ(90)*" + Loc(-dB_CC / 2#, -dW / 2# - dP_W1 / 2#, dD / 2# + dH - dP_H - dB_EXTRA), m_outputColl, "ClampPart4", "ClampPart4"
        AddCylinder dB_H, dB_DIA / 2#, "ROTZ(90)*" + Loc(dB_CC / 2#, -dW / 2# - dP_W1 / 2#, dD / 2# + dH - dP_H - dB_EXTRA), m_outputColl, "ClampPart5", "ClampPart5"
        AddCylinder dB_H2, dB_DIA / 2#, "ROTY(270)*" + Loc(dB_CC / 2# + dB_DIA / 2#, -dW / 2# - dP_W1 / 2#, dD / 2# + dH + dGirder + dB_DIA / 2#), m_outputColl, "ClampPart6", "ClampPart6"

        AddBox dP_D, dP_W, dP_T, Loc(-dP_D / 2#, dW / 2# - (dP_W - dP_W1) + 0.002, dD / 2# + dH - dP_H), m_outputColl, "ClampPart7", "ClampPart7"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, dW / 2# + 0.002, dD / 2# + dH - dP_H + dP_T), m_outputColl, "ClampPart8", "ClampPart8"
        AddBox dP_D, dP_W1, dP_T, Loc(-dP_D / 2#, dW / 2# + 0.002, dD / 2# + dH - dP_H + 2 * dP_T), m_outputColl, "ClampPart9", "ClampPart9"

        AddCylinder dB_H, dB_DIA / 2#, "ROTZ(90)*" + Loc(-dB_CC / 2#, dW / 2# + dP_W1 / 2#, dD / 2# + dH - dP_H - dB_EXTRA), m_outputColl, "ClampPart10", "ClampPart10"
        AddCylinder dB_H, dB_DIA / 2#, "ROTZ(90)*" + Loc(dB_CC / 2#, dW / 2# + dP_W1 / 2#, dD / 2# + dH - dP_H - dB_EXTRA), m_outputColl, "ClampPart11", "ClampPart11"
        AddCylinder dB_H2, dB_DIA / 2#, "ROTY(270)*" + Loc(dB_CC / 2# + dB_DIA / 2#, dW / 2# + dP_W1 / 2#, dD / 2# + dH + dGirder + dB_DIA / 2#), m_outputColl, "ClampPart12", "ClampPart12"

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
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim lGirder As Long
    Dim lConnex As Long
    Dim sOrderNo As String
    Dim sHDefault As String
    Dim dH As Double
    Dim sPipeNomDia As String
    Dim sBearingType As String
    Dim sPartNumber As String
    Dim dWeight As Double
    Dim sWeight As String
    Dim sConnectorBom As String
    Dim dD As Double

    sBearingType = GetAttributeFromObject(oPart, "BearingType")
    sOrderNo = GetAttributeFromObject(oPart, "OrderNo")
    dH = GetAttributeFromObject(oPart, "H0")
    dD = GetAttributeFromObject(oPart, "D")
    lGirder = GetAttributeFromObject(pSupportComp, "Girder")
    lConnex = GetAttributeFromObject(pSupportComp, "Connex")

    If dH < 0.12 Then
        sHDefault = "100"
    End If
    If dH >= 0.12 And dH < 0.17 Then
        sHDefault = "150"
    End If
    If dH >= 0.17 Then
        sHDefault = "200"
    End If

    sPartNumber = oPart.PartNumber
    If dD < 0.1 Then
        sPipeNomDia = Trim(Mid(sPartNumber, 12, 2))
    Else
        sPipeNomDia = Trim(Mid(sPartNumber, 12, 3))
    End If

    Dim sType As String
    sType = Mid(sPartNumber, 5, 2)
    dWeight = MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASMI_TYPE_" & Trim(sType), "PartNumber", "JDPart", "'" & sPartNumber & "'")

    If sBearingType = "LVL" Then
        dWeight = dWeight - 1.7
    End If
    If sBearingType = "FVN" Then
        dWeight = dWeight + 1#
    End If

    sWeight = Trim(Str(dWeight))

    sConnectorBom = ""

    If lConnex = 1 And (sBearingType = "LVG" Or sBearingType = "FVN") Then
        If lGirder = 1 Then
            sConnectorBom = "+ 2, 304838, Guide MIC-PS90"
        Else
            sConnectorBom = "+ 2, 304839, Guide MIC-PS120"
        End If
    End If

    bstrBOMDesc = "Hilti " & sBearingType & " " & sOrderNo & sHDefault & " " & sPipeNomDia & " NW " & sConnectorBom

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
