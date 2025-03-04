VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MP_MX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MP_MX.cls
'   ProgID:         HS_Hilti_MQParts.MP_MX
'   Author:         JRK
'   Creation Date:  Friday, November 9, 2007

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
Private Const MODULE = "MP_MX"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MP_MX"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Implements IJHgrSymbolWCGServices
Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices

Private Sub IJHgrSymbolWCGServices_EvaluateLocalWCG(ByVal pSupportComp As Object, Weight As Double, CogX As Double, CogY As Double, CogZ As Double)
    Const METHOD = " IJHgrSymbolWCGServices_EvaluateLocalWCG"
    On Error GoTo ErrorHandler

    ' <Determine BOM description here>
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim dWeight As Double
    Dim dPi As Double
    Dim dD As Double
    Dim dS As Double
    Dim dB As Double
    Dim dSpacer As Double
    Dim dL As Double
    Dim dBoltH As Double
    Dim dBoltD As Double

    dD = GetAttributeFromObject(oPart, "D") * 1000
    dS = GetAttributeFromObject(oPart, "S") * 1000
    dB = GetAttributeFromObject(oPart, "B") * 1000
    dL = GetAttributeFromObject(oPart, "L") * 1000
    dBoltD = GetAttributeFromObject(oPart, "BoltD") * 1000
    dSpacer = 2
    dBoltH = dSpacer + 2 * dS + 2 * dBoltD

    dPi = 3.14159265358979

    dWeight = (((dD / 2# + dS) * (dD / 2# + dS) * dPi * dB) - (dD / 2# * dD / 2# * dPi * dB) _
             - ((dS * dB * dSpacer) * 2#) + (((dL / 2# - dD / 2# - dS) * dB * dS) * 4#) + _
               ((dPi * dBoltH * (dBoltD / 2#) * (dBoltD / 2#)) * 2#)) * 0.0000079

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
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "H", "Double"
    AddInput Inputs, "B", "Double"
    AddInput Inputs, "BoltD", "Double"
    AddInput Inputs, "L", "Double"
    AddInput Inputs, "S", "Double"
    AddInput Inputs, "FRec", "Double"
    AddInput Inputs, "BossDiaMet", "Double"
    AddInput Inputs, "BossDiaFrac", "Double"
    AddInput Inputs, "BOMAddMet", "String"
    AddInput Inputs, "BOMAddFrac", "String"
    AddInput Inputs, "BossCode", "String"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Connector"
    AddOutput Outputs, "PipeClampBody"
    AddOutput Outputs, "LeftBottom"
    AddOutput Outputs, "LeftTop"
    AddOutput Outputs, "LeftBolt"
    AddOutput Outputs, "RightBottom"
    AddOutput Outputs, "RightTop"
    AddOutput Outputs, "RightBolt"
    AddOutput Outputs, "Top"
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
    Dim dBoltD As Double
    Dim dL As Double
    Dim dS As Double
    Dim dFRec As Double
    Dim dBossDiaMet As Double
    Dim dBossDiaFrac As Double
    Dim sBossCode As String
    Dim sItemNo As String
    Dim sGroup As String
    Dim dSpacer As Double
    Dim dBoltH As Double
    Dim dBossDia As Double

    dD = arrayOfInputs(2)
    dH = arrayOfInputs(3)
    dB = arrayOfInputs(4)
    dBoltD = arrayOfInputs(5)
    dL = arrayOfInputs(6)
    dS = arrayOfInputs(7)
    dFRec = arrayOfInputs(8)
    dBossDiaMet = arrayOfInputs(9)
    dBossDiaFrac = arrayOfInputs(10)
    sBossCode = arrayOfInputs(13)
    sItemNo = arrayOfInputs(14)
    sGroup = arrayOfInputs(15)
    dSpacer = 0.002
    dBoltH = dSpacer + 2 * dS + 2 * dBoltD

    If UCase(Trim(sBossCode)) = "MET" Then
        dBossDia = dBossDiaMet
    Else
        dBossDia = dBossDiaFrac
    End If

    AddPort "Route", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"
    AddPort "Connector", 0, 0, dD / 2 + dS, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Connector", arrayOfInputs(1), "Connector"

    AddCylinder dB, dD / 2# + dS, "ROTY(90)*" + Loc(-dB / 2#, 0, 0), m_outputColl, "PipeClampBody", "PipeClampBody"

    AddBox dB, (dL - dD) / 2#, dS, Loc(-dB / 2#, -dL / 2#, -dS - dSpacer / 2#), m_outputColl, "LeftBottom", "LeftBottom"
    AddBox dB, (dL - dD) / 2#, dS, Loc(-dB / 2#, -dL / 2#, dSpacer / 2#), m_outputColl, "LeftTop", "LeftTop"
    AddCylinder dBoltH, dBoltD / 2#, "ROTZ(90)*" & Loc(0, -dL / 2# + ((dL - dD) / 2) / 2, -dBoltH / 2#), m_outputColl, "LeftBolt", "LeftBolt"

    AddBox dB, (dL - dD) / 2#, dS, Loc(-dB / 2#, dD / 2#, -dS - dSpacer / 2#), m_outputColl, "RightBottom", "RightBottom"
    AddBox dB, (dL - dD) / 2#, dS, Loc(-dB / 2#, dD / 2#, dSpacer / 2#), m_outputColl, "RightTop", "RightTop"
    AddCylinder dBoltH, dBoltD / 2#, "ROTZ(90)*" & Loc(0, dD / 2# + ((dL - dD) / 2) / 2, -dBoltH / 2#), m_outputColl, "RightBolt", "RightBolt"

    AddCylinder dH - dD / 2, dBossDia / 2, "ROTZ(90)*" & Loc(0, 0, dD / 2#), m_outputColl, "Top", "Top"

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
    Dim sBOMAdd As String
    Dim dWeight As Double
    Dim sPartNumber As String
    Dim dPi As Double
    Dim dD As Double
    Dim dS As Double
    Dim dB As Double
    Dim dSpacer As Double
    Dim dL As Double
    Dim dBoltH As Double
    Dim dBoltD As Double
    Dim sBossCode As String

    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sBossCode = Trim(GetAttributeFromObject(oPart, "BossCode"))
    dD = GetAttributeFromObject(oPart, "D") * 1000
    dS = GetAttributeFromObject(oPart, "S") * 1000
    dB = GetAttributeFromObject(oPart, "B") * 1000
    dL = GetAttributeFromObject(oPart, "L") * 1000
    dBoltD = GetAttributeFromObject(oPart, "BoltD") * 1000
    dSpacer = 2
    dBoltH = dSpacer + 2 * dS + 2 * dBoltD

    dPi = 3.14159265358979

    dWeight = Round((((dD / 2# + dS) * (dD / 2# + dS) * dPi * dB) - (dD / 2# * dD / 2# * dPi * dB) _
                   - ((dS * dB * dSpacer) * 2#) + (((dL / 2# - dD / 2# - dS) * dB * dS) * 4#) + _
                     ((dPi * dBoltH * (dBoltD / 2#) * (dBoltD / 2#)) * 2#)) * 0.0000079, 2)

    If InStr(sPartNumber, "MX") = 0 Then
        If UCase(sBossCode) = "MET" Then
            bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "")
        Else
            sBOMAdd = GetAttributeFromObject(oPart, "BOMAddFrac")
            If InStr(oPart.PartDescription, sBOMAdd) <> 0 Then
                bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "")
            Else
                bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "") & " " & sBOMAdd
            End If
        End If
    Else
        bstrBOMDesc = "Hilti, " & sItemNo & ", " & Replace(oPart.PartDescription, "Hilti ", "")
    End If
    
    Set oPartOcc = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub




