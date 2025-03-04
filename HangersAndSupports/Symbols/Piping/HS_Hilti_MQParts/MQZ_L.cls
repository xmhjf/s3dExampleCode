VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQZ_L"
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
'   26/05/16        PR          TR 288282    Unable to modify the Hilti MQ Parts 'Channel Endcaps' to 'Connection Plates'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQZ_L"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQZ_L"

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
    AddInput Inputs, "Length", "Double"
    AddInput Inputs, "W", "Double"
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "T", "Double"
    AddInput Inputs, "NumPlate", "Long"
    AddInput Inputs, "PlateGap", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "BoltD", "Double"
    AddInput Inputs, "WithBolt", "Long"
    AddInput Inputs, "NutThick", "Double"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Top"
    AddOutput Outputs, "Bottom"
    AddOutput Outputs, "Plate"
    AddOutput Outputs, "PlateOption"
    AddOutput Outputs, "Bolt"
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

    Dim LENGTH As Double
    Dim dW As Double
    Dim dD As Double
    Dim dT As Double
    Dim lNumPlate As Long
    Dim dPlateGap As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim dBoltD As Double
    Dim lWithBolt As Long
    Dim dNutThick As Double

    LENGTH = arrayOfInputs(2)
    dW = arrayOfInputs(3)
    dD = arrayOfInputs(4)
    dT = arrayOfInputs(5)
    lNumPlate = arrayOfInputs(6)
    dPlateGap = arrayOfInputs(7)
    sItemNo = arrayOfInputs(8)
    sGroup = arrayOfInputs(9)
    dBoltD = arrayOfInputs(10)
    lWithBolt = arrayOfInputs(11)
    dNutThick = arrayOfInputs(12)

    AddPort "Top", 0, 0, 0, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Top", arrayOfInputs(1), "Top"

    AddPort "Bottom", 0, 0, -dT, _
            1, 0, 0, _
            0, 0, 1, m_outputColl, "Bottom", arrayOfInputs(1), "Bottom"

    AddBox dD, dW, dT, Loc(-dD / 2, -dW / 2, -dT), m_outputColl, "Plate"

    If dPlateGap > LENGTH - ((2 * dNutThick) + dT) Then
        PF_EventHandler "Plate gap has exceeded the maximum distance based on the bolt length.", Err, MODULE, METHOD, True
    End If

    If lWithBolt = 2 Then
        AddCylinder LENGTH + ((2 * dNutThick) + dT), dBoltD / 2, Loc(0, 0, -(2 * dNutThick) - dT), m_outputColl, "Bolt"
       
    End If

    If lNumPlate = 2 Then
        AddBox dD, dW, dT, Loc(-dD / 2, -dW / 2, dPlateGap), m_outputColl, "Plate"
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

    Dim sItemNo As String
    Dim sPartNumber As String
    Dim dblLength As Double
    Dim dNutThick As Double
    Dim dT As Double
    
    dblLength = GetAttributeFromObject(pSupportComp, "Length")
    
    dNutThick = GetAttributeFromObject(oPart, "NutThick")
    dT = GetAttributeFromObject(oPart, "T")
    
    If (dblLength <= 0) Then
        dblLength = 1 + ((2 * dNutThick) + dT)
        
        Dim oMyPart As IJHgrSupportComponent
        Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
        
        Set oMyPart = pSupportComp
        oMyPart.GetOccAssembly oICH
        oICH.SetAttributeValue "Length", oMyPart, dblLength
        Set oMyPart = Nothing
        Set oICH = Nothing
    End If
    
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





