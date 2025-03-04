VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MFP_GP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MFP_GP.cls
'   ProgID:         HS_Hilti_MQParts.MFP_GP
'   Author:         JOGN
'   Creation Date:  Monday, November 28, 2007

'   Description:
'       Pre-assembled 90
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP         Addition to S3D
'   02/09/11          Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14/09/11          VSP         TR-CP-193697  Incorrect BOM Description and Warnings were observed in part placement
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MFP_GP"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MFP_GP"

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
    AddInput Inputs, "ConnType", "Long"
    AddInput Inputs, "H", "Double"
    AddInput Inputs, "A", "Double"
    AddInput Inputs, "BaseL", "Double"
    AddInput Inputs, "BaseW", "Double"
    AddInput Inputs, "BaseThick", "Double"
    AddInput Inputs, "RodDia", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Top"
    AddOutput Outputs, "Bottom"
    AddOutput Outputs, "Plate"
    AddOutput Outputs, "Hole"
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

    Dim lConnType As Long
    Dim dH As Double
    Dim dA As Double
    Dim dBaseL As Double
    Dim dBaseW As Double
    Dim dBaseThick As Double
    Dim dRodDia As Double
    Dim sItemNo As String
    Dim sGroup As String

    lConnType = arrayOfInputs(2)
    dH = arrayOfInputs(3)
    dA = arrayOfInputs(4)
    dBaseL = arrayOfInputs(5)
    dBaseW = arrayOfInputs(6)
    dBaseThick = arrayOfInputs(7)
    dRodDia = arrayOfInputs(8)
    sItemNo = arrayOfInputs(9)
    sGroup = arrayOfInputs(10)

    AddPort "Top", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Top", arrayOfInputs(1), "Top"

    AddPort "Bottom", 0, 0, dBaseThick, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Bottom", arrayOfInputs(1), "Bottom"

    AddBox dBaseL, dBaseW, dBaseThick, Loc(-dBaseL / 2, -dBaseW / 2, 0), m_outputColl, "Plate"
    AddCylinder dH, dRodDia / 2, Loc(0, 0, dBaseThick), m_outputColl, "Hole"

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
    Dim lStructCon As Long
    Dim sStructCon As String

    lStructCon = GetAttributeFromObject(pSupportComp, "ConnType")

    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    If Mid(Trim(sPartNumber), 14, 1) = "R" Then
        sStructCon = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeSS", "ConnType", lStructCon)

        If Trim(sStructCon) <> "None" Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 2, " & sStructCon
        Else
            bstrBOMDesc = Hilti_BuildBom(pSupportComp)
        End If
    Else
        If Right(Trim(sPartNumber), 1) = "F" Then
            sStructCon = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeHDG", "ConnType", lStructCon)

            If Trim(sStructCon) <> "None" Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 2, " & sStructCon
            Else
                bstrBOMDesc = Hilti_BuildBom(pSupportComp)
            End If
        Else
            sStructCon = ConvertValueToLongStringValue(pSupportComp, "IJOAMQConnTypeGalv", "ConnType", lStructCon)

            If Trim(sStructCon) <> "None" Then
                bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 2, " & sStructCon
            Else
                bstrBOMDesc = Hilti_BuildBom(pSupportComp)
            End If
        End If
    End If

    Set oPartOcc = Nothing
    Set oPart = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub









