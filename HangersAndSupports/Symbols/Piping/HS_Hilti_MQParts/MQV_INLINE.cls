VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQV_INLINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQV_INLINE.cls
'   ProgID:         HS_Hilti_MQParts.MQV_INLINE
'   Author:         SN
'   Creation Date:  Friday, November 02, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732    Detailed Physical aspect need to be removed from Smart Parts
'   28/03/12          KP        CR-CP-195618 Add Hiliti and Anvil PH2010 to default catalog
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MQV_INLINE"  'Used for error messages
Private Const m_progID = "HS_Hilti_MQParts.MQV_INLINE"

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
    AddInput Inputs, "ChanLength", "Double"
    AddInput Inputs, "Width", "Double"
    AddInput Inputs, "Depth", "Double"
    AddInput Inputs, "Thick", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "NumPush", "Long"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Structure"
    AddOutput Outputs, "Box"
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
    m_SymbolDefinition.Name = m_SymbolDefinition.ProgId
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
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Dim sPartNumber As String

    Dim dLength As Double
    Dim dWidth As Double
    Dim dDepth As Double
    Dim dThick As Double
    Dim dOffset As Double
    Dim sItemNo As String
    Dim sGroup As String
    Dim lNumPush As Long

    dLength = arrayOfInputs(2)
    dWidth = arrayOfInputs(3)
    dDepth = arrayOfInputs(4)
    dThick = arrayOfInputs(5)
    sItemNo = arrayOfInputs(6)
    sGroup = arrayOfInputs(7)
    lNumPush = arrayOfInputs(8)
    dOffset = 0.004

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        sPartNumber = ""
    Else
        oPartOcc.GetPart oPart
        sPartNumber = oPart.PartNumber
    End If
    Set oSupportComp = Nothing

    AddPort "Structure", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"

    If InStr(sPartNumber, "P4") = 0 Then
        AddComposite 8, "EXTRUDED", "", "COMP"
        AddLine -dLength / 2#, -dWidth / 2#, -dDepth + dOffset, -dLength / 2#, dThick - dWidth / 2#, -dDepth + dOffset, m_outputColl, "", "LeftBottom"
        AddLine -dLength / 2#, dThick - dWidth / 2#, -dDepth + dOffset, -dLength / 2#, dThick - dWidth / 2#, dOffset - dThick, m_outputColl, "", "InLeftBtoT"
        AddLine -dLength / 2#, dThick - dWidth / 2#, dOffset - dThick, -dLength / 2#, -dThick + dWidth / 2#, dOffset - dThick, m_outputColl, "", "InCenter"
        AddLine -dLength / 2#, -dThick + dWidth / 2#, dOffset - dThick, -dLength / 2#, -dThick + dWidth / 2#, -dDepth + dOffset, m_outputColl, "", "InRightTtoB"
        AddLine -dLength / 2#, -dThick + dWidth / 2#, -dDepth + dOffset, -dLength / 2#, dWidth / 2#, -dDepth + dOffset, m_outputColl, "", "RightBottom"
        AddLine -dLength / 2#, dWidth / 2#, -dDepth + dOffset, -dLength / 2#, dWidth / 2#, dOffset, m_outputColl, "", "OutRightBtoT"
        AddLine -dLength / 2#, dWidth / 2#, dOffset, -dLength / 2#, -dWidth / 2#, dOffset, m_outputColl, "", "OutCenter"
        AddLine -dLength / 2#, -dWidth / 2#, dOffset, -dLength / 2#, -dWidth / 2#, -dDepth + dOffset, m_outputColl, "", "OutCenter"
        AddExtrusion dLength, 0, 0, 1, m_outputColl, "Box", "Box"
    Else
        AddBox dLength, dWidth, dThick, Loc(-dLength / 2#, -dWidth / 2#, 0), m_outputColl, "Box", "Box"
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
    Dim dDryWeight As Double
    Dim sPartNumber As String
    Dim dNumPush As Double

    sPartNumber = Trim(oPart.PartNumber)

    dDryWeight = GetAttributeFromObject(oPart, "DryWeight")
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    dNumPush = Val(GetAttributeFromObject(oPart, "NumPush"))

    If Right(Trim(sPartNumber), 1) = "R" Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 304012, Channel connector MQN-R"
    Else
        If Right(Trim(sPartNumber), 1) = "F" Then
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 387779, Channel Connector MQN-HDG plus"
        Else
            bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 4, 369623, Channel connector MQN"
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
