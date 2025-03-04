VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIC_UB90_L400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MIC_UB90_L400.cls
'   ProgID:         HS_Hilti_MIParts.MIC_UB90_L400
'   Author:         JRK
'   Creation Date:  Friday, Oct 19, 2007

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
Private Const MODULE = "MIC_UB90_L400"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MIC_UB90_L400"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Dim sPartNumber As String

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
    AddInput Inputs, "Finish", "Long"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "L", "Double"
    AddInput Inputs, "T", "Double"
    AddInput Inputs, "TopW", "Double"
    AddInput Inputs, "TopOffset1", "Double"
    AddInput Inputs, "TopOffset2", "Double"
    AddInput Inputs, "TopSlotL", "Double"
    AddInput Inputs, "TopSlotW", "Double"
    AddInput Inputs, "SideW", "Double"
    AddInput Inputs, "SideOffset1", "Double"
    AddInput Inputs, "SideOffset2", "Double"
    AddInput Inputs, "SideOffset3", "Double"
    AddInput Inputs, "SideSlotL", "Double"
    AddInput Inputs, "SideSlotW", "Double"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Structure1"
    AddOutput Outputs, "Structure2"
    AddOutput Outputs, "Structure3"
    AddOutput Outputs, "Structure4"
    AddOutput Outputs, "Structure5"
    AddOutput Outputs, "Structure6"
    AddOutput Outputs, "Box1"
    AddOutput Outputs, "Box2"
    AddOutput Outputs, "Box3"
    AddOutput Outputs, "Box4"
    AddOutput Outputs, "Box5"
    AddOutput Outputs, "Box6"
    AddOutput Outputs, "Box7"
    AddOutput Outputs, "Box8"
    AddOutput Outputs, "Box9"
    AddOutput Outputs, "Box10"
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

    Dim lFinish As Long
    Dim sItemNo As String
    Dim dL As Double
    Dim dT As Double
    Dim dTopW As Double
    Dim dTopOffset1 As Double
    Dim dTopOffset2 As Double
    Dim dTopSlotL As Double
    Dim dTopSlotW As Double
    Dim dSideW As Double
    Dim dSideOffset1 As Double
    Dim dSideOffset2 As Double
    Dim dSideOffset3 As Double
    Dim dSideSlotL As Double
    Dim dSideSlotW As Double

    lFinish = arrayOfInputs(2)
    sItemNo = arrayOfInputs(3)
    dL = arrayOfInputs(4)
    dT = arrayOfInputs(5)
    dTopW = arrayOfInputs(6)
    dTopOffset1 = arrayOfInputs(7)
    dTopOffset2 = arrayOfInputs(8)
    dTopSlotL = arrayOfInputs(9)
    dTopSlotW = arrayOfInputs(10)
    dSideW = arrayOfInputs(11)
    dSideOffset1 = arrayOfInputs(12)
    dSideOffset2 = arrayOfInputs(13)
    dSideOffset3 = arrayOfInputs(14)
    dSideSlotL = arrayOfInputs(15)
    dSideSlotW = arrayOfInputs(16)

    AddPort "Structure1", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure1", arrayOfInputs(1), "Structure1"
    AddPort "Structure2", -dTopOffset1 + dT, 0, -(dSideW - dT) / 2, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure2", arrayOfInputs(1), "Structure2"
    AddPort "Structure3", 0, dL / 2 - dTopOffset2 - dTopSlotL / 3, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure3", arrayOfInputs(1), "Structure3"
    AddPort "Structure4", 0, dL / 2 - dTopOffset2 - dTopSlotL * 2 / 3, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure4", arrayOfInputs(1), "Structure4"
    AddPort "Structure5", 0, -dL / 2 + dTopOffset2 + dTopSlotL / 3, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure5", arrayOfInputs(1), "Structure5"
    AddPort "Structure6", 0, -dL / 2 + dTopOffset2 + dTopSlotL * 2 / 3, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure6", arrayOfInputs(1), "Structure6"

    'Create the Graphics
    AddBox dTopW, dL, dT, Loc(-dTopOffset1, -dL / 2#, -dT), m_outputColl, "Box1", "Box1"
    AddBox dT, dL, dSideW - dT, Loc(-dTopOffset1, -dL / 2#, -dSideW), m_outputColl, "Box2", "Box2"

    AddBox dTopSlotW, dTopSlotL, dT, Loc(-dTopSlotW / 2#, -dL / 2# + dTopOffset2, -dT), m_outputColl, "Box3", "Box3"
    AddBox dTopSlotW, dTopSlotL, dT, Loc(-dTopSlotW / 2#, dL / 2# - dTopOffset2 - dTopSlotL, -dT), m_outputColl, "Box4", "Box4"

    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, dSideOffset1 * 3# - dSideSlotL / 2#, -dSideOffset2 - dT), m_outputColl, "Box5", "Box5"
    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, dSideOffset1 - dSideSlotL / 2#, -dSideOffset2 - dT), m_outputColl, "Box6", "Box6"
    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, -dSideOffset1 * 2# - dSideSlotL / 2#, -dSideOffset2 - dT), m_outputColl, "Box7", "Box7"

    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, -dSideOffset1 * 3# - dSideSlotL / 2#, -dSideOffset3 - dT), m_outputColl, "Box8", "Box8"
    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, -dSideOffset1 - dSideSlotL / 2#, -dSideOffset3 - dT), m_outputColl, "Box9", "Box9"
    AddBox dT, dSideSlotL, dSideSlotW, Loc(-dTopOffset1, dSideOffset1 * 2# - dSideSlotL / 2#, -dSideOffset3 - dT), m_outputColl, "Box10", "Box10"

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
    Dim sBOMWeight As String
    Dim sPartNumber As String

    sPartNumber = oPart.PartNumber
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sBOMWeight = Trim(Str(Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASMI_MulitiClampsConn", "PartNumber", "JDPart", "'" & Trim(sPartNumber) & "'")))

    bstrBOMDesc = "Hilti, " + sItemNo + ", Connector U-bolt MIC-UB90-L400"

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub









