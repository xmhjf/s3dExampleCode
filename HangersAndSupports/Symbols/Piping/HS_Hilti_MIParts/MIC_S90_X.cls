VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_Base_Flange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Anvil_FIG95.cls
'   ProgID:         HS_Hilti_MIParts.MI_Base_Flange
'   Author:         JOGN
'   Creation Date:  September 24, 2007

'   Description:
'       Crossbeam connector to Steel
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP         Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts  
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MI_Base_Flange"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_Base_Flange"

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
    AddInput Inputs, "PlateThick", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Finish", "Long"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "NozzelLength", "Double"
    AddInput Inputs, "NozzelWidth", "Double"
    AddInput Inputs, "NozzelDepth", "Double"
    AddInput Inputs, "PlateDepth", "Double"
    AddInput Inputs, "PlateWidth", "Double"
    AddInput Inputs, "PortOffset", "Double"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "StructureInt"
    AddOutput Outputs, "StructureExt"
    AddOutput Outputs, "Plate"
    AddOutput Outputs, "girderATT"
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

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = oSupportComp    ' The part occurence
    oPartOcc.GetPart oPart
    Dim sPartNumber As String
    sPartNumber = oPart.PartNumber

    Set oSupportComp = Nothing
    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dFlangeThick As Double
    Dim dWebThick As Double
    Dim dPipeDia As Double
    Dim MIC_S90_AUX As String
    Dim dPlateDepth As Double
    Dim dPlateWidth As Double

    Dim dPlateThick As Double
    Dim sItemNo As String
    Dim lFinish As Long
    Dim sGroup As String
    Dim dNozzelLength As Double
    Dim dNozzelWidth As Double
    Dim dNozzelDepth As Double
    Dim dPortOffset As Double

    dPlateThick = arrayOfInputs(2)
    sItemNo = arrayOfInputs(3)
    lFinish = arrayOfInputs(4)
    sGroup = arrayOfInputs(5)
    dNozzelLength = arrayOfInputs(6)
    dNozzelWidth = arrayOfInputs(7)
    dNozzelDepth = arrayOfInputs(8)
    dPlateDepth = arrayOfInputs(9)
    dPlateWidth = arrayOfInputs(10)
    dPortOffset = arrayOfInputs(11)

    If oICH Is Nothing Then    ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275    ' 6" pipe dia
    Else
        Hilti_InitializeMyHH oICH
        ' Changed to use hshlpr instead of hs_hgrhlpr

        Hilti_DestroyMyHH oICH

        If dSteelWidth > 0.3 Then
            PF_EventHandler "Steel Width can not exceed 300 mm.", Err, MODULE, METHOD, True
        Else
            If dSteelWidth > 0.165 And sPartNumber = "MIC-S90-A" Then
                PF_EventHandler "Steel Width can not exceed 165 mm. For Part Number MIC-S90-A please select MIC-S90-B or MIC-S90-C", Err, MODULE, METHOD, True
            ElseIf dSteelWidth > 0.235 And sPartNumber = "MIC-S90-B" Then
                PF_EventHandler "Steel Width can not exceed 165 mm. For Part Number MIC-S90-A please select", Err, MODULE, METHOD, True
            End If
        End If

        If dFlangeThick < 0.003 Or dFlangeThick > 0.036 Then
            PF_EventHandler "Flange thickness must be between 3 to 36 mm", Err, MODULE, METHOD, True

        End If

        Set oICH = Nothing
    End If

    AddPort "StructureInt", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureInt", arrayOfInputs(1), "StructureInt"

    AddPort "StructureExt", dPlateThick + dPortOffset, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureExt", arrayOfInputs(1), "StructureExt"

    AddBox dPlateThick, dPlateWidth, dPlateDepth, Loc(0, -dPlateWidth / 2, -dPlateDepth / 2#), m_outputColl, "Plate"
    AddBox dNozzelLength, dNozzelWidth, dNozzelDepth, Loc(dPlateThick, -dNozzelWidth / 2, -dNozzelDepth / 2), m_outputColl, "girderATT"

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

    Dim strFinish As String
    Dim lFinish As Long
    Dim sItemNo As Long

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub


