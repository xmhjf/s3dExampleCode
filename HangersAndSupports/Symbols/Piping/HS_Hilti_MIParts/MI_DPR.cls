VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_DPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MI_DPR.cls
'   ProgID:         HS_Hilti_MIParts.MI_DPR
'   Author:         JRK
'   Creation Date:  Friday, October 26, 2007

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
Private Const MODULE = "MI_DPR"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_DPR"

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
    AddInput Inputs, "Spacer_W", "Double"
    AddInput Inputs, "Spacer_D", "Double"
    AddInput Inputs, "CompLength", "Double"
    AddInput Inputs, "CompHeight", "Double"
    AddInput Inputs, "RollerDia", "Double"
    AddInput Inputs, "PlateThick", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "Finish", "Long"
    AddInput Inputs, "Group", "String"
    AddInput Inputs, "PlateDepth", "Double"
    AddInput Inputs, "PlateWidth", "Double"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Structure"
    AddOutput Outputs, "BasePlate"
    AddOutput Outputs, "Spacer"
    AddOutput Outputs, "RightRoller"
    AddOutput Outputs, "LeftRoller"
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

    Dim dSpacerW As Double
    Dim dSpacerD As Double
    Dim dSpacerH As Double
    Dim dCompLength As Double
    Dim dCompHeight As Double
    Dim dRollerDia As Double
    Dim dPlateThick As Double
    Dim dPlateDepth As Double
    Dim dPlateWidth As Double
    Dim dRollerAngle As Double
    Dim dE As Double
    Dim dRollerL As Double
    Dim dPipeDia As Double
    Dim dInsulat As Double

    dSpacerW = arrayOfInputs(2)
    dSpacerD = arrayOfInputs(3)
    dCompLength = arrayOfInputs(4)    'UpperWidth
    dCompHeight = arrayOfInputs(5)    'OuterH
    dRollerDia = arrayOfInputs(6)
    dPlateThick = arrayOfInputs(7)
    dPlateDepth = arrayOfInputs(11)
    dPlateWidth = arrayOfInputs(12)
    dRollerAngle = 25#

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dInsulat = 0
        dPipeDia = 0.219075
    Else
        Hilti_InitializeMyHH oICH
        GetPipeDiameterAndInsulatTh 1, dPipeDia, dInsulat
        Hilti_DestroyMyHH oICH
        Set oICH = Nothing
    End If

    If dPipeDia < 0.219 Or dPipeDia > 0.41 Then
        PF_EventHandler "Pipe size not valid.", Err, MODULE, METHOD, True
    End If

    dE = 0.532 * dPipeDia + 0.054
    dSpacerD = dRollerDia + 0.01
    dSpacerH = dE - dPipeDia / 2# - 0.005 - dPlateThick
    dRollerL = 0.11     ' calc from approx angle and height and assumed spacer width

    AddPort "Route", 0, 0, 0, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"
    AddPort "Structure", 0, 0, -dE, _
            -1, 0, 0, _
            0, 0, -1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"

    'Add clamp boxes
    AddBox dPlateDepth, dPlateWidth, dPlateThick, Loc(-dPlateDepth / 2, -dPlateWidth / 2, -dE), m_outputColl, "BasePlate", "BasePlate"
    AddBox dSpacerD, dSpacerW, dSpacerH, Loc(-dSpacerD / 2, -dSpacerW / 2, -dE + dPlateThick), m_outputColl, "Spacer", "Spacer"
    'Add rollers
    AddCylinder dRollerL, dRollerDia / 2, "ROTX(" + Str(270 + dRollerAngle) + ")*" + Loc(0, dSpacerW / 2#, -dE + dPlateThick + dSpacerH / 2), m_outputColl, "RightRoller", "RightRoller"
    AddCylinder dRollerL, dRollerDia / 2, "ROTX(" + Str(90 - dRollerAngle) + ")*" + Loc(0, -dSpacerW / 2#, -dE + dPlateThick + dSpacerH / 2), m_outputColl, "LeftRoller", "LeftRoller"

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
    Dim sWeight As String

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
