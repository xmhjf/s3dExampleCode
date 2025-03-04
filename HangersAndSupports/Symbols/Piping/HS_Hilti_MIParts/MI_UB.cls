VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_UB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MI_UB.cls
'   ProgID:         HS_Hilti_MIParts.MI_UB
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
Private Const MODULE = "MI_UB"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_UB"

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_SymbolDefinition As IMSSymbolEntities.DSymbolDefinition
Private m_outputColl As IJDOutputCollection
Dim oCodeListData As IJDCodeListMetaData

Implements IJDUserSymbolServices
Implements IJHgrSymbolBOMServices
Implements IJHgrSymbolWCGServices

Private Sub IJHgrSymbolWCGServices_EvaluateLocalWCG(ByVal pSupportComp As Object, Weight As Double, CogX As Double, CogY As Double, CogZ As Double)
    Const METHOD = " IJHgrSymbolWCGServices_EvaluateLocalWCG"
    On Error GoTo ErrorHandler

    Dim dWeight As Double
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp
    oPartOcc.GetPart oPart

    Dim lGirder As Long
    Dim lAngles As Long
    Dim dRodDia As Double
    Dim sPartNumber As String

    sPartNumber = oPart.PartNumber
    lGirder = GetAttributeFromObject(pSupportComp, "Girder")
    lAngles = GetAttributeFromObject(pSupportComp, "Angles")
    dRodDia = GetAttributeFromObject(oPart, "RodDia")
    dWeight = MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASMI_StandardUBolt", "PartNumber", "JDPart", "'" & sPartNumber & "'")

    If lGirder = 1 And lAngles = 1 And dRodDia < 0.016 Then
        dWeight = dWeight + 1.26
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia < 0.016 Then
        dWeight = dWeight + 1.38
    End If

    If lGirder = 1 And lAngles = 1 And dRodDia > 0.015 Then
        dWeight = dWeight + 1.24
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia > 0.015 Then
        dWeight = dWeight + 1.36
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
    AddInput Inputs, "Finish", "Long"
    AddInput Inputs, "Girder", "Long"
    AddInput Inputs, "Angles", "Long"
    AddInput Inputs, "Lining", "Long"
    AddInput Inputs, "RodDia", "Double"
    AddInput Inputs, "H0", "Double"
    AddInput Inputs, "E", "Double"
    AddInput Inputs, "D", "Double"
    AddInput Inputs, "ItemNo", "String"
    AddInput Inputs, "AddInputs_Done", ""

    'Outputs
    Dim Outputs As tOutputs
    AddOutput Outputs, "Route"
    AddOutput Outputs, "Structure"
    AddOutput Outputs, "Bend"
    AddOutput Outputs, "R"
    AddOutput Outputs, "L"
    AddOutput Outputs, "Box1"
    AddOutput Outputs, "Box2"
    AddOutput Outputs, "Box3"
    AddOutput Outputs, "Box4"
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
    Dim lGirder As Long
    Dim lAngles As Long
    Dim lLining As Long
    Dim dRodDia As Double
    Dim dH As Double
    Dim dE As Double
    Dim sItemNo As String
    Dim dAnglesOffset As Double
    Dim dPipeDia As Double
    Dim dAnglesSideW As Double
    Dim dAnglesLength As Double
    Dim dAnglesTopW As Double
    Dim dAnglesThickness As Double
    Dim dOffset As Double

    lFinish = arrayOfInputs(2)
    lGirder = arrayOfInputs(3)
    lAngles = arrayOfInputs(4)
    lLining = arrayOfInputs(5)
    dRodDia = arrayOfInputs(6)
    dH = arrayOfInputs(7)
    dE = arrayOfInputs(8)
    dPipeDia = arrayOfInputs(9)
    sItemNo = arrayOfInputs(10)

    dAnglesLength = 0.09
    dAnglesTopW = 0.049
    dAnglesOffset = 0.0285
    dAnglesThickness = 0.006

    If lGirder = 1 And lAngles = 1 And dRodDia < 0.016 Then
        dAnglesSideW = 0.09
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia < 0.016 Then
        dAnglesSideW = 0.096
    End If

    If lGirder = 1 And lAngles = 1 And dRodDia > 0.015 Then
        dAnglesSideW = 0.09
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia > 0.015 Then
        dAnglesSideW = 0.096
    End If

    If dE < dAnglesLength Then
        dOffset = 0.02
    Else
        dOffset = 0
    End If

    If lAngles = 2 Then
        AddPort "Route", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"
        AddPort "Structure", 0, 0, -dPipeDia / 2, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"
    Else
        AddPort "Route", 0, 0, 0, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "Route", arrayOfInputs(1), "Route"
        AddPort "Structure", -dAnglesOffset, 0, -dPipeDia / 2 - dAnglesSideW / 2 - dRodDia / 2, _
                1, 0, 0, _
                0, 0, -1, m_outputColl, "Structure", arrayOfInputs(1), "Structure"
    End If


    'Create the Graphics
    AddElbow dRodDia / 2#, dE / 2#, 180, "ROTZ(90)*" + Loc(0, 0, 0), m_outputColl, "Bend", "Bend"
    AddCylinder dH - dE / 2# - dRodDia / 2#, dRodDia / 2#, Loc(0, dE / 2#, -dH + dE / 2# + dRodDia / 2#), m_outputColl, "R", "R"
    AddCylinder dH - dE / 2# - dRodDia / 2#, dRodDia / 2#, Loc(0, -dE / 2#, -dH + dE / 2# + dRodDia / 2#), m_outputColl, "L", "L"

    If lAngles = 1 Then
        AddBox dAnglesTopW, dAnglesLength, dAnglesThickness, Loc(-dAnglesOffset, dE / 2# - dAnglesLength / 2# + dOffset, -dPipeDia / 2# - dRodDia / 2# - dAnglesThickness), m_outputColl, "Box1", "Box1"
        AddBox dAnglesThickness, dAnglesLength, dAnglesSideW - dAnglesThickness, Loc(-dAnglesOffset, dE / 2# - dAnglesLength / 2# + dOffset, -dPipeDia / 2# - dRodDia / 2# - dAnglesThickness - dAnglesSideW + dAnglesThickness), m_outputColl, "Box2", "Box2"
        AddBox dAnglesTopW, dAnglesLength, dAnglesThickness, Loc(-dAnglesOffset, -dE / 2# - dAnglesLength / 2# - dOffset, -dPipeDia / 2# - dRodDia / 2# - dAnglesThickness), m_outputColl, "Box3", "Box3"
        AddBox dAnglesThickness, dAnglesLength, dAnglesSideW - dAnglesThickness, Loc(-dAnglesOffset, -dE / 2# - dAnglesLength / 2# - dOffset, -dPipeDia / 2# - dRodDia / 2# - dAnglesThickness - dAnglesSideW + dAnglesThickness), m_outputColl, "Box4", "Box4"
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
    Dim lAngles As Long
    Dim lLining As Long
    Dim dRodDia As Double
    Dim sAnglesBom As String
    Dim sItemName As String
    Dim sItemNo As String
    Dim sPartNumber As String
    Dim dBOMWeight As Double
    Dim sBOMWeight As String

    sPartNumber = Trim(oPart.PartNumber)
    lGirder = GetAttributeFromObject(pSupportComp, "Girder")
    lAngles = GetAttributeFromObject(pSupportComp, "Angles")
    lLining = GetAttributeFromObject(pSupportComp, "Lining")
    dRodDia = GetAttributeFromObject(oPart, "RodDia")
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")

    dBOMWeight = Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", "JCUHASMI_StandardUBolt", "PartNumber", "JDPart", "'" & sPartNumber & "'")

    If dBOMWeight < 1 Then
        sBOMWeight = "0" & Trim(Str(dBOMWeight))
    Else
        sBOMWeight = Trim(Str(dBOMWeight))
    End If

    If lGirder = 1 And lAngles = 1 And dRodDia < 0.016 Then
        sAnglesBom = "+ 1, 304831, Connector U-Bolt MIC-UB90-M12"
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia < 0.016 Then
        sAnglesBom = "+ 1, 304833, Connector U-Bolt MIC-UB120-M12"
    End If

    If lGirder = 1 And lAngles = 1 And dRodDia > 0.015 Then
        sAnglesBom = "+ 1, 304834, Connector U-Bolt MIC-UB90-M16"
    End If

    If lGirder = 2 And lAngles = 1 And dRodDia > 0.015 Then
        sAnglesBom = "+ 1, 304835, Connector U-Bolt MIC-UB120-M16"
    End If

    If lAngles = 2 Then
        If lLining = 2 Then
            bstrBOMDesc = "Hilti, " + sItemNo + ", " + Replace(oPart.PartDescription, "Hilti ", "")
        Else
            bstrBOMDesc = "Hilti, " + sItemNo + ", " + Replace(oPart.PartDescription, "Hilti ", "") + " LINED"
        End If
    Else
        If lLining = 2 Then
            bstrBOMDesc = "Hilti, " + sItemNo + ", " + Replace(oPart.PartDescription, "Hilti ", "") + " " + sAnglesBom
        Else
            bstrBOMDesc = "Hilti, " + sItemNo + ", " + Replace(oPart.PartDescription, "Hilti ", "") + " LINED " + sAnglesBom
        End If
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub

