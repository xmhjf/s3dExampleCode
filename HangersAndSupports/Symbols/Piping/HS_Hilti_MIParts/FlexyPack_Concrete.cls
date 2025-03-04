VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexyPack_Concrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   FlexyPack_Concrete.cls
'   ProgID:         HS_Hilti_MIParts.FlexyPack_Concrete
'   Author:         SN
'   Creation Date:  October 18, 2007

'   Description:
'       FlexyPack connector to Steel
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP         Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "FlexyPack_Concrete"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.FlexyPack_Concrete"

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

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Dim dPlateWeight As Double
    Dim dNozzelWeight As Double
    Dim lConnectionSize As Long
    Dim record As String
    Dim pParameters(0) As Parameters

    Dim dPlateWidth As Double

    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    dPlateWeight = GetAttributeFromObject(oPart, "WaterWeight")
    lConnectionSize = GetAttributeFromObject(pSupportComp, "ConnectionSize")

    pParameters(0) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
    record = ReadParametricData("JUAMIFlexyPack_Con_AUX", "where ConnectionSize = ?")
    dNozzelWeight = GetNParamUsingParametricQuery(record, "NozzelWeight", pParameters)

    Weight = dPlateWeight + dNozzelWeight

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
    Const iNumInputs As Integer = 16
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 5
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "PlateThick"
    sDataType(2) = "Double"

    sInputName(3) = "ItemNo"
    sDataType(3) = "String"

    sInputName(4) = "Finish"
    sDataType(4) = "Long"

    sInputName(5) = "Group"
    sDataType(5) = "String"

    sInputName(6) = "NozzelLength"
    sDataType(6) = "Double"

    sInputName(7) = "NozzelWidth"
    sDataType(7) = "Double"

    sInputName(8) = "NozzelDepth"
    sDataType(8) = "Double"

    sInputName(9) = "PlateDepth"
    sDataType(9) = "Double"

    sInputName(10) = "PlateWidth"
    sDataType(10) = "Double"

    sInputName(11) = "PortOffset"
    sDataType(11) = "Double"

    sInputName(12) = "NozzelYOffset"
    sDataType(12) = "Double"

    sInputName(13) = "NozzelZOffset"
    sDataType(13) = "Double"

    sInputName(14) = "NozzelRotation"
    sDataType(14) = "Double"

    sInputName(15) = "ConnectionSize"
    sDataType(15) = "Long"

    sInputName(16) = "M16Anchors"
    sDataType(16) = "Long"

    'Outputs
    sOutputName(1) = "StructureInt"
    sOutputName(2) = "StructureMid"
    sOutputName(3) = "StructureExt"
    sOutputName(4) = "Plate"
    sOutputName(5) = "girderATT"

    SetupInputsAndOutputs pSymbolDefinition, iNumInputs, iNumOutputs, sInputName(), sDataType(), sOutputName(), m_progID

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

    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dFlangeThick As Double
    Dim dWebThick As Double
    Dim dPipeDia As Double
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
    Dim lConnectionSize As Long
    Dim dNozzelYOffset As Double
    Dim dNozzelZOffset As Double
    Dim dNozzelRotation As Double
    Dim lM16Anchors As Long
    Dim dAngle As Double
    Dim record As String

    Dim dNozzelYZ(2) As Double
    Dim dNozzel1(2) As Double
    Dim dNozzel2(2) As Double
    Dim dNozzel3(2) As Double
    Dim dNozzel4(2) As Double

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
    dNozzelYOffset = arrayOfInputs(12)
    dNozzelZOffset = arrayOfInputs(13)
    dNozzelRotation = arrayOfInputs(14)
    lConnectionSize = arrayOfInputs(15)
    lM16Anchors = arrayOfInputs(16)

    Dim pParameters(0 To 0) As Parameters
    pParameters(0) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
    record = ReadParametricData("JUAMIFlexyPack_Con_AUX", "where ConnectionSize = ?")
    dNozzelDepth = GetNParamUsingParametricQuery(record, "NozzelDepth", pParameters)
    dNozzelWidth = GetNParamUsingParametricQuery(record, "NozzelWidth", pParameters)

    'Translate the angle value from Radians to Degrees for using the rotation function(ROTX).
    If dNozzelRotation = 0# Then
        dAngle = 0#
    Else
        dAngle = dNozzelRotation * 180 / GetPi
    End If

    AddPort "StructureInt", 0, 0, 0, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureInt", arrayOfInputs(1), "StructureInt"

    AddPort "StructureMid", dPlateThick, -dNozzelYOffset, -dNozzelZOffset, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureMid", arrayOfInputs(1), "StructureMid"

    AddPort "StructureExt", dPlateThick + dPortOffset, -dNozzelYOffset, -dNozzelZOffset, _
            1, 0, 0, _
            0, 0, -1, m_outputColl, "StructureExt", arrayOfInputs(1), "StructureExt"

    AddBox dPlateThick, dPlateWidth, dPlateDepth, Loc(0, -dPlateWidth / 2, -dPlateDepth / 2#), m_outputColl, "Plate"

    dNozzelYZ(1) = dNozzelWidth / 2#
    dNozzelYZ(2) = dNozzelDepth / 2#
    dNozzel1(1) = -dNozzelYZ(1) * Cos(dNozzelRotation) - dNozzelYZ(2) * Sin(dNozzelRotation)
    dNozzel1(2) = -dNozzelYZ(1) * Sin(dNozzelRotation) + dNozzelYZ(2) * Cos(dNozzelRotation)
    dNozzel2(1) = -dNozzelYZ(1) * Cos(dNozzelRotation) + dNozzelYZ(2) * Sin(dNozzelRotation)
    dNozzel2(2) = -dNozzelYZ(1) * Sin(dNozzelRotation) - dNozzelYZ(2) * Cos(dNozzelRotation)
    dNozzel3(1) = dNozzelYZ(1) * Cos(dNozzelRotation) + dNozzelYZ(2) * Sin(dNozzelRotation)
    dNozzel3(2) = dNozzelYZ(1) * Sin(dNozzelRotation) - dNozzelYZ(2) * Cos(dNozzelRotation)
    dNozzel4(1) = dNozzelYZ(1) * Cos(dNozzelRotation) - dNozzelYZ(2) * Sin(dNozzelRotation)
    dNozzel4(2) = dNozzelYZ(1) * Sin(dNozzelRotation) + dNozzelYZ(2) * Cos(dNozzelRotation)
    'create nozzel graphic
    AddComposite 4, "EXTRUDED", "", "COMP"
    AddLine dPlateThick, dNozzel1(1) - dNozzelYOffset, dNozzel1(2) - dNozzelZOffset, dPlateThick, dNozzel2(1) - dNozzelYOffset, dNozzel2(2) - dNozzelZOffset, m_outputColl, "", "TOP"
    AddLine dPlateThick, dNozzel2(1) - dNozzelYOffset, dNozzel2(2) - dNozzelZOffset, dPlateThick, dNozzel3(1) - dNozzelYOffset, dNozzel3(2) - dNozzelZOffset, m_outputColl, "", "TOP"
    AddLine dPlateThick, dNozzel3(1) - dNozzelYOffset, dNozzel3(2) - dNozzelZOffset, dPlateThick, dNozzel4(1) - dNozzelYOffset, dNozzel4(2) - dNozzelZOffset, m_outputColl, "", "TOP"
    AddLine dPlateThick, dNozzel4(1) - dNozzelYOffset, dNozzel4(2) - dNozzelZOffset, dPlateThick, dNozzel1(1) - dNozzelYOffset, dNozzel1(2) - dNozzelZOffset, m_outputColl, "", "TOP"
    AddExtrusion dNozzelLength, 0, 0, 1, m_outputColl, "girderATT", "girderATT"

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

    Dim sPlateItemNo As String
    Dim sNozzleItemNo As String
    Dim sGroupName As String
    Dim sPlateDesc As String
    Dim sNozzleDesc As String
    
    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    
    Set oSupportComp = pSupportComp
    oSupportComp.GetOccAssembly oICH

    Dim lConnectionSize As Long
    Dim record As String
    Dim pParameters(0) As Parameters
    lConnectionSize = GetAttributeFromObject(pSupportComp, "ConnectionSize") '
    sGroupName = GetAttributeFromObject(pSupportComp, "Group")

    pParameters(0) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
    record = ReadParametricData("JUAMIFlexyPack_Con_AUX", "where ConnectionSize = ?")
    sPlateItemNo = GetSParamUsingParametricQuery(record, "PlateItemNo", pParameters)
    sNozzleItemNo = GetSParamUsingParametricQuery(record, "NozzelItemNo", pParameters)

    pParameters(0) = FillParameterForBSTR("Item_Nb", sPlateItemNo)
    record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
    sPlateDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters)

    pParameters(0) = FillParameterForBSTR("Item_Nb", sNozzleItemNo)
    record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
    sNozzleDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters)

    If Not oICH Is Nothing Then
        oICH.SetAttributeValue "ItemNo", oSupportComp, sPlateItemNo & ", " & sNozzleItemNo
        
        Set oICH = Nothing
    End If
    
     bstrBOMDesc = "Hilti, " & sPlateItemNo & ", " & Replace(sPlateDesc, "Hilti ", "") _
                        & " + 1, " & sNozzleItemNo & ", " & Replace(sNozzleDesc, "Hilti ", "")
    
    Set oSupportComp = Nothing
    Set oPartOcc = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub




