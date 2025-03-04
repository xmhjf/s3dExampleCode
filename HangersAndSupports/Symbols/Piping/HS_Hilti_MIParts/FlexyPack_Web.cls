VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexyPack_Web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   FlexyPack_Web.cls
'   ProgID:         HS_Hilti_MIParts.FlexyPack_Web
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
Private Const MODULE = "FlexyPack_Web"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.FlexyPack_Web"

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
    Dim oSupportComp As IJHgrSupportComponent
    Set oPartOcc = pSupportComp
    Dim oPart As IJDPart
    oPartOcc.GetPart oPart
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oSupportComp = pSupportComp
    oSupportComp.GetOccAssembly oICH

    ' new
    Set oSupportComp = Nothing

    Dim dPlateWeight As Double
    Dim dNozzelWeight As Double
    Dim dFlexyWeight As Double
    Dim dThreadedRodWeight As Double
    Dim dSelfLockingNutWeight As Double
    Dim dPipeDia As Double
    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dFlangeThick As Double
    Dim dWebThick As Double
    Dim lConnectionSize As Long
    Dim lConnectToSteel As Long
    Dim record As String
    Dim pParameters(0 To 2) As Parameters
        
    lConnectionSize = GetAttributeFromObject(oPart, "ConnectionSize")
    lConnectToSteel = GetAttributeFromObject(pSupportComp, "ConnectToSteel")

    If oICH Is Nothing Then        ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275        ' 6" pipe dia
    Else
        Hilti_InitializeMyHH oICH

        'newcode

        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dSteelDepth = SupportStruct.dDepth
            dSteelWidth = SupportStruct.dWidth
            dWebThick = SupportStruct.dWebThickness
            dFlangeThick = SupportStruct.dFlangeThickness
        Else
            dSteelDepth = GetAttributeFromObject(pSupportComp, "SteelDepth")
            dSteelWidth = GetAttributeFromObject(pSupportComp, "SteelWidth")
            dWebThick = GetAttributeFromObject(pSupportComp, "WebThick")
            dFlangeThick = GetAttributeFromObject(pSupportComp, "FlangeThick")
        End If

        Set oStruct = Nothing
        
        Hilti_DestroyMyHH oICH

        If lConnectToSteel = 1 Then
            If dSteelDepth <= 0.3 And dSteelDepth >= 0.075 Then
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelDepth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelDepth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_Web_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                dPlateWeight = GetNParamUsingParametricQuery(record, "PlateWeight", pParameters)
                dNozzelWeight = GetNParamUsingParametricQuery(record, "NozzelWeight", pParameters)
                dFlexyWeight = GetNParamUsingParametricQuery(record, "FlexyWeight", pParameters)
                dThreadedRodWeight = GetNParamUsingParametricQuery(record, "ThreadedRodWeight", pParameters)
                dSelfLockingNutWeight = GetNParamUsingParametricQuery(record, "SelfLockingNutWeight", pParameters)
                           
            End If
        Else
            If dSteelWidth <= 0.3 And dSteelWidth >= 0.075 Then
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelWidth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelWidth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_Web_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                dPlateWeight = GetNParamUsingParametricQuery(record, "PlateWeight", pParameters)
                dNozzelWeight = GetNParamUsingParametricQuery(record, "NozzelWeight", pParameters)
                dFlexyWeight = GetNParamUsingParametricQuery(record, "FlexyWeight", pParameters)
                dThreadedRodWeight = GetNParamUsingParametricQuery(record, "ThreadedRodWeight", pParameters)
                dSelfLockingNutWeight = GetNParamUsingParametricQuery(record, "SelfLockingNutWeight", pParameters)
           End If
        End If
    End If

    Weight = dFlexyWeight + dPlateWeight + 4 * dThreadedRodWeight + 8 * dSelfLockingNutWeight
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
    Const iNumInputs As Integer = 22    '18
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    'outputs
    Const iNumOutputs As Integer = 10
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

    sInputName(16) = "RodDia"
    sDataType(16) = "Double"

    sInputName(17) = "RodOffset"
    sDataType(17) = "Double"

    sInputName(18) = "ConnectToSteel"
    sDataType(18) = "Long"

    ' newcode
    sInputName(19) = "SteelDepth"
    sDataType(19) = "Double"

    sInputName(20) = "SteelWidth"
    sDataType(20) = "Double"

    sInputName(21) = "WebThick"
    sDataType(21) = "Double"

    sInputName(22) = "FlangeThick"
    sDataType(22) = "Double"


    'Outputs
    sOutputName(1) = "StructureInt"
    sOutputName(2) = "StructureMid"
    sOutputName(3) = "StructureExt"
    sOutputName(4) = "PlateFront"
    sOutputName(5) = "girderATT"
    sOutputName(6) = "PlateBack"
    sOutputName(7) = "Rod1"
    sOutputName(8) = "Rod2"
    sOutputName(9) = "Rod3"
    sOutputName(10) = "Rod4"

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

    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly oICH

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = oSupportComp       ' The part occurence
    oPartOcc.GetPart oPart
    Dim sPartNumber As String
    sPartNumber = oPart.PartNumber

    Set oSupportComp = Nothing
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
    Dim record As String
    Dim pParameters(0 To 2) As Parameters
    Dim dRodDia As Double
    Dim dRodOffset As Double
    Dim dRodL As Double
    Dim lConnectToSteel As Long
    Dim sErrMsg As String

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
    dRodDia = arrayOfInputs(16)
    dRodOffset = arrayOfInputs(17)
    lConnectToSteel = arrayOfInputs(18)

    If oICH Is Nothing Then          ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275          ' 6" pipe dia
        dRodL = 2 * dPlateThick + 2 * 0.01 + 0.02
    Else
        Hilti_InitializeMyHH oICH
        ' Changed to use hshlpr instead of hs_hgrhlpr

        ' Steel Dimensions
        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dSteelDepth = SupportStruct.dDepth
            dSteelWidth = SupportStruct.dWidth
            dWebThick = SupportStruct.dWebThickness
            dFlangeThick = SupportStruct.dFlangeThickness
        Else
            dSteelDepth = arrayOfInputs(19)
            dSteelWidth = arrayOfInputs(20)
            dWebThick = arrayOfInputs(21)
            dFlangeThick = arrayOfInputs(22)
        End If

        Set oStruct = Nothing
        
        Hilti_DestroyMyHH oICH

        If lConnectToSteel = 1 Then
            'calculate the length of four rods.
            dRodL = (2 * dPlateThick) + dSteelWidth + (2 * 0.01) + 0.02

            If dSteelDepth > 0.3 Or dSteelDepth < 0.075 Then
                PF_EventHandler "Steel Depth must be between 75 to 300 mm.", Err, MODULE, METHOD, True
            Else
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelDepth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelDepth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_FL_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                dPlateWidth = GetNParamUsingParametricQuery(record, "PlateWidth", pParameters)
                dPlateDepth = GetNParamUsingParametricQuery(record, "PlateDepth", pParameters)
                dNozzelWidth = GetNParamUsingParametricQuery(record, "NozzelWidth", pParameters)
                dNozzelDepth = GetNParamUsingParametricQuery(record, "NozzelDepth", pParameters)
            End If
        Else
            'calculate the length of four rods.
            dRodL = (2 * dPlateThick) + dSteelDepth + (2 * 0.01) + 0.02

            If dSteelWidth > 0.3 Or dSteelWidth < 0.075 Then
                PF_EventHandler "Steel Depth must be between 75 to 300 mm.", Err, MODULE, METHOD, True
            Else
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelWidth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelWidth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_FL_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                dPlateWidth = GetNParamUsingParametricQuery(record, "PlateWidth", pParameters)
                dPlateDepth = GetNParamUsingParametricQuery(record, "PlateDepth", pParameters)
                dNozzelWidth = GetNParamUsingParametricQuery(record, "NozzelWidth", pParameters)
                dNozzelDepth = GetNParamUsingParametricQuery(record, "NozzelDepth", pParameters)
            End If
        End If

        Set oICH = Nothing
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

    AddBox dPlateThick, dPlateWidth, dPlateDepth, Loc(0, -dPlateWidth / 2, -dPlateDepth / 2#), m_outputColl, "PlateFront"

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

    If lConnectToSteel = 1 Then
        AddBox dPlateThick, dPlateWidth, dPlateDepth, Loc(-dSteelWidth - dPlateThick, -dPlateWidth / 2, -dPlateDepth / 2#), m_outputColl, "PlateBack"
        AddCylinder dRodL, dRodDia / 2, Loc(dRodOffset / 2, dSteelDepth / 2 + dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod1"
        AddCylinder dRodL, dRodDia / 2, Loc(dRodOffset / 2, -dSteelDepth / 2 - dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod2"
        AddCylinder dRodL, dRodDia / 2, Loc(-dRodOffset / 2, dSteelDepth / 2 + dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod3"
        AddCylinder dRodL, dRodDia / 2, Loc(-dRodOffset / 2, -dSteelDepth / 2 - dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod4"
    Else
        AddBox dPlateThick, dPlateWidth, dPlateDepth, Loc(-dSteelDepth - dPlateThick, -dPlateWidth / 2, -dPlateDepth / 2#), m_outputColl, "PlateBack"
        AddCylinder dRodL, dRodDia / 2, Loc(dRodOffset / 2, dSteelWidth / 2 + dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod1"
        AddCylinder dRodL, dRodDia / 2, Loc(dRodOffset / 2, -dSteelWidth / 2 - dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod2"
        AddCylinder dRodL, dRodDia / 2, Loc(-dRodOffset / 2, dSteelWidth / 2 + dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod3"
        AddCylinder dRodL, dRodDia / 2, Loc(-dRodOffset / 2, -dSteelWidth / 2 - dRodDia / 2, -dRodL + dPlateThick + 0.02) + " * ROTY(90)", m_outputColl, "Rod4"
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
    Set oPartOcc = pSupportComp      ' The part occurence
    oPartOcc.GetPart oPart           ' The associated catalog part
    
    Dim oSupportComp As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oSupportComp = pSupportComp
    oSupportComp.GetOccAssembly oICH
    Dim dPlateWeight As Double
    Dim dNozzelWeight As Double
    Dim dFlexyWeight As Double
    Dim dThreadedRodWeight As Double
    Dim dSelfLockingNutWeight As Double
    Dim dPipeDia As Double
    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dWebThick As Double
    Dim dFlangeThick As Double
    
    Dim sGroupName As String
    Dim sPlateItemNo As String
    Dim sNozzleItemNo As String
    Dim sFlexyItemNo As Double
    Dim sThreadedRodItemNo As Double
    Dim sSelfLockingNutItemNo As Double
    Dim sPlateDesc As String
    Dim sNozzleDesc As String
    Dim sThreadedRodDesc As String
    Dim sSelfLockingNutDesc As String
    Dim lConnectionSize As Long
    Dim lConnectToSteel As Long
    Dim record As String
    Dim pParameters(0 To 2) As Parameters
    Dim pParameters1(0) As Parameters
    
    lConnectionSize = GetAttributeFromObject(oPart, "ConnectionSize")
    lConnectToSteel = GetAttributeFromObject(pSupportComp, "ConnectToSteel")
    sGroupName = GetAttributeFromObject(pSupportComp, "Group")
    
    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275 ' 6" pipe dia
    Else
        Hilti_InitializeMyHH oICH
        ' Changed to use hshlpr instead of hs_hgrhlpr

        ' Steel Dimensions
        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dSteelDepth = SupportStruct.dDepth
            dSteelWidth = SupportStruct.dWidth
            dWebThick = SupportStruct.dWebThickness
            dFlangeThick = SupportStruct.dFlangeThickness
        Else
            dSteelDepth = GetAttributeFromObject(pSupportComp, "SteelDepth")
            dSteelWidth = GetAttributeFromObject(pSupportComp, "SteelWidth")
            dWebThick = GetAttributeFromObject(pSupportComp, "WebThick")
            dFlangeThick = GetAttributeFromObject(pSupportComp, "FlangeThick")
        End If

        Set oStruct = Nothing
        
        Hilti_DestroyMyHH oICH
      
        If lConnectToSteel = 1 Then
            If dSteelDepth <= 0.3 And dSteelDepth >= 0.075 Then
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelDepth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelDepth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_Web_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                sPlateItemNo = GetSParamUsingParametricQuery(record, "PlateItemNo", pParameters)
                sNozzleItemNo = GetSParamUsingParametricQuery(record, "NozzelItemNo", pParameters)
                sThreadedRodItemNo = GetNParamUsingParametricQuery(record, "ThreadedRodItemNo", pParameters)
                sSelfLockingNutItemNo = GetNParamUsingParametricQuery(record, "SelfLockingNutItemNo", pParameters)

                pParameters1(0) = FillParameterForBSTR("Item_Nb", sPlateItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sPlateDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sNozzleItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sNozzleDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sThreadedRodItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sThreadedRodDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sSelfLockingNutItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sSelfLockingNutDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
            End If
        Else
            If dSteelWidth <= 0.3 And dSteelWidth >= 0.075 Then
                pParameters(0) = FillParameterForDouble("BeamWidthFrom", Str(dSteelWidth + 0.0001))
                pParameters(1) = FillParameterForDouble("BeamWidthTo", Str(dSteelWidth - 0.0001))
                pParameters(2) = FillParameterForBSTR("ConnectionSize", Str(lConnectionSize))
                record = ReadParametricData("JUAMIFlexyPack_Web_AUX", "where BeamWidthFrom <= ? and BeamWidthTo >= ? and ConnectionSize = ?")
                sPlateItemNo = GetSParamUsingParametricQuery(record, "PlateItemNo", pParameters)
                sNozzleItemNo = GetSParamUsingParametricQuery(record, "NozzelItemNo", pParameters)
                sThreadedRodItemNo = GetNParamUsingParametricQuery(record, "ThreadedRodItemNo", pParameters)
                sSelfLockingNutItemNo = GetNParamUsingParametricQuery(record, "SelfLockingNutItemNo", pParameters)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sPlateItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sPlateDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sNozzleItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sNozzleDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sThreadedRodItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sThreadedRodDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
                
                pParameters1(0) = FillParameterForBSTR("Item_Nb", sSelfLockingNutItemNo)
                record = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb = ?")
                sSelfLockingNutDesc = GetSParamUsingParametricQuery(record, "Part_Description", pParameters1)
            End If
        End If
        
        oICH.SetAttributeValue "ItemNo", oSupportComp, sPlateItemNo & ", " & sNozzleItemNo
        Hilti_DestroyMyHH oICH
    End If
    
    bstrBOMDesc = "Hilti, " & sPlateItemNo & ", " & Replace(sPlateDesc, "Hilti ", "") _
                        & " + 1, " & sPlateItemNo & ", " & Replace(sPlateDesc, "Hilti ", "") _
                        & " + 1, " & sNozzleItemNo & ", " & Replace(sNozzleDesc, "Hilti ", "") _
                        & " + 4, " & sThreadedRodItemNo & ", " & Replace(sThreadedRodDesc, "Hilti ", "") _
                        & " + 8, " & sSelfLockingNutItemNo & ", " & Replace(sSelfLockingNutDesc, "Hilti ", "")
    
    Set oSupportComp = Nothing
    Set oPartOcc = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub




