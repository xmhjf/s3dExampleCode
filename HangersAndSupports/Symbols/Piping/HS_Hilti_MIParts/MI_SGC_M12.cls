VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_SGC_M12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Anvil_FIG95.cls
'   ProgID:         HS_Hilti_MIParts.MI_SGC_M12
'   Author:         JOGN
'   Creation Date:  September 20, 2007

'   Description:
'       Crossbeam connector to Steel
'
'   Change History:
'   Date            who         change description
'   16-02-2010      VSP          Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   22/03/13        SK          DM 231005  Hilti Beam Clamp gives warning when it should not
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "MI_SGC_M12"  'Used for error messages
Private Const m_progID = "HS_Hilti_MIParts.MI_SGC_M12"

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

    ' Inputs
    Const iNumInputs As Integer = 13
    Dim sInputName(iNumInputs) As String
    Dim sDataType(iNumInputs) As String

    ' Outputs
    Const iNumOutputs As Integer = 4
    Dim sOutputName(iNumOutputs) As String

    'Inputs
    sInputName(2) = "ItemNo"
    sDataType(2) = "String"

    sInputName(3) = "Finish"
    sDataType(3) = "Long"

    sInputName(4) = "Group"
    sDataType(4) = "String"

    sInputName(5) = "BoltSize"
    sDataType(5) = "Double"

    sInputName(6) = "Bolt_Length"
    sDataType(6) = "Double"

    sInputName(7) = "ClampWidth"
    sDataType(7) = "Double"

    sInputName(8) = "Clamp_Depth"
    sDataType(8) = "Double"

    sInputName(9) = "ClampHeight"
    sDataType(9) = "Double"

    sInputName(10) = "SpacerDepth"
    sDataType(10) = "Double"

    sInputName(11) = "SpacerWidth"
    sDataType(11) = "Double"

    sInputName(12) = "PortOffset"
    sDataType(12) = "Double"
    
    sInputName(13) = "FlangeThick"
    sDataType(13) = "Double"

    sOutputName(1) = "Structure"
    sOutputName(2) = "Clamp"
    sOutputName(3) = "Spacer"
    sOutputName(4) = "Bolt"

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
    Set oSupportComp = Nothing

    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dFlangeThick As Double
    Dim dWebThick As Double
    Dim dPipeDia As Double

    Dim sItemNo As String
    Dim lFinish As Long
    Dim sGroup As String
    Dim dBoltSize As Double
    Dim dBolt_Length As Double
    Dim dClampWidth As Double
    Dim dClamp_Depth As Double
    Dim dClampHeight As Double
    Dim dPortOffset As Double
    Dim Spacer_Desc As String
    Dim Spacer_Height As Double
    Dim Spacer_Weight As Double
    Dim dSpacerDepth As Double
    Dim dSpacerWidth As Double

    sItemNo = arrayOfInputs(2)
    lFinish = arrayOfInputs(3)
    sGroup = arrayOfInputs(4)
    dBoltSize = arrayOfInputs(5)
    dBolt_Length = arrayOfInputs(6)
    dClampWidth = arrayOfInputs(7)
    dClamp_Depth = arrayOfInputs(8)
    dClampHeight = arrayOfInputs(9)
    dSpacerDepth = arrayOfInputs(10)
    dSpacerWidth = arrayOfInputs(11)
    dPortOffset = arrayOfInputs(12)

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275    ' 6" pipe dia
    Else
        Hilti_InitializeMyHH oICH
        'Changed to work with hshlpr instead of hs_hgrhlpr

        ' Steel Dimensions
        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dFlangeThick = SupportStruct.dFlangeThickness
        Else
            dFlangeThick = arrayOfInputs(13)
        End If
        
        Set oStruct = Nothing
        
        Hilti_DestroyMyHH oICH


        If dFlangeThick < 0.003 Or dFlangeThick > 0.036 Then
            PF_EventHandler "Flange thickness must be between 3 to 36 mm", Err, MODULE, METHOD, True
        End If

        If dFlangeThick < 0.003 Or dFlangeThick > 0.036 Then
            PF_EventHandler "Must use spacers if flange thicknesses lies outside the range of 3 mm to 36 mm", Err, MODULE, METHOD, True
        End If

        If dFlangeThick > 0.003 Or dFlangeThick < 0.036 Then
            Spacer_Desc = " "
            Spacer_Height = 0
            Spacer_Height = 0
            Spacer_Weight = 0
        Else
            If dFlangeThick <= 0.02 Then
                Spacer_Desc = " + MI-GCD-6MM 283590 0.03 kg"
                Spacer_Height = 0.006
                Spacer_Weight = 0.03
            Else
                Spacer_Desc = " + MI-GCD-12MM 285591 0.07 kg"
                Spacer_Height = 0.012
                Spacer_Weight = 0.07
            End If
        End If
        Set oICH = Nothing
    End If

    AddPort "Structure", 0, 0, 0, _
            0, 0, -1, _
            -1, 0, 0, m_outputColl, "Structure", arrayOfInputs(1), "Structure"

    AddBox dClamp_Depth, dClampWidth, dFlangeThick + dClampHeight - Spacer_Height, Loc(-dClamp_Depth / 2, -dClampWidth / 2, Spacer_Height), m_outputColl, "Clamp"

    If Spacer_Height > 0 Then
        AddBox dSpacerDepth, dSpacerWidth, Spacer_Height, Loc(-dSpacerDepth / 2, -dSpacerWidth / 2, 0), m_outputColl, "Spacer"
    End If
    AddCylinder dBolt_Length, dBoltSize / 2, Loc(0, 0, -dBolt_Length + dFlangeThick + dClampHeight + 0.012), m_outputColl, "Bolt"

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

    Dim sItemNo As Long

    sItemNo = GetAttributeFromObject(pSupportComp, "ItemNo")
    sPartNumber = oPart.PartNumber

    bstrBOMDesc = Hilti_BuildBom(pSupportComp)

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
