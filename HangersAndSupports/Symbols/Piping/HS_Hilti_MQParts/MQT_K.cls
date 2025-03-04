VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MQT_K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   MQT_K.cls
'   ProgID:         HS_Hilti_MQParts.MQT_K
'   Author:         JRK
'   Creation Date:  26.Nov.2007
'   Description: Channel
'
'   Change History:
'   Date              who            change description
'   16-02-2010        VSP            Addition to S3D
'   19-08-2010        YLP            TR-CP-186061  Failed to place some of the Hilti_MQ parts
'   08/02/2011        Ramya          TR # 183457  Length and Weight properties are incorrect for a Hilti Part in a Specific case
'   02/09/11          Ramya          TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014        Chethan        DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'   12-05-2016        PVK            TR-CP-294544	Errors when placing HS_Hilti_MQParts using 3rd Party App
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
'Create Constant Used for error messages
Private Const MODULE = "MQT_K"
Private Const m_progID = "HS_Hilti_MQParts.MQT_K"

Private Const INPUT_1 = "Part"
Private Const INPUT_2 = "CardinalPoint"
Private Const INPUT_3 = "Length"
Private Const INPUT_4 = "MaterialGrade"
Private Const INPUT_5 = "Orientation"
Private Const INPUT_6 = "BeginOverLength"
Private Const INPUT_7 = "EndOverLength"
Private Const INPUT_8 = "BeginMiter"
Private Const INPUT_9 = "EndMiter"
Private Const INPUT_10 = "L"
Private Const INPUT_11 = "W"
Private Const INPUT_12 = "ItemNo"
Private Const INPUT_13 = "Group"
Private Const INPUT_14 = "ChannelSize"
Private Const INPUT_15 = "H"
Private Const INPUT_16 = "T"
Private Const INPUT_17 = "BoltSP"
Private Const INPUT_18 = "D"
Private Const INPUT_19 = "SteelDepth"
Private Const INPUT_20 = "FlangeThick"

Private Const OUTPUT_BC = "BeginCap"
Private Const OUTPUT_EC = "EndCap"
Private Const OUTPUT_NT = "Neutral"
Private Const OUTPUT_BC_SUR = "BeginCapSurface"
Private Const OUTPUT_EC_SUR = "EndCapSurface"
Private Const STRUCTUREINT = "StructureInt"
Private Const STRUCTUREEXT = "StructureExt"
Private Const OUTPUT_Plate1 = "Plate1"
Private Const OUTPUT_Plate2 = "Plate2"
Private Const OUTPUT_Bolt1 = "Bolt1"
Private Const OUTPUT_Bolt2 = "Bolt2"
Private Const OUTPUT_Bolt3 = "Bolt3"
Private Const OUTPUT_Bolt4 = "Bolt4"

Private Const NUM_INPUTS = 20
Private Const NUM_PORTS = 30
Private Const NUM_OUTPUTS = NUM_PORTS + 13

Implements IJHgrSymbolWCGServices
Implements IJDUserSymbolServices
Implements IJHgrSymbolDwgServices
Implements IJHgrSymbolBOMServices

'********************************************************************
' Function: IJDUserSymbolServices_EditOccurence
' This method is called each time the user wants to edit the symbol
' occurrence (Modify its parameter values or its representation for
' instance). Prefer to register a custom command to fulfill this
' functionality, into the symbol definition.
' Currently obsolete (as of Oct 99).
'
' Parameters: pSymbolOccurence As Object
'             TransactionManager As Object
' Returns: Boolean
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

'********************************************************************
' Function: IJDUserSymbolServices_GetDefinitionName
' Return the Definition Name
'
' Parameters: definitionParamaters As Variant
' Returns: String
'********************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_progID
End Function

'********************************************************************
' Function: IJDUserSymbolServices_InstanciateDefinition
' Instantiates a persistent symbol definition object and initializes
' it for the first time.
'
' Parameters: CodeBase As String
'             defParamaters As Variant
'             ActiveConnection As Object
' Returns: Object
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_progID
    oSymbolDefinition.CodeBase = CodeBase

    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'********************************************************************
' Sub: IJDUserSymbolServices_InitializeSymbolDefinition
' Cleans the previous definition up and initializes the new one
' (Input, Output, Representation, RepresenationEvaluation,…).
'
' Parameters: pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler

    'Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    'DEFINE INPUTS FOR ALL REPRESENTATIONS
    Dim SymbolInput(1 To NUM_INPUTS) As IMSSymbolEntities.IJDInput

    Dim ii As Integer
    For ii = 1 To NUM_INPUTS
        Set SymbolInput(ii) = New IMSSymbolEntities.DInput
    Next

    'Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent  'not persistent PC

    Dim PC1 As IMSSymbolEntities.IJDParameterContent
    Set PC1 = New IMSSymbolEntities.DParameterContent    'not persistent PC
    PC1.Type = igString
    PC.Type = igValue

    'Following indented lines will be removed very soon. This has been
    'included to improve the performance. These lines will be removed as soon as
    'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.
    'set the CMCacheForPart method in the definition
    Dim libDesc As New DLibraryDescription
    Dim mCookie As Long
    Dim libCookie As Long

    libDesc.Name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = "HgrSupportSymbolUtilities.CustomMethods"

    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc

    'Get the lib/method cookie
    libCookie = libDesc.Cookie
    mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForHgrPart", libCookie)

    'There should be one input defined for refData part.
    'This should be the first input
    SymbolInput(1).Name = INPUT_1
    SymbolInput(1).Description = "Part"
    SymbolInput(1).IJDInputStdCustomMethod.SetCMCache libCookie, mCookie

    SymbolInput(2).Name = INPUT_2
    SymbolInput(2).Description = "Cardinality Point of the Beam"
    SymbolInput(2).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 1
    SymbolInput(2).DefaultParameterValue = PC

    SymbolInput(3).Name = INPUT_3
    SymbolInput(3).Description = "Length"        'may be changed later by constraints
    SymbolInput(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.5
    SymbolInput(3).DefaultParameterValue = PC

    SymbolInput(4).Name = INPUT_4
    SymbolInput(4).Description = "MaterialGrade of the Beam"    'codelist
    SymbolInput(4).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "A36"
    SymbolInput(4).DefaultParameterValue = PC1

    SymbolInput(5).Name = INPUT_5
    SymbolInput(5).Description = "Orientation"               'codelist
    SymbolInput(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 10#
    SymbolInput(5).DefaultParameterValue = PC

    SymbolInput(6).Name = INPUT_6
    SymbolInput(6).Description = "Begin OverLength of the Beam"    'positive value always elongate the Beam
    SymbolInput(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(6).DefaultParameterValue = PC

    SymbolInput(7).Name = INPUT_7
    SymbolInput(7).Description = "End OverLength of the Beam"    'positive value always elongate the beam
    SymbolInput(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(7).DefaultParameterValue = PC

    SymbolInput(8).Name = INPUT_8
    SymbolInput(8).Description = "BeginMiter"        'positive value always elongate the beam
    SymbolInput(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(8).DefaultParameterValue = PC

    SymbolInput(9).Name = INPUT_9
    SymbolInput(9).Description = "EndMiter"          'positive value always elongate the beam
    SymbolInput(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(9).DefaultParameterValue = PC

    SymbolInput(10).Name = INPUT_10
    SymbolInput(10).Description = "L"               'codelist
    SymbolInput(10).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.015
    SymbolInput(10).DefaultParameterValue = PC

    SymbolInput(11).Name = INPUT_11
    SymbolInput(11).Description = "W"          'positive value always elongate the beam
    SymbolInput(11).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.06
    SymbolInput(11).DefaultParameterValue = PC

    SymbolInput(12).Name = INPUT_12
    SymbolInput(12).Description = "ItemNo"
    SymbolInput(12).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "284241"
    SymbolInput(12).DefaultParameterValue = PC1

    SymbolInput(13).Name = INPUT_13
    SymbolInput(13).Description = "Group"       'codelist
    SymbolInput(13).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "Bracket"
    SymbolInput(13).DefaultParameterValue = PC1

    SymbolInput(14).Name = INPUT_14
    SymbolInput(14).Description = "ChannelSize"    'codelist
    SymbolInput(14).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "MQ-41"
    SymbolInput(14).DefaultParameterValue = PC1

    SymbolInput(15).Name = INPUT_15
    SymbolInput(15).Description = "H"          'codelist
    SymbolInput(15).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.1
    SymbolInput(15).DefaultParameterValue = PC

    SymbolInput(16).Name = INPUT_16
    SymbolInput(16).Description = "T"          'codelist
    SymbolInput(16).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.005
    SymbolInput(16).DefaultParameterValue = PC

    SymbolInput(17).Name = INPUT_17
    SymbolInput(17).Description = "BoltSP"      'codelist
    SymbolInput(17).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.08
    SymbolInput(17).DefaultParameterValue = PC

    SymbolInput(18).Name = INPUT_18
    SymbolInput(18).Description = "D"           'codelist
    SymbolInput(18).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.008
    SymbolInput(18).DefaultParameterValue = PC

    SymbolInput(19).Name = INPUT_19
    SymbolInput(19).Description = "SteelDepth"
    SymbolInput(19).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.25
    SymbolInput(19).DefaultParameterValue = PC

    SymbolInput(20).Name = INPUT_20
    SymbolInput(20).Description = "FlangeThick"
    SymbolInput(20).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.015
    SymbolInput(20).DefaultParameterValue = PC

    'Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition

    For ii = 1 To NUM_INPUTS
        oInputs.SetInput SymbolInput(ii), ii
        Set SymbolInput(ii) = Nothing
    Next

    'DEFINE OUTPUTS FOR Physical REPRESENTATION
    'Create the output
    Dim SymbolOutput(1 To NUM_OUTPUTS) As IMSSymbolEntities.IJDOutput
    Dim OutputName As String

    Set SymbolOutput(1) = New IMSSymbolEntities.DOutput
    SymbolOutput(1).Name = OUTPUT_BC
    SymbolOutput(1).Description = OUTPUT_BC
    SymbolOutput(1).Properties = 0

    Set SymbolOutput(2) = New IMSSymbolEntities.DOutput
    SymbolOutput(2).Name = OUTPUT_EC
    SymbolOutput(2).Description = OUTPUT_EC
    SymbolOutput(2).Properties = 0

    Set SymbolOutput(3) = New IMSSymbolEntities.DOutput
    SymbolOutput(3).Name = OUTPUT_NT
    SymbolOutput(3).Description = OUTPUT_NT
    SymbolOutput(3).Properties = 0

    Set SymbolOutput(4) = New IMSSymbolEntities.DOutput
    SymbolOutput(4).Name = OUTPUT_BC_SUR
    SymbolOutput(4).Description = OUTPUT_BC_SUR
    SymbolOutput(4).Properties = 0

    Set SymbolOutput(5) = New IMSSymbolEntities.DOutput
    SymbolOutput(5).Name = OUTPUT_EC_SUR
    SymbolOutput(5).Description = OUTPUT_EC_SUR
    SymbolOutput(5).Properties = 0

    Set SymbolOutput(6) = New IMSSymbolEntities.DOutput
    SymbolOutput(6).Name = STRUCTUREINT
    SymbolOutput(6).Description = STRUCTUREINT
    SymbolOutput(6).Properties = 0

    Set SymbolOutput(7) = New IMSSymbolEntities.DOutput
    SymbolOutput(7).Name = STRUCTUREEXT
    SymbolOutput(7).Description = STRUCTUREEXT
    SymbolOutput(7).Properties = 0

    Set SymbolOutput(8) = New IMSSymbolEntities.DOutput
    SymbolOutput(8).Name = OUTPUT_Plate1
    SymbolOutput(8).Description = OUTPUT_Plate1
    SymbolOutput(8).Properties = 0

    Set SymbolOutput(9) = New IMSSymbolEntities.DOutput
    SymbolOutput(9).Name = OUTPUT_Plate2
    SymbolOutput(9).Description = OUTPUT_Plate2
    SymbolOutput(9).Properties = 0

    Set SymbolOutput(10) = New IMSSymbolEntities.DOutput
    SymbolOutput(10).Name = OUTPUT_Bolt1
    SymbolOutput(10).Description = OUTPUT_Bolt1
    SymbolOutput(10).Properties = 0

    Set SymbolOutput(11) = New IMSSymbolEntities.DOutput
    SymbolOutput(11).Name = OUTPUT_Bolt2
    SymbolOutput(11).Description = OUTPUT_Bolt2
    SymbolOutput(11).Properties = 0

    Set SymbolOutput(12) = New IMSSymbolEntities.DOutput
    SymbolOutput(12).Name = OUTPUT_Bolt3
    SymbolOutput(12).Description = OUTPUT_Bolt3
    SymbolOutput(12).Properties = 0

    Set SymbolOutput(13) = New IMSSymbolEntities.DOutput
    SymbolOutput(13).Name = OUTPUT_Bolt4
    SymbolOutput(13).Description = OUTPUT_Bolt4
    SymbolOutput(13).Properties = 0

    For ii = 1 To NUM_PORTS
        Set SymbolOutput(ii + 13) = New IMSSymbolEntities.DOutput
        OutputName = "Port" + CStr(ii)
        SymbolOutput(ii + 13).Name = OutputName
        SymbolOutput(ii + 13).Description = OutputName
        SymbolOutput(ii + 13).Properties = 0
    Next

    'DEFINE THE "SimplePhysical" REPRESENTATION
    Dim oRep1 As IMSSymbolEntities.IJDRepresentation
    Set oRep1 = New IMSSymbolEntities.DRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep1

    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.

    oRep1.Name = "Symbolic"
    oRep1.Description = "Symbolic Represntation of the 3d flexible"
    oRep1.RepresentationId = SimplePhysical    'define a aspect 0 (Simple_physical)
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    'Set the output
    For ii = 1 To NUM_OUTPUTS
        oOutputs.SetOutput SymbolOutput(ii)
    Next

    Dim libDesc1 As New DLibraryDescription

    libDesc1.Name = "mySelfAsLib1"
    libDesc1.Type = imsLIBRARY_IS_ACTIVEX
    libDesc1.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc1.Source = m_progID

    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc1
    libCookie = libDesc1.Cookie

    Dim mthCookie As Long
    mthCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("Symbolic", libCookie)
    oRep1.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mthCookie

    'Set the representation to definition
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    oReps.SetRepresentation oRep1

    Dim oVbFuncSymbolicRep As IJDRepresentationEvaluation

    Dim oScripts As IMSSymbolEntities.IJDRepresentationEvaluations
    Set oScripts = pSymbolDefinition

    'DEFINE DetailPhysical REPRESENTATION WITH DIFFERENT OUTPUTS
    'We are reusing the previously allocated IMSSymbolEntities.DRepresentation
    'used to describe the Simple Physical Representation.  Clear outputs first.
    oOutputs.RemoveAllOutput

    'Create the output
    Dim DetailSymbolOutput As IMSSymbolEntities.IJDOutput

    Set DetailSymbolOutput = New IMSSymbolEntities.DOutput

    For ii = 1 To NUM_OUTPUTS
        oOutputs.SetOutput SymbolOutput(ii)
    Next

    oRep1.Name = "Detailed"
    oRep1.Description = "Detailed Represntation of the 3d flexible"
    oRep1.RepresentationId = DetailPhysical          'Detailed Physical
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    mthCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("Detailed", libCookie)

    'set the customer evaluation method
    oRep1.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mthCookie

    'Set the representation to definition
    oReps.SetRepresentation oRep1

    'Set the script associated to the Detailed representation
    Dim oVbFuncDetailedRep As DRepresentationEvaluation

    'DEFINE Maintenance Envelop REPRESENTATION WITH DIFFERENT OUTPUTS
    oRep1.Name = "Maintenance"
    oRep1.Description = "Maintenance Represntation of the 3d flexible"
    oRep1.RepresentationId = Maintenance     'Maintenance
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    'Set the customer evaluation method
    oRep1.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mthCookie

    'Set the representation to definition
    oReps.SetRepresentation oRep1

    'Set the script associated to the Maintenance representation
    Dim oVbFuncMaintenanceRep As DRepresentationEvaluation

    Set oReps = Nothing
    Set oRep1 = Nothing
    Set oScripts = Nothing
    Set oVbFuncSymbolicRep = Nothing
    Set oVbFuncDetailedRep = Nothing
    Set oVbFuncMaintenanceRep = Nothing

    'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
    'WHICH ARE GRAPHIC ENTITIES.

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'********************************************************************
' Sub: IJDUserSymbolServices_InvokeRepresentation
'
' Parameters: sblOcc As Object
'             repName As String
'             outputcoll As Object
'             arrayOfInputs()
'********************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())

End Sub

'********************************************************************
' Sub: Symbolic()
' This procedure will assist with the creation of the Symbolic
' representation of the profile part.  The symbol method creates
' the HgrPorts that will be used for assembly purpose.
'
' Parameters: pRepSCM As IJDRepresentationStdCustomMethod
'********************************************************************
Public Sub Symbolic(ByVal pRepSCM As IJDRepresentationStdCustomMethod)
    Call PhysicalRepresentation(pRepSCM, SymbolicRepresentation)
End Sub

'********************************************************************
' Sub: Detailed()
' This procedure will assist with the creation of the detailed
' representation of the profiled part
'
' Parameters: pRepSCM As IJDRepresentationStdCustomMethod
'********************************************************************
Public Sub Detailed(ByVal pRepSCM As IJDRepresentationStdCustomMethod)
        'Call PhysicalRepresentation(pRepSCM, DetailedRepresentation)
End Sub

'********************************************************************
' Sub: PhysicalRepresentation()
' This procedure will be used to create teh physical representation
' of the profiled part
'
' Parameters: pRepSCM As IJDRepresentationStdCustomMethod
'             RepType As RepresentationType
'********************************************************************
Private Sub PhysicalRepresentation(ByVal pRepSCM As IJDRepresentationStdCustomMethod, RepType As RepresentationType)
    Const METHOD = "PhysicalRepresentation"
    On Error GoTo ErrorHandler
    Dim oGeomFactory As New GeometryFactory
    Dim pRepDG As IJDRepresentationDuringGame
    Dim pDefinition As IJDSymbolDefinition
    Dim Inputs As IJDInputs
    Dim m_outputColl As IJDOutputCollection
    Dim oCrossSectionColl As IJDTargetObjectCol
    Dim oCrossSectionServices As CrossSectionServices
    Dim oHgrSymbolPortHlpr As IJHgrSymbolPortHlpr
    Dim oPart As Object
    Dim oTempPart As IJDPart
    Dim CSName As String
    Dim CSStandard As String
    Dim SectionType As String
    Dim LENGTH As Double
    Dim WIDTH As Double
    Dim Thickness As Double
    Dim BeginOverLength As Double
    Dim EndOverLength As Double
    Dim dLength As Double
    Dim cardinalX As Double
    Dim cardinalY As Double
    Dim oPorts As IJElements
    Dim ii As Integer
    Dim CardinalPoint As Long
    Dim oCrossSection As Object
    Dim varWidth As Variant
    Dim dblWidth As Double
    Dim varDepth As Variant
    Dim dblDepth As Double
    Dim lPortCount As Integer
    Dim oBeginCapSurface As Object
    Dim oEndCapSurface As Object
    Dim dL As Double
    Dim dW As Double
    Dim dH As Double
    Dim dT As Double
    Dim dBoltSP As Double
    Dim dD As Double
    Dim dFlangeD As Double
    Dim dFlangeT As Double
    Dim dBoltL As Double

    Dim oCatServices As New GSCADHgrSupCatalogServices.CatalogServices

    Set pRepDG = pRepSCM
    Set pDefinition = pRepDG.Definition
    Set Inputs = pDefinition.IJDInputs
    Set m_outputColl = pRepDG.OutputCollection

    Set oHgrSymbolPortHlpr = New HgrSymbolPortHlpr
    Set oCrossSectionServices = New CrossSectionServices

    Set oPart = Inputs.GetInputAtIndex(1).IJDInputDuringGame.Result
    Set oTempPart = oPart
    CardinalPoint = Inputs.GetInputAtIndex(2).IJDInputDuringGame.Result.UomValue
    BeginOverLength = Inputs.GetInputAtIndex(6).IJDInputDuringGame.Result.UomValue
    EndOverLength = Inputs.GetInputAtIndex(7).IJDInputDuringGame.Result.UomValue
    dL = Inputs.GetInputAtIndex(10).IJDInputDuringGame.Result.UomValue
    dW = Inputs.GetInputAtIndex(11).IJDInputDuringGame.Result.UomValue
    dH = Inputs.GetInputByIndex(15).IJDInputDuringGame.Result.UomValue
    dT = Inputs.GetInputByIndex(16).IJDInputDuringGame.Result.UomValue
    dBoltSP = Inputs.GetInputByIndex(17).IJDInputDuringGame.Result.UomValue
    dD = Inputs.GetInputByIndex(18).IJDInputDuringGame.Result.UomValue
    dLength = Inputs.GetInputByIndex(3).IJDInputDuringGame.Result.UomValue
    LENGTH = dLength

    Dim oSupComp As IJHgrSupportComponent
    Dim oICH As IJHgrInputConfigHlpr
    Set oSupComp = pRepDG.Definition.IJDDefinitionPlayerEx.PlayingSymbol
    oSupComp.GetOccAssembly oICH

    If oICH Is Nothing Then  ' The OCH will be null the first time thru until the symbol is computed the first time
        dFlangeD = 0.4
        LENGTH = dFlangeD
      Else
            Hilti_InitializeMyHH oICH
            If Not oStructCollection Is Nothing Then
                    If oStructCollection.Count > 0 Then
                        Dim SupportStruct As hsSteelMember
                        SupportStruct = GetSupportingSectionData(1)
                        dFlangeD = SupportStruct.dDepth
                        dFlangeT = SupportStruct.dFlangeThickness
                    End If
                Else
                    dFlangeD = Inputs.GetInputByIndex(19).IJDInputDuringGame.Result.UomValue
                    dFlangeT = Inputs.GetInputByIndex(20).IJDInputDuringGame.Result.UomValue
            End If

        LENGTH = dFlangeD - dFlangeT * 2 - 0.025
                
        If LENGTH > 0.4 Then
            PF_EventHandler "Steel Depth can not exceed 400 mm.", Err, MODULE, METHOD, True
        End If

        If LENGTH < 0.1 Then
            PF_EventHandler "Steel Depth can not be less than 100 mm.", Err, MODULE, METHOD, True
        End If

        Set oICH = Nothing
    End If

    Set oPorts = oCatServices.GetConnectionPortColl( _
                 m_outputColl.ResourceManager, oPart, _
                 CardinalPoint, _
                 BeginOverLength, LENGTH, EndOverLength, _
                 cardinalX, cardinalY)

    Dim OutputName As String

    lPortCount = oPorts.Count - 2

    For ii = 1 To lPortCount
        OutputName = "Port" + CStr(ii)
        m_outputColl.AddOutput OutputName, oPorts.Item(ii)
    Next

    Set oBeginCapSurface = oPorts.Item(lPortCount + 1)

    Set oEndCapSurface = oPorts.Item(lPortCount + 2)

    m_outputColl.AddOutput OUTPUT_BC_SUR, oBeginCapSurface

    m_outputColl.AddOutput OUTPUT_EC_SUR, oEndCapSurface

    oHgrSymbolPortHlpr.GetHgrAssociation oPart, HGR_CROSSSECREL, oCrossSectionColl
    Set oCrossSection = oCrossSectionColl.Item(1)

    oCrossSectionServices.GetCrossSectionAttributeValue oCrossSection, "Width", varWidth
    dblWidth = varWidth
    oCrossSectionServices.GetCrossSectionAttributeValue oCrossSection, "Depth", varDepth
    dblDepth = varDepth

    Dim oX As Double
    Dim oY As Double
    Dim oZ As Double

    Dim oPortFac As IJHgrPortFactory
    Dim oPort1 As IJHgrPort
    Dim oPort2 As IJHgrPort
    
    Dim oElements As IJElements

    Set oPortFac = New HgrPortFactory

    oX = cardinalX
    oY = cardinalY
    oZ = 0#

    Set oElements = oPortFac.CreateHgrPortsFromPart(m_outputColl.ResourceManager, oPart)

    Set oPort1 = GetPortByNameFromCollection(oElements, "BeginCap")
    oPort1.PutOrientation oX + dblWidth / 2#, oY + dblDepth / 2#, oZ - dT, _
                          0, 0, 1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_BC, oPort1

    Set oPort2 = GetPortByNameFromCollection(oElements, "EndCap")
    oPort2.PutOrientation oX + dblWidth / 2, oY + dblDepth / 2#, oZ + LENGTH + dT, _
                          0, 0, -1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_EC, oPort2

    dBoltL = 0.0125 - dT + dL + 0.0125 - dT
    AddBox dH, dW, dL, Loc(dblDepth / 2 - dH / 2, dblWidth / 2 - dW / 2, -dT - BeginOverLength), m_outputColl, "Plate1", "Plate1"
    AddBox dH, dW, dL, Loc(dblDepth / 2 - dH / 2, dblWidth / 2 - dW / 2, LENGTH - dL + dT + EndOverLength), m_outputColl, "Plate2", "Plate2"
    AddCylinder dBoltL, dD / 2, Loc(dblDepth / 2 - dBoltSP / 2, dblWidth / 2, -dT + dL / 2 - dBoltL / 2 - BeginOverLength), m_outputColl, "Bolt1", "Bolt1"
    AddCylinder dBoltL, dD / 2, Loc(dblDepth / 2 + dBoltSP / 2, dblWidth / 2, -dT + dL / 2 - dBoltL / 2 - BeginOverLength), m_outputColl, "Bolt2", "Bolt2"
    AddCylinder dBoltL, dD / 2, Loc(dblDepth / 2 - dBoltSP / 2, dblWidth / 2, -dT + dL / 2 - dBoltL / 2 + LENGTH - 0.005 + EndOverLength), m_outputColl, "Bolt3", "Bolt3"
    AddCylinder dBoltL, dD / 2, Loc(dblDepth / 2 + dBoltSP / 2, dblWidth / 2, -dT + dL / 2 - dBoltL / 2 + LENGTH - 0.005 + EndOverLength), m_outputColl, "Bolt4", "Bolt4"

    Set oPort1 = Nothing
    Set oPort2 = Nothing
    Set oPortFac = Nothing
    oPorts.Clear
    Set pRepDG = Nothing
    Set oCrossSectionColl = Nothing
    Set oCrossSectionServices = Nothing
    Set oBeginCapSurface = Nothing
    Set oEndCapSurface = Nothing
    Set oTempPart = Nothing
    Set oHgrSymbolPortHlpr = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub

'********************************************************************
' Function: IJHgrSymbolDwgServices_GetDrawingGeometry()
'
' Parameters: pSupportComp As Object
' Returns: Object
'********************************************************************
Private Function IJHgrSymbolDwgServices_GetDrawingGeometry(ByVal pSupportComp As Object) As Object
    Const METHOD = "IJHgrSymbolDwgServices_GetDrawingGeometry"
    On Error GoTo ErrorHandler
    Dim oCat As SP3DHgrSupGeometryServices.GeometryServices

    Set oCat = New GeometryServices

    Set IJHgrSymbolDwgServices_GetDrawingGeometry = oCat.GetHgrBeamGeometry(pSupportComp)

    Set oCat = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Function

'********************************************************************
' Sub: IJHgrSymbolWCGServices_EvaluateLocalWCG
' This function will evaluate the local weight and center of gravity
'
' Parameters: PSupportComp As Object
'             Weight As Double
'             CogX As Double
'             CogY As Double
'             CogZ As Double
'********************************************************************
Private Sub IJHgrSymbolWCGServices_EvaluateLocalWCG(ByVal pSupportComp As Object, Weight As Double, CogX As Double, CogY As Double, CogZ As Double)
    Const METHOD = "IJHgrSymbolWCGServices_EvaluateLocalWCG"
    On Error GoTo ErrorHandler
    Dim oPartOcc As PartOcc
    Dim oHgrPart As IJDPart
    Dim dblBeginLength As Double
    Dim dblEndLength As Double
    Dim dblLength As Double
    Dim dblTotalLength As Double
    Dim dblWeight As Double
    Dim varWeight As Variant
    Dim oHgrPortHlpr As HgrSymbolPortHlpr
    Dim oCrossSectionColl As IJDTargetObjectCol
    Dim oCrossSectionServices As CrossSectionServices
    Dim oCSOcc As Object
    Dim oCSObj As Object
    Dim oIJDObject As IJDObject
    Dim dFlangeD As Double
    Dim dFlangeT As Double
    
    Set oPartOcc = pSupportComp
    oPartOcc.GetPart oHgrPart
    
    Dim oMyPart As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr

    Set oMyPart = pSupportComp
    oMyPart.GetOccAssembly oICH
            
    dblLength = GetAttributeFromObject(pSupportComp, "Length")
    dblBeginLength = GetAttributeFromObject(pSupportComp, "BeginOverLength")
    dblEndLength = GetAttributeFromObject(pSupportComp, "EndOverLength")
        If (dblLength <= 0) Then
            dblLength = 0.4
            oICH.SetAttributeValue "Length", oMyPart, dblLength
        End If
    
    Hilti_InitializeMyHH oICH
        
    If Not oStructCollection Is Nothing Then
                If oStructCollection.Count > 0 Then
                    Dim SupportStruct As hsSteelMember
                    SupportStruct = GetSupportingSectionData(1)
                    dFlangeD = SupportStruct.dDepth
                    dFlangeT = SupportStruct.dFlangeThickness
                End If
         Else
                dFlangeD = GetAttributeFromObject(oHgrPart, "SteelDepth")
                dFlangeT = GetAttributeFromObject(oHgrPart, "FlangeThick")
    End If
                    
        dblLength = dFlangeD - dFlangeT * 2 - 0.025
        oICH.SetAttributeValue "Length", oMyPart, dblLength
        Set oMyPart = Nothing
        Set oICH = Nothing
        dblTotalLength = dblBeginLength + dblLength + dblEndLength

    Set oHgrPortHlpr = New HgrSymbolPortHlpr
    Call oHgrPortHlpr.GetHgrAssociation(oHgrPart, 3, oCrossSectionColl)

    Set oCSObj = oCrossSectionColl.Item(1)

    Set oCrossSectionServices = New CrossSectionServices

    oCrossSectionServices.PlaceCrossSectionOccurrence GetModelResourceManager(), oCSObj, True, oCSOcc

    oCrossSectionServices.GetCardinalPoint oCSOcc, 10, CogX, CogY

    CogZ = dblTotalLength / 2#

    Set oIJDObject = oCSOcc

    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "UnitWeight", varWeight

    dblWeight = varWeight

    Weight = dblWeight * dblTotalLength

    Set oCrossSectionColl = Nothing
    Set oCSOcc = Nothing
    Set oCSObj = Nothing
    Set oIJDObject = Nothing
    Set oCrossSectionServices = Nothing
    Set oHgrPortHlpr = Nothing
    Set oPartOcc = Nothing
    Set oHgrPart = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    ' <Determine BOM description here>
    Dim sItemNo As Long

    sItemNo = GetAttributeFromObject(pSupportComp, "ItemNo")
    bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
