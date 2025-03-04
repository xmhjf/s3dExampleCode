VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MI_Bracket_Flange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   MI_Bracket_Flange.cls
'   ProgID:         HS_Hilti_MIParts.MI_Bracket_Flange
'   Author:         JOGN
'   Creation Date:  October 17, 2007
'   Description: Channel
'
'   Change History:
'   Date              who            change description
'   16-02-2010        VSP            Addition to S3D
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
'Create Constant Used for error messages
Private Const MODULE = "MI_Bracket_Flange"
Private Const m_progID = "HS_Hilti_MIParts.MI_Bracket_Flange"

Private Const INPUT_1 = "Part"
Private Const INPUT_2 = "CardinalPoint"
Private Const INPUT_3 = "MaterialGrade"
Private Const INPUT_4 = "SteelWidthFrom"
Private Const INPUT_5 = "SteelWidthTo"
Private Const INPUT_6 = "Orientation"
Private Const INPUT_7 = "EndOverLength"
Private Const INPUT_8 = "BeginOverLength"
Private Const INPUT_9 = "Length"
Private Const INPUT_10 = "BeginMiter"
Private Const INPUT_11 = "EndMiter"
Private Const INPUT_12 = "Thickness"
Private Const INPUT_13 = "PlateThick"
Private Const INPUT_14 = "ItemNo"
Private Const INPUT_15 = "Finish"
Private Const INPUT_16 = "Group"
Private Const INPUT_17 = "GirderL"
Private Const INPUT_18 = "GirderW"
Private Const INPUT_19 = "GirderD"
Private Const INPUT_20 = "PlateDepth"
Private Const INPUT_21 = "PlateWidth"
Private Const INPUT_22 = "PortOffset"

Private Const STRUCTUREINT = "StructureInt"
Private Const OUTPUT_BC = "BeginCap"
Private Const OUTPUT_EC = "EndCap"
Private Const OUTPUT_EC_CEN = "EndCapCenter"
Private Const OUTPUT_ECR = "EndCapRight"
Private Const OUTPUT_ECL = "EndCapLeft"
Private Const OUTPUT_ECT = "EndCapTop"
Private Const OUTPUT_ECB = "EndCapBottom"
Private Const OUTPUT_PLATE = "Plate"
Private Const OUTPUT_EC_SUR = "EndCapSurface"

Private Const NUM_INPUTS = 22
Private Const NUM_PORTS = 30
Private Const NUM_OUTPUTS = NUM_PORTS + 10

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
    oSymbolDefinition.name = oSymbolDefinition.ProgId

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
    Set PC = New IMSSymbolEntities.DParameterContent     'not persistent PC

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

    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = "HgrSupportSymbolUtilities.CustomMethods"

    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc

    'Get the lib/method cookie
    libCookie = libDesc.Cookie
    mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForHgrPart", libCookie)

    'There should be one input defined for refData part.
    'This should be the first input
    SymbolInput(1).name = INPUT_1
    SymbolInput(1).Description = "Part"
    SymbolInput(1).IJDInputStdCustomMethod.SetCMCache libCookie, mCookie

    SymbolInput(2).name = INPUT_2
    SymbolInput(2).Description = "Cardinality Point of the Beam"
    SymbolInput(2).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 1
    SymbolInput(2).DefaultParameterValue = PC

    SymbolInput(3).name = INPUT_3
    SymbolInput(3).Description = "MaterialGrade of the Beam"    'codelist
    SymbolInput(3).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "A36"
    SymbolInput(3).DefaultParameterValue = PC1

    SymbolInput(4).name = INPUT_4
    SymbolInput(4).Description = "SteelWidthFrom"
    SymbolInput(4).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.075
    SymbolInput(4).DefaultParameterValue = PC

    SymbolInput(5).name = INPUT_5
    SymbolInput(5).Description = "SteelWidthTo"
    SymbolInput(5).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.165
    SymbolInput(5).DefaultParameterValue = PC

    SymbolInput(6).name = INPUT_6
    SymbolInput(6).Description = "Orientation"    'codelist
    SymbolInput(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 10#
    SymbolInput(6).DefaultParameterValue = PC

    SymbolInput(7).name = INPUT_7
    SymbolInput(7).Description = "End OverLength of the Beam"    'positive value always elongate the beam
    SymbolInput(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(7).DefaultParameterValue = PC

    SymbolInput(8).name = INPUT_8
    SymbolInput(8).Description = "Begin OverLength of the Beam"    'positive value always elongate the Beam
    SymbolInput(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(8).DefaultParameterValue = PC

    SymbolInput(9).name = INPUT_9
    SymbolInput(9).Description = "Length of the Beam"    'may be changed later by constraints
    SymbolInput(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.5
    SymbolInput(9).DefaultParameterValue = PC

    SymbolInput(10).name = INPUT_10
    SymbolInput(10).Description = "BeginMiter"       'positive value always elongate the beam
    SymbolInput(10).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(10).DefaultParameterValue = PC

    SymbolInput(11).name = INPUT_11
    SymbolInput(11).Description = "EndMiter"        'positive value always elongate the beam
    SymbolInput(11).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(11).DefaultParameterValue = PC

    SymbolInput(12).name = INPUT_12
    SymbolInput(12).Description = "Thickness"       'positive value always elongate the beam
    SymbolInput(12).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.035
    SymbolInput(12).DefaultParameterValue = PC

    SymbolInput(13).name = INPUT_13
    SymbolInput(13).Description = "PlateThick"      'positive value always elongate the beam
    SymbolInput(13).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.012
    SymbolInput(13).DefaultParameterValue = PC

    SymbolInput(14).name = INPUT_14
    SymbolInput(14).Description = "ItemNo"
    SymbolInput(14).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "267774"
    SymbolInput(14).DefaultParameterValue = PC1

    SymbolInput(15).name = INPUT_15
    SymbolInput(15).Description = "Finish"    'codelist
    SymbolInput(15).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "Hot Dip Galvanized"
    SymbolInput(15).DefaultParameterValue = PC1

    SymbolInput(16).name = INPUT_16
    SymbolInput(16).Description = "Group"    'codelist
    SymbolInput(16).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "Bracket"
    SymbolInput(16).DefaultParameterValue = PC1

    SymbolInput(17).name = INPUT_17
    SymbolInput(17).Description = "GirderL"
    SymbolInput(17).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.5
    SymbolInput(17).DefaultParameterValue = PC

    SymbolInput(18).name = INPUT_18
    SymbolInput(18).Description = "GirderW"
    SymbolInput(18).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.09
    SymbolInput(18).DefaultParameterValue = PC

    SymbolInput(19).name = INPUT_19
    SymbolInput(19).Description = "GirderD"
    SymbolInput(19).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.09
    SymbolInput(19).DefaultParameterValue = PC

    SymbolInput(20).name = INPUT_20
    SymbolInput(20).Description = "PlateDepth"
    SymbolInput(20).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.22
    SymbolInput(20).DefaultParameterValue = PC

    SymbolInput(21).name = INPUT_21
    SymbolInput(21).Description = "PlateWidth"
    SymbolInput(21).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.28
    SymbolInput(21).DefaultParameterValue = PC

    SymbolInput(22).name = INPUT_22
    SymbolInput(22).Description = "PortOffset"
    SymbolInput(22).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.008
    SymbolInput(22).DefaultParameterValue = PC

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
    SymbolOutput(1).name = STRUCTUREINT
    SymbolOutput(1).Description = STRUCTUREINT
    SymbolOutput(1).Properties = 0

    Set SymbolOutput(2) = New IMSSymbolEntities.DOutput
    SymbolOutput(2).name = OUTPUT_BC
    SymbolOutput(2).Description = OUTPUT_BC
    SymbolOutput(2).Properties = 0

    Set SymbolOutput(3) = New IMSSymbolEntities.DOutput
    SymbolOutput(3).name = OUTPUT_EC
    SymbolOutput(3).Description = OUTPUT_EC
    SymbolOutput(3).Properties = 0

    Set SymbolOutput(4) = New IMSSymbolEntities.DOutput
    SymbolOutput(4).name = OUTPUT_EC_CEN
    SymbolOutput(4).Description = OUTPUT_EC_CEN
    SymbolOutput(4).Properties = 0

    Set SymbolOutput(5) = New IMSSymbolEntities.DOutput
    SymbolOutput(5).name = OUTPUT_ECR
    SymbolOutput(5).Description = OUTPUT_ECR
    SymbolOutput(5).Properties = 0

    Set SymbolOutput(6) = New IMSSymbolEntities.DOutput
    SymbolOutput(6).name = OUTPUT_ECL
    SymbolOutput(6).Description = OUTPUT_ECL
    SymbolOutput(6).Properties = 0

    Set SymbolOutput(7) = New IMSSymbolEntities.DOutput
    SymbolOutput(7).name = OUTPUT_ECT
    SymbolOutput(7).Description = OUTPUT_ECT
    SymbolOutput(7).Properties = 0

    Set SymbolOutput(8) = New IMSSymbolEntities.DOutput
    SymbolOutput(8).name = OUTPUT_ECB
    SymbolOutput(8).Description = OUTPUT_ECB
    SymbolOutput(8).Properties = 0

    Set SymbolOutput(9) = New IMSSymbolEntities.DOutput
    SymbolOutput(9).name = OUTPUT_PLATE
    SymbolOutput(9).Description = OUTPUT_PLATE
    SymbolOutput(9).Properties = 0
    Set SymbolOutput(10) = New IMSSymbolEntities.DOutput
    SymbolOutput(10).name = OUTPUT_EC_SUR
    SymbolOutput(10).Description = OUTPUT_EC_SUR
    SymbolOutput(10).Properties = 0

    For ii = 1 To NUM_PORTS
        Set SymbolOutput(ii + 10) = New IMSSymbolEntities.DOutput
        OutputName = "Port" + CStr(ii)
        SymbolOutput(ii + 10).name = OutputName
        SymbolOutput(ii + 10).Description = OutputName
        SymbolOutput(ii + 10).Properties = 0
    Next

    'DEFINE THE "SimplePhysical" REPRESENTATION
    Dim oRep1 As IMSSymbolEntities.IJDRepresentation
    Set oRep1 = New IMSSymbolEntities.DRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep1


    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.

    oRep1.name = "Symbolic"
    oRep1.Description = "Symbolic Represntation of the 3d flexible"
    oRep1.RepresentationId = SimplePhysical      'define a aspect 0 (Simple_physical)
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    'Set the output
    For ii = 1 To NUM_OUTPUTS
        oOutputs.SetOutput SymbolOutput(ii)
    Next

    Dim libDesc1 As New DLibraryDescription

    libDesc1.name = "mySelfAsLib1"
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

    oRep1.name = "Detailed"
    oRep1.Description = "Detailed Represntation of the 3d flexible"
    oRep1.RepresentationId = DetailPhysical     'Detailed Physical
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    mthCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("Detailed", libCookie)

    'set the customer evaluation method
    oRep1.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mthCookie

    'Set the representation to definition
    oReps.SetRepresentation oRep1

    'Set the script associated to the Detailed representation
    Dim oVbFuncDetailedRep As DRepresentationEvaluation

    'DEFINE Maintenance Envelop REPRESENTATION WITH DIFFERENT OUTPUTS
    oRep1.name = "Maintenance"
    oRep1.Description = "Maintenance Represntation of the 3d flexible"
    oRep1.RepresentationId = Maintenance    'Maintenance
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
    'pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
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
' This procedure will be used to create the physical representation
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
    Dim CSName As String
    Dim CSStandard As String
    Dim SectionType As String
    Dim LENGTH As Double
    Dim WIDTH As Double
    Dim Thickness As Double
    Dim BeginOverLength As Double
    Dim EndOverLength As Double
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
    Dim oEndCapSurface As Object
    Dim oBeginCapCenter As Object
    Dim oEndCapCenter As Object
    Dim oEndCapRight As Object
    Dim oEndCapLeft As Object
    Dim oEndCapTop As Object
    Dim oEndCapBottom As Object
    Dim dGirderL As Double
    Dim dPlateThick As Double
    Dim dPlateDepth As Double
    Dim dPlateWidth As Double
    Dim dPortOffset As Double
    Dim oCatServices As New GSCADHgrSupCatalogServices.CatalogServices
    Dim dSteelWidthFrom As Double
    Dim dSteelWidthTo As Double
    Dim oPartOcc As PartOcc

    Dim oSupComp As IJHgrSupportComponent
    Dim oICH As IJHgrInputConfigHlpr


    Dim dSteelDepth As Double
    Dim dSteelWidth As Double
    Dim dFlangeThick As Double
    Dim dWebThick As Double
    Dim dPipeDia As Double
    Dim sErrMsg As String

    Set pRepDG = pRepSCM
    Set pDefinition = pRepDG.Definition

    Set Inputs = pDefinition.IJDInputs
    Set m_outputColl = pRepDG.OutputCollection

    Set oHgrSymbolPortHlpr = New HgrSymbolPortHlpr
    Set oCrossSectionServices = New CrossSectionServices

    Set oPart = Inputs.GetInputAtIndex(1).IJDInputDuringGame.Result
    CardinalPoint = Inputs.GetInputAtIndex(2).IJDInputDuringGame.Result.UomValue
    dSteelWidthFrom = Inputs.GetInputByIndex(4).IJDInputDuringGame.Result.UomValue
    dSteelWidthTo = Inputs.GetInputByIndex(5).IJDInputDuringGame.Result.UomValue
    BeginOverLength = Inputs.GetInputAtIndex(8).IJDInputDuringGame.Result.UomValue
    EndOverLength = Inputs.GetInputAtIndex(7).IJDInputDuringGame.Result.UomValue
    dPlateThick = Inputs.GetInputByIndex(13).IJDInputDuringGame.Result.UomValue
    dGirderL = Inputs.GetInputByIndex(17).IJDInputDuringGame.Result.UomValue
    dPlateDepth = Inputs.GetInputByIndex(20).IJDInputDuringGame.Result.UomValue
    dPlateWidth = Inputs.GetInputByIndex(21).IJDInputDuringGame.Result.UomValue
    dPortOffset = Inputs.GetInputByIndex(22).IJDInputDuringGame.Result.UomValue
    'LENGTH = dGirderL
    LENGTH = Inputs.GetInputAtIndex(9).IJDInputDuringGame.Result.UomValue
          
    
    Set oSupComp = pRepDG.Definition.IJDDefinitionPlayerEx.PlayingSymbol
    oSupComp.GetOccAssembly oICH

    If oICH Is Nothing Then    ' The OCH will be null the first time thru until the symbol is computed the first time
        dPipeDia = 0.168275    ' 6" pipe dia
    Else
        Hilti_InitializeMyHH oICH
        
        Dim oStruct As IJElements
        oICH.GetSupportingCollection oStruct
        
        If Not oStruct Is Nothing Then
            Dim SupportStruct As hsSteelMember
            SupportStruct = GetSupportingSectionData(1)
            dSteelWidth = SupportStruct.dWidth
            dFlangeThick = SupportStruct.dFlangeThickness
        
                If dSteelWidth > 0.3 Then
                    PF_EventHandler "Steel Width can not exceed 300 mm.", Err, MODULE, METHOD, True
                Else
                    If dSteelWidth < dSteelWidthFrom Or dSteelWidth > dSteelWidthTo Then
                        sErrMsg = "Steel Width must be between " & dSteelWidthFrom * 1000 & " to " & dSteelWidthTo * 1000 & " mm"
                        PF_EventHandler sErrMsg, Err, MODULE, METHOD, True
                    End If
                End If
        
                If dFlangeThick < 0.003 Or dFlangeThick > 0.036 Then
                    PF_EventHandler "Flange thickness must be between 3 to 36 mm", Err, MODULE, METHOD, True
                End If
        
        End If
        
        Set oStruct = Nothing
        
        Hilti_DestroyMyHH oICH
        Set oICH = Nothing
    End If
    
    LENGTH = LENGTH - dPlateThick
    Set oPorts = oCatServices.GetConnectionPortColl( _
                 m_outputColl.ResourceManager, oPart, _
                 CardinalPoint, _
                 BeginOverLength, LENGTH, EndOverLength, _
                 cardinalX, cardinalY)
    LENGTH = LENGTH + dPlateThick
    Dim OutputName As String

    lPortCount = oPorts.Count - 1

    For ii = 1 To lPortCount
        OutputName = "Port" + CStr(ii)
        m_outputColl.AddOutput OutputName, oPorts.Item(ii)
    Next

    Set oEndCapSurface = oPorts.Item(lPortCount + 1)
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
    Dim oPort3 As IJHgrPort
    Dim oPort4 As IJHgrPort
    Dim oPort5 As IJHgrPort
    Dim oPort6 As IJHgrPort
    Dim oPort7 As IJHgrPort
    Dim oPort8 As IJHgrPort
    Dim oElements As IJElements

    Set oPortFac = New HgrPortFactory

    oX = cardinalX
    oY = cardinalY
    oZ = 0#

    Set oElements = oPortFac.CreateHgrPortsFromPart(m_outputColl.ResourceManager, oPart)

    Set oPort1 = GetPortByNameFromCollection(oElements, "StructureInt")
    oPort1.PutOrientation oX + dblWidth / 2#, oY + dblDepth / 2#, oZ - dPlateThick, _
                          0, 0, 1, 0, -1, 0
    m_outputColl.AddOutput STRUCTUREINT, oPort1

    Set oPort2 = GetPortByNameFromCollection(oElements, "BeginCap")
    oPort2.PutOrientation oX + dblWidth / 2#, oY + dblDepth / 2#, oZ, _
                          0, 0, 1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_BC, oPort2

    Set oPort3 = GetPortByNameFromCollection(oElements, "EndCap")
    oPort3.PutOrientation oX + dblWidth / 2#, oY + dblDepth / 2#, oZ + LENGTH - dPlateThick, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_EC, oPort3

    Set oPort4 = GetPortByNameFromCollection(oElements, "EndCapCenter")
    oPort4.PutOrientation oX + dblWidth / 2#, oY + dblDepth / 2#, oZ + LENGTH - dPlateThick - 0.025, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_EC_CEN, oPort4

    Set oPort5 = GetPortByNameFromCollection(oElements, "EndCapRight")
    oPort5.PutOrientation oX + dblWidth, oY + dblDepth / 2, oZ + LENGTH - dPlateThick, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_ECR, oPort5

    Set oPort6 = GetPortByNameFromCollection(oElements, "EndCapLeft")
    oPort6.PutOrientation oX, oY + dblDepth / 2, oZ + LENGTH - dPlateThick, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_ECL, oPort6

    Set oPort7 = GetPortByNameFromCollection(oElements, "EndCapTop")
    oPort7.PutOrientation oX + dblWidth / 2, oY + dblDepth, oZ + LENGTH - dPlateThick, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_ECT, oPort7

    Set oPort8 = GetPortByNameFromCollection(oElements, "EndCapBottom")
    oPort8.PutOrientation oX + dblWidth / 2, oY, oZ + LENGTH - dPlateThick, _
                          0, 0, -1, 0, -1, 0
    m_outputColl.AddOutput OUTPUT_ECB, oPort8

    AddBox dPlateWidth, dPlateDepth, dPlateThick, Loc(oX + dblWidth / 2# - dPlateWidth / 2, oY + dblDepth / 2# - dPlateDepth / 2, oZ - dPlateThick - BeginOverLength), m_outputColl, "Plate"

    Set oPort1 = Nothing
    Set oPort2 = Nothing
    Set oPortFac = Nothing
    oPorts.Clear
    Set pRepDG = Nothing
    Set oCrossSectionColl = Nothing
    Set oCrossSectionServices = Nothing
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

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    ' <Determine BOM description here>
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim oMyPart As IJHgrSupportComponent
    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
    
    Set oMyPart = pSupportComp
    oMyPart.GetOccAssembly oICH
    
    Dim strFinish As String
    Dim lFinish As Long
    Dim sItemNo As Long
    Dim dblLength As Double
    Dim dblBeginLength As Double
    Dim dblEndLength As Double
    Dim dGirderL As Double
    Dim dblTotalLength As Double
    Dim sPartNumber As String
    Dim MinLen As Double
      
    MinLen = GetAttributeFromObject(oPart, "MinLen")
    dblLength = GetAttributeFromObject(pSupportComp, "Length")
    dGirderL = GetAttributeFromObject(pSupportComp, "GirderL")
    dblBeginLength = GetAttributeFromObject(pSupportComp, "BeginOverLength")
    dblEndLength = GetAttributeFromObject(pSupportComp, "EndOverLength")

    
    If dblLength = 0 Then   'Intially Length will be zero
        dblLength = 0.5
    End If
    
    If ((dblLength + dblBeginLength + dblEndLength) > dGirderL) Then
        PF_EventHandler "Length is greater than " & dGirderL * 1000 & " mm", Err, MODULE, METHOD, True
        dblLength = dGirderL - dblBeginLength - dblEndLength
        oICH.SetAttributeValue "Length", oMyPart, dblLength
    End If
    If ((dblLength + dblBeginLength + dblEndLength) < MinLen) Then
        PF_EventHandler "Length is less than " & MinLen * 1000 & " mm", Err, MODULE, METHOD, True
    End If


    Set oMyPart = Nothing
    Set oICH = Nothing


    dblTotalLength = dblBeginLength + dblLength + dblEndLength
    
    If (dblLength < dGirderL) Then
        bstrBOMDesc = Hilti_BuildBomForBracket(pSupportComp, (dblLength + dblBeginLength + dblEndLength))
    Else
        bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub


