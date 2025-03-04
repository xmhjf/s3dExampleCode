VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hgr2CChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   Hgr2CChannel.cls
'   ProgID:         HS_Hilti_MQParts.Hgr2CChannel
'   Author:         JOGN
'   Creation Date:  07.Sept.2007
'   Description: Channel
'
'   Change History:
'   Date              who            change description
'   16-02-2010        VSP            Addition to S3D
'   08/02/2011        Ramya          TR # 183457  Length and Weight properties are incorrect for a Hilti Part in a Specific case
'   02/09/11        Ramya       TR 191732  Detailed Physical aspect need to be removed from Smart Parts
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
'Create Constant Used for error messages
Private Const MODULE = "Hgr2CChannel"
Private Const m_progID = "HS_Hilti_MQParts.Hgr2CChannel"

Private Const INPUT_1 = "Part"
Private Const INPUT_2 = "CardinalPoint"
Private Const INPUT_3 = "Length"
Private Const INPUT_4 = "MaterialGrade"
Private Const INPUT_5 = "Orientation"
Private Const INPUT_6 = "BeginOverLength"
Private Const INPUT_7 = "EndOverLength"
Private Const INPUT_8 = "BeginMiter"
Private Const INPUT_9 = "EndMiter"
Private Const INPUT_10 = "Finish"
Private Const INPUT_11 = "Thick"
Private Const INPUT_12 = "ItemNo"
Private Const INPUT_13 = "Size"
Private Const INPUT_14 = "StructCon"

Private Const OUTPUT_BC = "BeginCap"
Private Const OUTPUT_EC = "EndCap"
Private Const OUTPUT_NT = "Neutral"
Private Const OUTPUT_EC_T = "EndCapTop"
Private Const OUTPUT_EC_B = "EndCapBottom"
Private Const OUTPUT_BC_T = "BeginCapTop"
Private Const OUTPUT_BC_B = "BeginCapBottom"
Private Const OUTPUT_BC_SUR = "BeginCapSurface"
Private Const OUTPUT_EC_SUR = "EndCapSurface"

Private Const NUM_INPUTS = 14
Private Const NUM_PORTS = 30
Private Const NUM_OUTPUTS = NUM_PORTS + 9

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
    Set PC = New IMSSymbolEntities.DParameterContent      'not persistent PC

    Dim PC1 As IMSSymbolEntities.IJDParameterContent
    Set PC1 = New IMSSymbolEntities.DParameterContent     'not persistent PC
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
    SymbolInput(3).Description = "Length of the Beam"            'may be changed later by constraints
    SymbolInput(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.5
    SymbolInput(3).DefaultParameterValue = PC

    SymbolInput(4).name = INPUT_4
    SymbolInput(4).Description = "MaterialGrade of the Beam"     'codelist
    SymbolInput(4).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "A36"
    SymbolInput(4).DefaultParameterValue = PC1

    SymbolInput(5).name = INPUT_5
    SymbolInput(5).Description = "Orientation"                   'codelist
    SymbolInput(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 10#
    SymbolInput(5).DefaultParameterValue = PC

    SymbolInput(6).name = INPUT_6
    SymbolInput(6).Description = "Begin OverLength of the Beam"  'positive value always elongate the Beam
    SymbolInput(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(6).DefaultParameterValue = PC

    SymbolInput(7).name = INPUT_7
    SymbolInput(7).Description = "End OverLength of the Beam"    'positive value always elongate the beam
    SymbolInput(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(7).DefaultParameterValue = PC

    SymbolInput(8).name = INPUT_8
    SymbolInput(8).Description = "BeginMiter"             'positive value always elongate the beam
    SymbolInput(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(8).DefaultParameterValue = PC

    SymbolInput(9).name = INPUT_9
    SymbolInput(9).Description = "EndMiter"                'positive value always elongate the beam
    SymbolInput(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    SymbolInput(9).DefaultParameterValue = PC

    SymbolInput(10).name = INPUT_10
    SymbolInput(10).Description = "Finish"                 'codelist
    SymbolInput(10).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "Hot Dip Galvanized"
    SymbolInput(10).DefaultParameterValue = PC1

    SymbolInput(11).name = INPUT_11
    SymbolInput(11).Description = "Thick"                   'positive value always elongate the beam
    SymbolInput(11).Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igValue
    PC.UomValue = 0.002
    SymbolInput(11).DefaultParameterValue = PC

    SymbolInput(12).name = INPUT_12
    SymbolInput(12).Description = "ItemNo"
    SymbolInput(12).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "304798"
    SymbolInput(12).DefaultParameterValue = PC1

    SymbolInput(13).name = INPUT_13
    SymbolInput(13).Description = "Size"                    'codelist
    SymbolInput(13).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "3m"
    SymbolInput(13).DefaultParameterValue = PC1

    SymbolInput(14).name = INPUT_14
    SymbolInput(14).Description = "StructCon"               'codelist
    SymbolInput(14).Properties = igINPUT_IS_A_PARAMETER
    PC1.Type = igString
    PC1.String = "No"
    SymbolInput(14).DefaultParameterValue = PC1

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
    SymbolOutput(1).name = OUTPUT_BC
    SymbolOutput(1).Description = OUTPUT_BC
    SymbolOutput(1).Properties = 0

    Set SymbolOutput(2) = New IMSSymbolEntities.DOutput
    SymbolOutput(2).name = OUTPUT_EC
    SymbolOutput(2).Description = OUTPUT_EC
    SymbolOutput(2).Properties = 0

    Set SymbolOutput(3) = New IMSSymbolEntities.DOutput
    SymbolOutput(3).name = OUTPUT_NT
    SymbolOutput(3).Description = OUTPUT_NT
    SymbolOutput(3).Properties = 0

    Set SymbolOutput(4) = New IMSSymbolEntities.DOutput
    SymbolOutput(4).name = OUTPUT_EC_T
    SymbolOutput(4).Description = OUTPUT_EC_T
    SymbolOutput(4).Properties = 0

    Set SymbolOutput(5) = New IMSSymbolEntities.DOutput
    SymbolOutput(5).name = OUTPUT_EC_B
    SymbolOutput(5).Description = OUTPUT_EC_B
    SymbolOutput(5).Properties = 0

    Set SymbolOutput(6) = New IMSSymbolEntities.DOutput
    SymbolOutput(6).name = OUTPUT_BC_T
    SymbolOutput(6).Description = OUTPUT_BC_T
    SymbolOutput(6).Properties = 0

    Set SymbolOutput(7) = New IMSSymbolEntities.DOutput
    SymbolOutput(7).name = OUTPUT_BC_B
    SymbolOutput(7).Description = OUTPUT_BC_B
    SymbolOutput(7).Properties = 0

    Set SymbolOutput(8) = New IMSSymbolEntities.DOutput
    SymbolOutput(8).name = OUTPUT_BC_SUR
    SymbolOutput(8).Description = OUTPUT_BC_SUR
    SymbolOutput(8).Properties = 0

    Set SymbolOutput(9) = New IMSSymbolEntities.DOutput
    SymbolOutput(9).name = OUTPUT_EC_SUR
    SymbolOutput(9).Description = OUTPUT_EC_SUR
    SymbolOutput(9).Properties = 0

    For ii = 1 To NUM_PORTS
        Set SymbolOutput(ii + 9) = New IMSSymbolEntities.DOutput
        OutputName = "Port" + CStr(ii)
        SymbolOutput(ii + 9).name = OutputName
        SymbolOutput(ii + 9).Description = OutputName
        SymbolOutput(ii + 9).Properties = 0
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
    oRep1.RepresentationId = SimplePhysical             'define a aspect 0 (Simple_physical)
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
    oRep1.RepresentationId = DetailPhysical              'Detailed Physical
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
    Dim lStructCon As Long
    Dim oCrossSection As Object
    Dim varWidth As Variant
    Dim dblWidth As Double
    Dim varDepth As Variant
    Dim dblDepth As Double
    Dim lPortCount As Integer
    Dim oBeginCapSurface As Object
    Dim oEndCapSurface As Object
    Dim oNeutral As Object
    Dim oBeginCapCenter As Object
    Dim oEndCapCenter As Object
    Dim oBeginCapRight As Object
    Dim oBeginCapLeft As Object
    Dim oBeginCapTop As Object
    Dim oBeginCapBottom As Object
    Dim oEndCapRight As Object
    Dim oEndCapLeft As Object
    Dim oEndCapTop As Object
    Dim oEndCapBottom As Object

    Dim oCatServices As New GSCADHgrSupCatalogServices.CatalogServices

    Set pRepDG = pRepSCM
    Set pDefinition = pRepDG.Definition
    Set Inputs = pDefinition.IJDInputs
    Set m_outputColl = pRepDG.OutputCollection

    Set oHgrSymbolPortHlpr = New HgrSymbolPortHlpr
    Set oCrossSectionServices = New CrossSectionServices

    Set oPart = Inputs.GetInputAtIndex(1).IJDInputDuringGame.Result
    CardinalPoint = Inputs.GetInputAtIndex(2).IJDInputDuringGame.Result.UomValue
    LENGTH = Inputs.GetInputAtIndex(3).IJDInputDuringGame.Result.UomValue
    BeginOverLength = Inputs.GetInputAtIndex(6).IJDInputDuringGame.Result.UomValue
    EndOverLength = Inputs.GetInputAtIndex(7).IJDInputDuringGame.Result.UomValue
    lStructCon = Inputs.GetInputByIndex(14).IJDInputDuringGame.Result.UomValue

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
    Dim oPort3 As IJHgrPort
    Dim oPort4 As IJHgrPort
    Dim oPort5 As IJHgrPort
    Dim oPort6 As IJHgrPort
    Dim oPort7 As IJHgrPort
    Dim oPort8 As IJHgrPort
    Dim oPort9 As IJHgrPort

    Dim oElements As IJElements

    Set oPortFac = New HgrPortFactory

    oX = cardinalX
    oY = cardinalY
    oZ = 0#

    Set oElements = oPortFac.CreateHgrPortsFromPart(m_outputColl.ResourceManager, oPart)

    Set oPort1 = GetPortByNameFromCollection(oElements, "BeginCap")
    oPort1.PutOrientation oX + dblWidth / 2, oY + dblDepth / 2, 0, _
                          0, 0, 1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_BC, oPort1

    Set oPort2 = GetPortByNameFromCollection(oElements, "EndCap")
    oPort2.PutOrientation oX + dblWidth / 2, oY + dblDepth / 2, oZ + LENGTH, _
                          0, 0, -1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_EC, oPort2

    Set oPort3 = GetPortByNameFromCollection(oElements, "Neutral")
    oPort3.PutOrientation oX + dblWidth / 2, oY + dblDepth / 2, oZ + LENGTH / 2#, _
                          0, 0, 1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_NT, oPort3

    Set oPort4 = GetPortByNameFromCollection(oElements, "EndCapTop")
    oPort4.PutOrientation oX + dblWidth, oY + dblDepth / 2, LENGTH, _
                          0, 0, -1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_EC_T, oPort4

    Set oPort5 = GetPortByNameFromCollection(oElements, "EndCapBottom")
    oPort5.PutOrientation oX, oY + dblDepth / 2, LENGTH, _
                          0, 0, -1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_EC_B, oPort5

    Set oPort6 = GetPortByNameFromCollection(oElements, "BeginCapTop")
    oPort6.PutOrientation oX + dblWidth, oY + dblDepth / 2, 0, _
                          0, 0, 1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_BC_T, oPort6

    Set oPort7 = GetPortByNameFromCollection(oElements, "BeginCapBottom")
    oPort7.PutOrientation oX, oY + dblDepth / 2, 0, _
                          0, 0, 1, -1, 0, 0
    m_outputColl.AddOutput OUTPUT_BC_B, oPort7

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

    dblLength = GetAttributeFromObject(pSupportComp, "Length")
    dblBeginLength = GetAttributeFromObject(pSupportComp, "BeginOverLength")
    dblEndLength = GetAttributeFromObject(pSupportComp, "EndOverLength")

    If (dblLength <= 0) Then
        dblLength = 0.5
        
        Dim oMyPart As IJHgrSupportComponent
        Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr
        
        Set oMyPart = pSupportComp
        oMyPart.GetOccAssembly oICH
        oICH.SetAttributeValue "Length", oMyPart, dblLength
        Set oMyPart = Nothing
        Set oICH = Nothing
    End If

    dblTotalLength = dblBeginLength + dblLength + dblEndLength

    Set oPartOcc = pSupportComp

    oPartOcc.GetPart oHgrPart

    Set oHgrPortHlpr = New HgrSymbolPortHlpr
    Call oHgrPortHlpr.GetHgrAssociation(oHgrPart, 3, oCrossSectionColl)

    Set oCSObj = oCrossSectionColl.Item(1)

    Set oCrossSectionServices = New CrossSectionServices

    oCrossSectionServices.PlaceCrossSectionOccurrence GetModelResourceManager(), oCSObj, True, oCSOcc

    oCrossSectionServices.GetCardinalPoint oCSOcc, 10, CogX, CogY

    CogZ = dblTotalLength / 2#

    Set oIJDObject = oCSOcc

    oIJDObject.Remove

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
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim strFinish As String
    Dim lFinish As Long
    Dim sItemNo As Long
    Dim dblLength As Double
    Dim dblBeginLength As Double
    Dim dblEndLength As Double
    Dim sSize As String
    Dim dSize As Double
    Dim dNumChannel As Double
    Dim dblTotalLength As Double
    Dim sPartNumber As String
    Dim dblDryWeight As Double
    Dim dNumConnect As Double
    Dim dTempLength As Double
    Dim lStructCon As Long
    Dim sConnType As String

    dblLength = GetAttributeFromObject(pSupportComp, "Length")
    dblBeginLength = GetAttributeFromObject(pSupportComp, "BeginOverLength")
    dblEndLength = GetAttributeFromObject(pSupportComp, "EndOverLength")
    lStructCon = GetAttributeFromObject(pSupportComp, "StructCon")

    dblTotalLength = dblBeginLength + dblLength + dblEndLength

    Dim oUomServices As UnitsOfMeasureServicesLib.UomVBInterface
    Set oUomServices = New UnitsOfMeasureServicesLib.UomVBInterface
    
    Dim sCutlen As String
    
    Dim xomFormat As IJUomVBFormat
    Set xomFormat = New UomVBFormat
    xomFormat.PrecisionType = PRECISIONTYPE_DECIMAL
    xomFormat.FractionalPrecision = 2
    xomFormat.UnitsDisplayed = True
    xomFormat.ReduceFraction = True
                
    oUomServices.FormatUnit UNIT_DISTANCE, dblTotalLength, sCutlen, xomFormat, DISTANCE_MILLIMETER

    sSize = GetAttributeFromObject(pSupportComp, "Size")
    sSize = Left(sSize, 1)
    dSize = Val(sSize)

    dNumChannel = Int(dblTotalLength / dSize)

    If (dNumChannel * dSize) < dblTotalLength Then
        dNumChannel = Int(dblTotalLength / dSize) + 1
    End If

    If lStructCon <> 3 Then
        dNumConnect = 2
        dTempLength = dblTotalLength

        Do While dTempLength > 0.5
            dNumConnect = dNumConnect + 1
            dTempLength = dTempLength - 0.5
        Loop
    End If

    sItemNo = GetAttributeFromObject(pSupportComp, "ItemNo")
    lFinish = GetAttributeFromObject(pSupportComp, "Finish")
    strFinish = ConvertValueToLongStringValue(pSupportComp, "IJUAMQFinish", "Finish", lFinish)
    sPartNumber = oPart.PartNumber
    sConnType = ConvertValueToLongStringValue(pSupportComp, "IJUAMQStrucCon", "StructCon", lStructCon)

    If Trim(sConnType) <> "No" Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " Length = " & sCutlen & ", Number of Channels " & dNumChannel & ", + " & dNumConnect & ", " & sConnType
    Else
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " Length = " & sCutlen & ", Number of Channels " & dNumChannel
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
