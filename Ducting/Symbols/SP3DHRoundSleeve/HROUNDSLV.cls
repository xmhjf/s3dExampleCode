VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HROUNDSLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HROUNDSLV.cls
'   ProgID:         SP3DHRoundSleeve.HROUNDSLV
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'           Hvac Round Sleeve
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New HVAC symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRoundSleeve.HROUNDSLV"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRoundSleeve.HROUNDSLV"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed SP3DHRoundSleeve.HROUNDSLV Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Clearance"
    '          3. "Thickness"
    '          4. "Length"
    
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    'Following indented lines will be removed very soon. This has been
    'included to improve the performance. These lines will be removed as soon as
    'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.
        ' set the CMCacheForPart method in the definition
        
    ' Create a new input by new operator
    Dim Inputs(1 To 4) As IMSSymbolEntities.DInput
    Dim Index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).Name = "Width"
    Inputs(1).Description = "Width of the Rect Duct"
    Inputs(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3
    Inputs(1).DefaultParameterValue = PC
    
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Clearance"
    Inputs(2).Description = "Clearance between ducts"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.006
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "Thickness"
    Inputs(3).Description = "Thickness of the Sleeve"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.0012
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "Length"
    Inputs(4).Description = "Width of the Sleeve"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.1
    Inputs(4).DefaultParameterValue = PC
       
    For Index = 1 To 4
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 2) As IMSSymbolEntities.DOutput
    
    For Index = 1 To 2
        Set O(Index) = New IMSSymbolEntities.DOutput
        O(Index).Properties = 0
    Next

   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Sleeve"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Sleeve"
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For Index = 1 To 2
        oRepPhysicalOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

   
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRoundSleeve.HROUNDSLV"

    
    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    
    Set PhysicalRepEval = Nothing
    
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    Set oRepPhysicalOutputs = Nothing
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    'If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
        'End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim Clearance As Double, Thickness As Double
    Dim Length As Double
    Dim HLength As Double
    Dim HClearance As Double
    Dim NozzleLength As Double, NozzleFlangeThickness As Double
    Dim CornerRadius As Double
    
    Dim FlangeWidth As Double
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Clearance = arrayOfInputs(3)
    Thickness = arrayOfInputs(4)
    Length = arrayOfInputs(5)
    Depth = Width
    FlangeWidth = Thickness
    HLength = Length / 2

'=====================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim nozzleIndex As Integer
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    
    Dim iPortIndex As Integer

    'Set HVAC nozzle parameters
    iPortIndex = 1
    Dim PortDepth As Double
    Dim CptOffset As Double
    PortDepth = 0#
    CptOffset = 0#
    
    EndPrep = 11
    CornerRadius = 0#
    HClearance = Clearance / 2
    NozzleFlangeThickness = HLength + HClearance
    NozzleLength = NozzleFlangeThickness + NEGLIGIBLE_THICKNESS
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    
    
    FlowDir = DistribFlow_BOTH

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -HClearance
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    iPortIndex = 2
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = HClearance
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    'dir.Set 0, 1, 0
    'iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
'BUILD HVACNOZZLE2 Of Sleeve
'=====================================
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRECTFF definition"
        Debug.Assert False
End Sub

