VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRECTFFDInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    HRECTFFDInc.cls
'
'Author
'   ProgID:         SP3DHRectFlatFlangeDInc.HRECTFFDInc
'   Author:         MS
'   Creation Date:  Friday, Sep 6 2002
'
'Description
'    Definition of HVAC Rectangular Flat Flange Where the Depth is inclined.
'
'Notes
'   This symbol is made on the basis of RectFlatFlange
'
'History:

'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007     PS      CR-132021  Modified the default value of Angle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection
    
Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRectFlatFlangeDInc.HRECTFFDInc"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRectFlatFlangeDInc.HRECTFFDInc"
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

    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    'Following indented lines will be removed very soon. This has been
    'included to improve the performance. These lines will be removed as soon as
    'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.
        ' set the CMCacheForPart method in the definition
        
    ' Create a new input by new operator
    Dim Inputs(1 To 9) As IMSSymbolEntities.DInput
    Dim Index As Integer
   
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
   
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Width"
    Inputs(2).Description = "Width of the Rect Duct"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "Depth"
    Inputs(3).Description = "Depth of the Rect Duct"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "Clearance"
    Inputs(4).Description = "Clearance between ducts"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.008
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "Thickness"
    Inputs(5).Description = "Thickness of the Flat Flange"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.06
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "FlangeWidth"
    Inputs(6).Description = "Width of the Flat Flange"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.04
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "CompressedGasket"
    Inputs(7).Description = "Compressed Gasket Thickness"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.004
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "Angle"
    Inputs(8).Description = "Depth Inclination Angle"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = (1 / 15) * Atn(1)
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "InsulationThickness"
    Inputs(9).Description = "Insulation Thickness"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(9).DefaultParameterValue = PC
    
    For Index = 2 To 9
        InputsIf.SetInput Inputs(Index), Index
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As IMSSymbolEntities.DOutput
    
    For Index = 1 To 4
        Set O(Index) = New IMSSymbolEntities.DOutput
        O(Index).Properties = 0
    Next

    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Flat Flange"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Flat Flange"
    
    O(3).Name = "RectFlange1"
    O(3).Description = "Dumb graphics of Flat Flange 1"
    
    O(4).Name = "RectFlangeIns"
    O(4).Description = "Dumb graphics of Insulation for the Flat Flanges"
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the HVAC Rect Flat Flange with Depth Inclination"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For Index = 1 To 3
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
    PhysicalRepEval.ProgId = "SP3DHRectFlatFlangeDInc.HRECTFFDInc"

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
    
    rep1.Name = "Insulation"
    rep1.Description = "Insulation Represntation of the Rectangle Flat Flange with Depth Inclination"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to Insulation. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = Insulation

    Dim oRepInsulationOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepInsulationOutputs = rep1
    
    oRepInsulationOutputs.SetOutput O(4)
    Set O(4) = Nothing
    
    'Set the representation to definition
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    'Define evaluation for Insulation representation
    Dim InsulationRepEval As DRepresentationEvaluation
    Set InsulationRepEval = New DRepresentationEvaluation
    InsulationRepEval.Name = "Insulation"
    InsulationRepEval.Description = "script for the Insulation representation"
    InsulationRepEval.Properties = igREPRESENTATION_HIDDEN
    InsulationRepEval.Type = igREPRESENTATION_VBFUNCTION
    InsulationRepEval.ProgId = "SP3DHRectFlatFlangeDInc.HRECTFFDInc"
    
    'Set the evaluations for the Physical representation on the definition
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation InsulationRepEval
    
    Set InsulationRepEval = Nothing
    
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    Set oRepInsulationOutputs = Nothing
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll

    If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
    ElseIf StrComp(repName, "Insulation") = 0 Then
        InsulationRep arrayOfInputs
    End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFFDInc
'=========================================================================

Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim Width As Double
    Dim Depth As Double
    Dim Clearance As Double
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim CompressedGasket As Double
    Dim Angle As Double
    Dim InsulationThickness As Double
    Dim oPart As PartFacelets.IJDPart
    Dim objRectFlange1 As Object
    
    Set oPart = arrayOfInputs(1)
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Clearance = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
    CompressedGasket = arrayOfInputs(7)
    Angle = arrayOfInputs(8)           'Measured in anti-clockwise direction from positive Y-axis
    InsulationThickness = arrayOfInputs(9)
    
       
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
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
                
'   Dumb graphics is used to place two flanges. End preparartion for the two ports is 301
'   and bLtWtGraphics parameter is set to TRUE

'Set HVAC nozzle 1 parameters
    iPortIndex = 1
    EndPrep = 11
    CornerRadius = 0#
    PortDepth = 0#
    CptOffset = CompressedGasket / 2
    FlowDir = DistribFlow_BOTH
               
    Dim FlangeWidth1 As Double
    Dim NozzleFlangeThickness1 As Double
    
    FlangeWidth1 = 0
    NozzleFlangeThickness1 = 0
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness1, FlangeWidth1, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, True, m_outputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -CptOffset
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NEGLIGIBLE_THICKNESS

    m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    iPortIndex = 2
    EndPrep = 301

    FlangeWidth1 = 0
    CptOffset = 0
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NozzleFlangeThickness1, FlangeWidth1, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, True, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    X = Clearance
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NEGLIGIBLE_THICKNESS

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

 'Insert code for the Dumb graphics of Rectangular flange 1
    CptOffset = CompressedGasket / 2
    
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim ProjVector As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
           
'Create Line String
    LineStrPoints(0) = -Depth / 2 * Tan(Angle) - FlangeWidth * Sin(Angle) - CptOffset
    LineStrPoints(1) = Depth / 2 + FlangeWidth * Cos(Angle)
    LineStrPoints(2) = -(Width / 2 + FlangeWidth)

    LineStrPoints(3) = -Depth / 2 * Tan(Angle) - FlangeWidth * Sin(Angle) - CptOffset
    LineStrPoints(4) = Depth / 2 + FlangeWidth * Cos(Angle)
    LineStrPoints(5) = (Width / 2 + FlangeWidth)

    LineStrPoints(6) = Depth / 2 * Tan(Angle) + FlangeWidth * Sin(Angle) - CptOffset
    LineStrPoints(7) = -(Depth / 2 + FlangeWidth * Cos(Angle))
    LineStrPoints(8) = (Width / 2 + FlangeWidth)

    LineStrPoints(9) = Depth / 2 * Tan(Angle) + FlangeWidth * Sin(Angle) - CptOffset
    LineStrPoints(10) = -(Depth / 2 + FlangeWidth * Cos(Angle))
    LineStrPoints(11) = -(Width / 2 + FlangeWidth)

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
    ProjVector.Set -Cos(Angle), -Sin(Angle), 0
    
    Set objRectFlange1 = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oLineString, _
                                        ProjVector.X, ProjVector.Y, ProjVector.Z, Thickness, True)

    'Add the output into the Collection
    m_outputColl.AddOutput "RectFlange1", objRectFlange1
    Set objRectFlange1 = Nothing
    
    Set ProjVector = Nothing
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRECTFFDInc definition"
        Debug.Assert False
End Sub

'=========================================================================
'CREATION OF INSULATION REPRESENTATION OF HRECTFFDInc
'=========================================================================

Private Sub InsulationRep(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim Width As Double
    Dim Depth As Double
    Dim Clearance As Double
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim CompressedGasket As Double
    Dim Angle As Double
    Dim InsulationThickness As Double
    
    Dim oPart As PartFacelets.IJDPart
    Dim objRectFlangeIns As Object
    
    Set oPart = arrayOfInputs(1)
'    assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Clearance = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
    CompressedGasket = arrayOfInputs(7)
    Angle = arrayOfInputs(8)
    InsulationThickness = arrayOfInputs(9)
    
        
'Insert code for the Dumb graphics of Insulation for Rect Flanges with Depth Inclination
    Dim CptOffset As Double
    CptOffset = CompressedGasket / 2
    
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim ProjVector As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
    
'Create Line String
    LineStrPoints(0) = -Depth / 2 * Tan(Angle) - (FlangeWidth + InsulationThickness) * Sin(Angle) - CptOffset _
                        - (Thickness + InsulationThickness) * Cos(Angle)
    LineStrPoints(1) = Depth / 2 + (FlangeWidth + InsulationThickness) * Cos(Angle) _
                        - (Thickness + InsulationThickness) * Sin(Angle)
    LineStrPoints(2) = -(Width / 2 + FlangeWidth)

    LineStrPoints(3) = -Depth / 2 * Tan(Angle) - (FlangeWidth + InsulationThickness) * Sin(Angle) - CptOffset _
                        - (Thickness + InsulationThickness) * Cos(Angle)
    LineStrPoints(4) = Depth / 2 + (FlangeWidth + InsulationThickness) * Cos(Angle) _
                        - (Thickness + InsulationThickness) * Sin(Angle)
    LineStrPoints(5) = (Width / 2 + FlangeWidth)

    LineStrPoints(6) = Depth / 2 * Tan(Angle) + (FlangeWidth + InsulationThickness) * Sin(Angle) - CptOffset _
                        - (Thickness + InsulationThickness) * Cos(Angle)
    LineStrPoints(7) = -(Depth / 2 + (FlangeWidth + InsulationThickness) * Cos(Angle)) _
                        - (Thickness + InsulationThickness) * Sin(Angle)
    LineStrPoints(8) = (Width / 2 + FlangeWidth)

    LineStrPoints(9) = Depth / 2 * Tan(Angle) + (FlangeWidth + InsulationThickness) * Sin(Angle) - CptOffset _
                        - (Thickness + InsulationThickness) * Cos(Angle)
    LineStrPoints(10) = -(Depth / 2 + (FlangeWidth + InsulationThickness) * Cos(Angle)) _
                            - (Thickness + InsulationThickness) * Sin(Angle)
    LineStrPoints(11) = -(Width / 2 + FlangeWidth)

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
        
    ProjVector.Set Cos(Angle), Sin(Angle), 0

    Set objRectFlangeIns = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oLineString, _
                                        ProjVector.X, ProjVector.Y, ProjVector.Z, _
                                        2 * (Thickness + InsulationThickness + CptOffset), True)

'Add the output into the Collection
    m_outputColl.AddOutput "RectFlangeIns", objRectFlangeIns
    Set objRectFlangeIns = Nothing
    Set ProjVector = Nothing
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    Set geomFactory = Nothing

    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Insulation representation", vbExclamation, "HRECTFFDInc definition"
        Debug.Assert False
End Sub
