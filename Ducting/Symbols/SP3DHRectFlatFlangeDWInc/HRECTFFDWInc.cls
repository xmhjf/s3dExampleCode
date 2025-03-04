VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HRECTFFDWInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    HRECTFFDWInc.cls
'
'   ProgID:         SP3DHRectFlatFlangeDWInc.HRECTFFDWInc
'   Author:         svsmylav
'   Creation Date:  Thursday, Dec 12 2002
'
'Description
'    Definition of HVAC Rectangular Flat Flange Where the Depth and Width are inclined.
'    Corners and co-ordinate system of the Rectangular duct are as follows:
'        Y
'        |
'    1-------2
'    |   |   |
'    |   O --|----> Z
'    |       |
'    4-------3
'
'Notes
'   This symbol is made on the basis of RectFlatFlange
'
'History:

'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007     PS      CR-132021  Modified the default value of Angle and Rotation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection
    
Const NEGLIGIBLE_THICKNESS = 0.0001

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRectFlatFlangeDWInc.HRECTFFDWInc"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  Dim oSymbolDefinition As imssymbolentities.DSymbolDefinition
  Dim oSymbolFactory As New imssymbolentities.DSymbolEntitiesFactory
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRectFlatFlangeDWInc.HRECTFFDWInc"
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


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As imssymbolentities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As imssymbolentities.IJDInputs
    Set InputsIf = pSymbolDefinition

    'Following indented lines will be removed very soon. This has been
    'included to improve the performance. These lines will be removed as soon as
    'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.
        ' set the CMCacheForPart method in the definition
        
    ' Create a new input by new operator
    Dim Inputs(1 To 10) As imssymbolentities.DInput
    Dim Index As Integer
   
    ' Create a defaultValue
    Dim PC As imssymbolentities.DParameterContent
    Set PC = New imssymbolentities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
   
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    Set Inputs(2) = New imssymbolentities.DInput
    Inputs(2).Name = "Width"
    Inputs(2).Description = "Width of the Rect Duct"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.3
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New imssymbolentities.DInput
    Inputs(3).Name = "Depth"
    Inputs(3).Description = "Depth of the Rect Duct"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2
    Inputs(3).DefaultParameterValue = PC
    
    Set Inputs(4) = New imssymbolentities.DInput
    Inputs(4).Name = "Clearance"
    Inputs(4).Description = "Clearance between ducts"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.008
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New imssymbolentities.DInput
    Inputs(5).Name = "Thickness"
    Inputs(5).Description = "Thickness of the Flat Flange"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.06
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New imssymbolentities.DInput
    Inputs(6).Name = "FlangeWidth"
    Inputs(6).Description = "Width of the Flat Flange"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.04
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New imssymbolentities.DInput
    Inputs(7).Name = "CompressedGasket"
    Inputs(7).Description = "Compressed Gasket Thickness"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.004
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New imssymbolentities.DInput
    Inputs(8).Name = "Angle"
    Inputs(8).Description = "Depth Inclination Angle"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = (1 / 15) * Atn(1)
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New imssymbolentities.DInput
    Inputs(9).Name = "Rotation"
    Inputs(9).Description = "Rotation"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = (4 / 15) * Atn(1)
    Inputs(9).DefaultParameterValue = PC
    
    Set Inputs(10) = New imssymbolentities.DInput
    Inputs(10).Name = "InsulationThickness"
    Inputs(10).Description = "Insulation Thickness"
    Inputs(10).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.025
    Inputs(10).DefaultParameterValue = PC
    
    For Index = 2 To 10
        InputsIf.SetInput Inputs(Index), Index
        Set Inputs(Index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 4) As imssymbolentities.DOutput
    
    For Index = 1 To 4
        Set O(Index) = New imssymbolentities.DOutput
        O(Index).Properties = 0
    Next

    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Flat Flange"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Flat Flange"
    
    O(3).Name = "RectFlange"
    O(3).Description = "Dumb graphics of Flat Flange"
    
    O(4).Name = "RectFlangeIns"
    O(4).Description = "Dumb graphics of Insulation for the Flat Flanges"
    
    'Define the representation "Physical"
    Dim rep1 As imssymbolentities.DRepresentation
    Set rep1 = New imssymbolentities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the HVAC Rect Flat Flange with Depth Inclination"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As imssymbolentities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For Index = 1 To 3
        oRepPhysicalOutputs.SetOutput O(Index)
        Set O(Index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As imssymbolentities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1
   
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRectFlatFlangeDWInc.HRECTFFDWInc"

    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As imssymbolentities.IJDRepresentationEvaluations
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

    Dim oRepInsulationOutputs As imssymbolentities.IJDOutputs
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
    InsulationRepEval.ProgId = "SP3DHRectFlatFlangeDWInc.HRECTFFDWInc"
    
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
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFFDWInc
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
    Dim Rotation As Double
    Dim InsulationThickness As Double
    
    Dim oPart As PartFacelets.IJDPart
    Dim objRectFlange As Object
    
    Set oPart = arrayOfInputs(1)
'   Assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Clearance = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
    CompressedGasket = arrayOfInputs(7)
    Angle = arrayOfInputs(8)        'Measured in anti-clockwise direction from positive Y-axis
    Rotation = arrayOfInputs(9)     'Measured in anti-clockwise direction from positive Z-axis
    InsulationThickness = arrayOfInputs(10)
    
       
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
    
'   Dumb graphics is used to place single flange. End preparartion for port1 is 11 and port2 is 301,
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

 'Insert code for the Dumb graphics of Rectangular flange 1
    CptOffset = CompressedGasket / 2
    
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim ProjVector As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
           
'   Approach used to create inclined flanges:
'       i) Consider mating face of inclined flange on left hand side. Intersection of the inclined
'          plane with the four planes(Walls) of the duct generates four-sided polygon.
'          Compute four corner points of the polygon and also the four mid-points of each side.
'       ii) To provide flangeWidth above the sides of polygon, compute change needed in direction
'           vectors at four mid-points and apply these changes to the corresponding corner points.
'       iii) Prepare line string from the above and project it along the normal of the
'            inclined plane to create the flange

'Create Line String
    Dim transMat        As New AutoMath.DT4x4
    Dim TanA As Double
    Dim TanR As Double
    Dim HD As Double
    Dim HW As Double

    TanA = Tan(Angle)   'Rotation about Z axis
    TanR = Tan(Rotation) 'Rotation about Y axis
    
    HD = Depth / 2
    HW = Width / 2
    
'   Prepare intersection points of inclined-plane(i.e. verticle plane rotated through
'   Rotation, Angle about Y-axis and  Z-axis respectively) with the four planes(walls) of duct.
    LineStrPoints(0) = -CptOffset - HD * TanA - HW * TanR
    LineStrPoints(1) = HD
    LineStrPoints(2) = -HW

    LineStrPoints(3) = -CptOffset - HD * TanA + HW * TanR
    LineStrPoints(4) = HD
    LineStrPoints(5) = HW

    LineStrPoints(6) = -CptOffset + HD * TanA + HW * TanR
    LineStrPoints(7) = -HD
    LineStrPoints(8) = HW

    LineStrPoints(9) = -CptOffset + HD * TanA - HW * TanR
    LineStrPoints(10) = -HD
    LineStrPoints(11) = -HW

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)

    Dim Mid12 As New AutoMath.DVector
    Dim Mid23 As New AutoMath.DVector
    Dim Mid34 As New AutoMath.DVector
    Dim Mid41 As New AutoMath.DVector
    Dim tmpV1 As New AutoMath.DVector
    Dim tmpV2 As New AutoMath.DVector
    Dim tmpV3 As New AutoMath.DVector
    Dim tmpV4 As New AutoMath.DVector
    Dim VecLen As Double
            
'   Deduct (-CptOffset, 0,0) from each Mid-point to get corresponding vector
'   Note: -(-CptOffset) is + CptOffset, and deduction of '0' is not done
    Mid12.Set (LineStrPoints(0) + LineStrPoints(3)) / 2 + CptOffset, _
                (LineStrPoints(1) + LineStrPoints(4)) / 2, _
                (LineStrPoints(2) + LineStrPoints(5)) / 2
    tmpV1.Set Mid12.X, Mid12.Y, Mid12.Z
    VecLen = Mid12.Length
    Mid12.[Scale] ((VecLen + FlangeWidth) / VecLen)
        
    Mid23.Set (LineStrPoints(3) + LineStrPoints(6)) / 2 + CptOffset, _
                (LineStrPoints(4) + LineStrPoints(7)) / 2, _
                (LineStrPoints(5) + LineStrPoints(8)) / 2
    tmpV2.Set Mid23.X, Mid23.Y, Mid23.Z
    VecLen = Mid23.Length
    Mid23.[Scale] ((VecLen + FlangeWidth) / VecLen)
    
    Mid34.Set (LineStrPoints(6) + LineStrPoints(9)) / 2 + CptOffset, _
                (LineStrPoints(7) + LineStrPoints(10)) / 2, _
                (LineStrPoints(8) + LineStrPoints(11)) / 2
    tmpV3.Set Mid34.X, Mid34.Y, Mid34.Z
    VecLen = Mid34.Length
    Mid34.[Scale] ((VecLen + FlangeWidth) / VecLen)
    
    Mid41.Set (LineStrPoints(0) + LineStrPoints(9)) / 2 + CptOffset, _
                (LineStrPoints(1) + LineStrPoints(10)) / 2, _
                (LineStrPoints(2) + LineStrPoints(11)) / 2
    tmpV4.Set Mid41.X, Mid41.Y, Mid41.Z
    VecLen = Mid41.Length
    Mid41.[Scale] ((VecLen + FlangeWidth) / VecLen)
    
'   Modify point 1
    LineStrPoints(0) = LineStrPoints(0) + (Mid12.X - tmpV1.X) + (Mid41.X - tmpV4.X)
    LineStrPoints(1) = LineStrPoints(1) + (Mid12.Y - tmpV1.Y) + (Mid41.Y - tmpV4.Y)
    LineStrPoints(2) = LineStrPoints(2) + (Mid12.Z - tmpV1.Z) + (Mid41.Z - tmpV4.Z)
    
'   Modify point 2
    LineStrPoints(3) = LineStrPoints(3) + (Mid12.X - tmpV1.X) + (Mid23.X - tmpV2.X)
    LineStrPoints(4) = LineStrPoints(4) + (Mid12.Y - tmpV1.Y) + (Mid23.Y - tmpV2.Y)
    LineStrPoints(5) = LineStrPoints(5) + (Mid12.Z - tmpV1.Z) + (Mid23.Z - tmpV2.Z)
    
'   Modify point 3
    LineStrPoints(6) = LineStrPoints(6) + (Mid23.X - tmpV2.X) + (Mid34.X - tmpV3.X)
    LineStrPoints(7) = LineStrPoints(7) + (Mid23.Y - tmpV2.Y) + (Mid34.Y - tmpV3.Y)
    LineStrPoints(8) = LineStrPoints(8) + (Mid23.Z - tmpV2.Z) + (Mid34.Z - tmpV3.Z)
    
'   Modify point 4
    LineStrPoints(9) = LineStrPoints(9) + (Mid34.X - tmpV3.X) + (Mid41.X - tmpV4.X)
    LineStrPoints(10) = LineStrPoints(10) + (Mid34.Y - tmpV3.Y) + (Mid41.Y - tmpV4.Y)
    LineStrPoints(11) = LineStrPoints(11) + (Mid34.Z - tmpV3.Z) + (Mid41.Z - tmpV4.Z)
    
    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
    
'   ProjVector is Cross-product of Mid12 and Mid41
    dir.Set Mid12.X, Mid12.Y, Mid12.Z
    Set ProjVector = dir.Cross(Mid41)
    
    Set objRectFlange = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oLineString, _
                                        ProjVector.X, ProjVector.Y, ProjVector.Z, Thickness, True)

'   Add the output into the Collection
    m_outputColl.AddOutput "RectFlange", objRectFlange
    Set objRectFlange = Nothing

    Set ProjVector = Nothing
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    Set transMat = Nothing
    Set Mid12 = Nothing
    Set Mid23 = Nothing
    Set Mid34 = Nothing
    Set Mid41 = Nothing
    Set tmpV1 = Nothing
    Set tmpV2 = Nothing
    Set tmpV3 = Nothing
    Set tmpV4 = Nothing

    Set geomFactory = Nothing
    Set dir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRECTFFDWInc definition"
        Debug.Assert False
End Sub

'=========================================================================
'CREATION OF INSULATION REPRESENTATION OF HRECTFFDWInc
'=========================================================================

Private Sub InsulationRep(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    
    Dim Width As Double
    Dim Depth As Double
    Dim Clearance As Double
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim CompressedGasket As Double
    Dim Angle As Double
    Dim Rotation As Double
    Dim InsulationThickness As Double
    
    Dim oPart As PartFacelets.IJDPart
    Dim objRectFlangeIns As Object
    
    Set oPart = arrayOfInputs(1)
'   assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Clearance = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
    CompressedGasket = arrayOfInputs(7)
    Angle = arrayOfInputs(8)    'Measured in anti-clockwise direction from positive Y-axis
    Rotation = arrayOfInputs(9) 'Measured in anti-clockwise direction from positive Z-axis
    InsulationThickness = arrayOfInputs(10)

    
'Insert code for the Dumb graphics of Insulation for Rect Flanges with Depth Inclination
    Dim CptOffset As Double
    CptOffset = CompressedGasket / 2
    
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim ProjVector As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
    Dim dir As New AutoMath.DVector

'   Approach used to create insulation for inclined flanges is same as the approach described
'   in 'Physical' procedure, However mating face of right side flange is considered.
'   FlangeWidth and InsulationThickness are to be considered for the modification of corner points.

'   Create Line String
    Dim transMat        As New AutoMath.DT4x4
    Dim TanA As Double
    Dim TanR As Double
    Dim HD As Double
    Dim HW As Double

    TanA = Tan(Angle)    'Rotation about Z axis
    TanR = Tan(Rotation) 'Rotation about Y axis
    
    HD = Depth / 2
    HW = Width / 2

'   Prepare intersection points of inclined-plane(i.e. verticle plane rotated through
'   Rotation, Angle about Y-axis and  Z-axis respectively) with the four planes(walls) of duct.
    LineStrPoints(0) = CptOffset - HD * TanA - HW * TanR
    LineStrPoints(1) = HD
    LineStrPoints(2) = -HW

    LineStrPoints(3) = CptOffset - HD * TanA + HW * TanR
    LineStrPoints(4) = HD
    LineStrPoints(5) = HW

    LineStrPoints(6) = CptOffset + HD * TanA + HW * TanR
    LineStrPoints(7) = -HD
    LineStrPoints(8) = HW

    LineStrPoints(9) = CptOffset + HD * TanA - HW * TanR
    LineStrPoints(10) = -HD
    LineStrPoints(11) = -HW

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)
        
'   Scale-up the line string so that the flange width and Insulation Thicknessis are
'   added all around

    Dim Mid12 As New AutoMath.DVector
    Dim Mid23 As New AutoMath.DVector
    Dim Mid34 As New AutoMath.DVector
    Dim Mid41 As New AutoMath.DVector
    Dim tmpV1 As New AutoMath.DVector
    Dim tmpV2 As New AutoMath.DVector
    Dim tmpV3 As New AutoMath.DVector
    Dim tmpV4 As New AutoMath.DVector
    Dim VecLen As Double
            
    Mid12.Set (LineStrPoints(0) + LineStrPoints(3)) / 2 - CptOffset, _
                (LineStrPoints(1) + LineStrPoints(4)) / 2, _
                (LineStrPoints(2) + LineStrPoints(5)) / 2
    tmpV1.Set Mid12.X, Mid12.Y, Mid12.Z
    VecLen = Mid12.Length
    Mid12.[Scale] ((VecLen + FlangeWidth + InsulationThickness) / VecLen)
    
    
    Mid23.Set (LineStrPoints(3) + LineStrPoints(6)) / 2 - CptOffset, _
                (LineStrPoints(4) + LineStrPoints(7)) / 2, _
                (LineStrPoints(5) + LineStrPoints(8)) / 2
    tmpV2.Set Mid23.X, Mid23.Y, Mid23.Z
    VecLen = Mid23.Length
    Mid23.[Scale] ((VecLen + FlangeWidth + InsulationThickness) / VecLen)
    
    Mid34.Set (LineStrPoints(6) + LineStrPoints(9)) / 2 - CptOffset, _
                (LineStrPoints(7) + LineStrPoints(10)) / 2, _
                (LineStrPoints(8) + LineStrPoints(11)) / 2
    tmpV3.Set Mid34.X, Mid34.Y, Mid34.Z
    VecLen = Mid34.Length
    Mid34.[Scale] ((VecLen + FlangeWidth + InsulationThickness) / VecLen)
    
    Mid41.Set (LineStrPoints(0) + LineStrPoints(9)) / 2 - CptOffset, _
                (LineStrPoints(1) + LineStrPoints(10)) / 2, _
                (LineStrPoints(2) + LineStrPoints(11)) / 2
    tmpV4.Set Mid41.X, Mid41.Y, Mid41.Z
    VecLen = Mid41.Length
    Mid41.[Scale] ((VecLen + FlangeWidth + InsulationThickness) / VecLen)
    
'   Modify point 1
    LineStrPoints(0) = LineStrPoints(0) + (Mid12.X - tmpV1.X) + (Mid41.X - tmpV4.X)
    LineStrPoints(1) = LineStrPoints(1) + (Mid12.Y - tmpV1.Y) + (Mid41.Y - tmpV4.Y)
    LineStrPoints(2) = LineStrPoints(2) + (Mid12.Z - tmpV1.Z) + (Mid41.Z - tmpV4.Z)
    
'   Modify point 2
    LineStrPoints(3) = LineStrPoints(3) + (Mid12.X - tmpV1.X) + (Mid23.X - tmpV2.X)
    LineStrPoints(4) = LineStrPoints(4) + (Mid12.Y - tmpV1.Y) + (Mid23.Y - tmpV2.Y)
    LineStrPoints(5) = LineStrPoints(5) + (Mid12.Z - tmpV1.Z) + (Mid23.Z - tmpV2.Z)
    
'   Modify point 3
    LineStrPoints(6) = LineStrPoints(6) + (Mid23.X - tmpV2.X) + (Mid34.X - tmpV3.X)
    LineStrPoints(7) = LineStrPoints(7) + (Mid23.Y - tmpV2.Y) + (Mid34.Y - tmpV3.Y)
    LineStrPoints(8) = LineStrPoints(8) + (Mid23.Z - tmpV2.Z) + (Mid34.Z - tmpV3.Z)
    
'   Modify point 4
    LineStrPoints(9) = LineStrPoints(9) + (Mid34.X - tmpV3.X) + (Mid41.X - tmpV4.X)
    LineStrPoints(10) = LineStrPoints(10) + (Mid34.Y - tmpV3.Y) + (Mid41.Y - tmpV4.Y)
    LineStrPoints(11) = LineStrPoints(11) + (Mid34.Z - tmpV3.Z) + (Mid41.Z - tmpV4.Z)
    
    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
    
'   ProjVector is Cross-product of Mid41 and Mid12
    dir.Set Mid41.X, Mid41.Y, Mid41.Z
    Set ProjVector = dir.Cross(Mid12)
    ProjVector.Length = 1       'Get unit vector
    
'   Apply translation to oLineString so that projection can start from the right most face
    transMat.LoadIdentity
    dir.Set ProjVector.X * (Thickness + InsulationThickness), _
                ProjVector.Y * (Thickness + InsulationThickness), _
                ProjVector.Z * (Thickness + InsulationThickness)
    transMat.Translate dir
    oLineString.Transform transMat
    ProjVector.Set -ProjVector.X, -ProjVector.Y, -ProjVector.Z 'Specify direction of projection
    Set objRectFlangeIns = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oLineString, _
                                        ProjVector.X, ProjVector.Y, ProjVector.Z, _
                                        2 * (Thickness + InsulationThickness + CptOffset), True)

'   Add the output into the Collection
    m_outputColl.AddOutput "RectFlangeIns", objRectFlangeIns
    Set objRectFlangeIns = Nothing

    Set ProjVector = Nothing
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    Set transMat = Nothing
    Set Mid12 = Nothing
    Set Mid23 = Nothing
    Set Mid34 = Nothing
    Set Mid41 = Nothing
    Set tmpV1 = Nothing
    Set tmpV2 = Nothing
    Set tmpV3 = Nothing
    Set tmpV4 = Nothing
    Set geomFactory = Nothing
    Set dir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Debug.Print "Error encountered while computing Insulation representation", vbExclamation, "HRECTFFDWInc definition"
        Debug.Assert False
End Sub
