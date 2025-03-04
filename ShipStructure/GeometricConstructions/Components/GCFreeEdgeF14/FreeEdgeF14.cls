VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdgeF14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const E_FAIL = &H80004005
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select line for first axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", False, 1)
    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select line for second axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", False, 1)
    Call pGeometricConstructionDefinition.AddInput("Circle", "Select circle", "IJCircle OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJCircle]", False, 1)
    '    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select line for first axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    '    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select line for second axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    '    Call pGeometricConstructionDefinition.AddInput("Circle", "Select circle", "IJCircle OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJCircle]", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", 8, 1, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", 8, 1, 0, 0, 0)
    
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oCSByLines1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines1 = oGCFactory.CreateEntity("CSByLines", pPOM, "001-CSByLines")
    oCSByLines1.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oCSByLines1.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oCSByLines1.Parameter("AxesRoles") = 1
    oCSByLines1.Parameter("CSOrientation") = 1
    oCSByLines1.Evaluate

    Dim oCSByLines2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines2 = oGCFactory.CreateEntity("CSByLines", pPOM, "002-CSByLines")
    oCSByLines2.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oCSByLines2.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oCSByLines2.Parameter("AxesRoles") = 1
    oCSByLines2.Parameter("CSOrientation") = 1
    oCSByLines2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("CutBack")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant4.Evaluate

    Dim oPointFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS5 = oGCFactory.CreateEntity("PointFromCS", pPOM, "005-PointFromCS")
    oPointFromCS5.Inputs("CoordinateSystem").Add oCSByLines2
    oPointFromCS5.Parameter("X") = 0#
    oPointFromCS5.Parameter("Y") = 3#
    oPointFromCS5.Parameter("Z") = 0#
    oPointFromCS5.Evaluate

    Dim oPlaneFromCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS6 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "006-PlaneFromCS")
    oPlaneFromCS6.Inputs("CoordinateSystem").Add oCSByLines2
    oPlaneFromCS6.Parameter("LookingAxis") = 3
    oPlaneFromCS6.Parameter("Offset") = 0#
    oPlaneFromCS6.Parameter("Range") = 10#
    oPlaneFromCS6.Evaluate
    
    Dim oLineFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7 = oGCFactory.CreateEntity("LineFromCS", pPOM, "007-LineFromCS")
    oLineFromCS7.Inputs("CoordinateSystem").Add oCSByLines2
    oLineFromCS7.Parameter("LookingAxis") = 3
    oLineFromCS7.Parameter("Length") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
    oLineFromCS7.Evaluate
    
    Dim oCurveByProjection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection8 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "008-CurveByProjection")
    oCurveByProjection8.Inputs("Curve").Add pGeometricConstruction.Input("Circle")
    oCurveByProjection8.Inputs("Surface").Add oPlaneFromCS6
    oCurveByProjection8.Inputs("Line").Add oLineFromCS7
    oCurveByProjection8.Parameter("TrackFlag") = 1
    oCurveByProjection8.Public = True
    oCurveByProjection8.Evaluate
    
    Dim oPointFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "009-PointFromCS")
    oPointFromCS9.Inputs("CoordinateSystem").Add oCSByLines2
    oPointFromCS9.Parameter("X") = 0#
    oPointFromCS9.Parameter("Y") = 0#
    oPointFromCS9.Parameter("Z") = 0#
    oPointFromCS9.Evaluate

    Dim oPointFromCS10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS10 = oGCFactory.CreateEntity("PointFromCS", pPOM, "010-PointFromCS")
    oPointFromCS10.Inputs("CoordinateSystem").Add oCSByLines1
    oPointFromCS10.Parameter("X") = 4#
    oPointFromCS10.Parameter("Y") = 0#
    oPointFromCS10.Parameter("Z") = 0#
    oPointFromCS10.Evaluate

    Dim oPointFromCS11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS11 = oGCFactory.CreateEntity("PointFromCS", pPOM, "011-PointFromCS")
    oPointFromCS11.Inputs("CoordinateSystem").Add oCSByLines2
    oPointFromCS11.Parameter("X") = 4#
    oPointFromCS11.Parameter("Y") = 0#
    oPointFromCS11.Parameter("Z") = 0#
    oPointFromCS11.Evaluate

    Dim oLineByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints12 = oGCFactory.CreateEntity("LineByPoints", pPOM, "012-LineByPoints")
    oLineByPoints12.Inputs("StartPoint").Add oPointFromCS11
    oLineByPoints12.Inputs("EndPoint").Add oPointFromCS10
    oLineByPoints12.Evaluate

    Dim oPointAtCurveMiddle13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle13 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "013-PointAtCurveMiddle")
    oPointAtCurveMiddle13.Inputs("Curve").Add oLineByPoints12
    oPointAtCurveMiddle13.Evaluate

    Dim oLineByPoints14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints14 = oGCFactory.CreateEntity("LineByPoints", pPOM, "014-LineByPoints")
    oLineByPoints14.Inputs("StartPoint").Add oPointFromCS9
    oLineByPoints14.Inputs("EndPoint").Add oPointAtCurveMiddle13
    oLineByPoints14.Evaluate

    Dim oCircleConcentric15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCircleConcentric15 = oGCFactory.CreateEntity("CircleConcentric", pPOM, "015-CircleConcentric")
    oCircleConcentric15.Inputs("Circle").Add oCurveByProjection8 ' pGeometricConstruction.Inputs("Circle").Item(1)
    oCircleConcentric15.Inputs("TrackPoint").Add oPointAtCurveMiddle13
    'oCircleConcentric15.Inputs("Parameter").Add oParamDistConstant4
    oCircleConcentric15.Parameter("Offset") = oParamDistConstant4
    oCircleConcentric15.Parameter("TrackFlag") = 1
    oCircleConcentric15.Evaluate

    Dim oPointByCurves16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves16 = oGCFactory.CreateEntity("PointByCurves", pPOM, "016-PointByCurves")
    oPointByCurves16.Inputs("Curve1").Add oLineByPoints14
    oPointByCurves16.Inputs("Curve2").Add oCircleConcentric15
    oPointByCurves16.Inputs("TrackPoint").Add oPointFromCS9
    oPointByCurves16.Parameter("TrackFlag") = 2
    oPointByCurves16.Evaluate

    Dim oLineAtAngle17 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle17 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "017-LineAtAngle")
    oLineAtAngle17.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oLineAtAngle17.Inputs("Point").Add oPointByCurves16
    oLineAtAngle17.Inputs("Plane").Add oPlaneFromCS6
    oLineAtAngle17.Inputs("TrackPoint").Add oPointFromCS11
    oLineAtAngle17.Parameter("Angle") = 0#
    oLineAtAngle17.Parameter("TrackFlag") = 1
    oLineAtAngle17.Evaluate

    Dim oLineAtAngle18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle18 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "018-LineAtAngle")
    oLineAtAngle18.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oLineAtAngle18.Inputs("Point").Add oPointByCurves16
    oLineAtAngle18.Inputs("Plane").Add oPlaneFromCS6
    oLineAtAngle18.Inputs("TrackPoint").Add oPointFromCS10
    oLineAtAngle18.Parameter("Angle") = 0#
    oLineAtAngle18.Parameter("TrackFlag") = 1
    oLineAtAngle18.Evaluate

    Dim oLineAtAngle19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle19 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "019-LineAtAngle")
    oLineAtAngle19.Inputs("Line").Add oLineByPoints14
    oLineAtAngle19.Inputs("Point").Add oPointByCurves16
    oLineAtAngle19.Inputs("Plane").Add oPlaneFromCS6
    oLineAtAngle19.Inputs("TrackPoint").Add oPointFromCS5
    oLineAtAngle19.Parameter("Angle") = 1.570796
    oLineAtAngle19.Parameter("TrackFlag") = 1
    oLineAtAngle19.Evaluate

    Dim oLineAtAngle20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle20 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "020-LineAtAngle")
    oLineAtAngle20.Inputs("Line").Add oLineByPoints14
    oLineAtAngle20.Inputs("Point").Add oPointByCurves16
    oLineAtAngle20.Inputs("Plane").Add oPlaneFromCS6
    oLineAtAngle20.Inputs("TrackPoint").Add oPointFromCS5
    oLineAtAngle20.Parameter("Angle") = -1.570796
    oLineAtAngle20.Parameter("TrackFlag") = 1
    oLineAtAngle20.Evaluate

    Dim oPointByCurves21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves21 = oGCFactory.CreateEntity("PointByCurves", pPOM, "021-PointByCurves")
    oPointByCurves21.Inputs("Curve1").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oPointByCurves21.Inputs("Curve2").Add oLineAtAngle19
    oPointByCurves21.Parameter("TrackFlag") = 1
    oPointByCurves21.Evaluate

    Dim oPointByCurves22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves22 = oGCFactory.CreateEntity("PointByCurves", pPOM, "022-PointByCurves")
    oPointByCurves22.Inputs("Curve1").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oPointByCurves22.Inputs("Curve2").Add oLineAtAngle20
    oPointByCurves22.Inputs("TrackPoint").Add oPointFromCS5
    oPointByCurves22.Parameter("TrackFlag") = 1
    oPointByCurves22.Evaluate

    Dim oLineByPoints23 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints23 = oGCFactory.CreateEntity("LineByPoints", pPOM, "023-LineByPoints")
    oLineByPoints23.Inputs("StartPoint").Add oPointByCurves16
    oLineByPoints23.Inputs("EndPoint").Add oPointByCurves22
    oLineByPoints23.Evaluate

    Dim oParamDistMeasureLength24 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength24 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "024-ParamDistMeasureLength")
    oParamDistMeasureLength24.Inputs("Curve").Add oLineByPoints23
    oParamDistMeasureLength24.Parameter("Value") = 0
    oParamDistMeasureLength24.Evaluate

    Dim oPointAlongCurve25 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve25 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "025-PointAlongCurve")
    oPointAlongCurve25.Inputs("Curve").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oPointAlongCurve25.Inputs("Point").Add oPointByCurves21
    oPointAlongCurve25.Inputs("TrackPoint").Add oPointFromCS11
    'oPointAlongCurve25.Inputs("Parameter").Add oParamDistMeasureLength24
    oPointAlongCurve25.Parameter("Distance") = oParamDistMeasureLength24
    oPointAlongCurve25.Parameter("TrackFlag") = 1
    oPointAlongCurve25.Evaluate

    Dim oPointAlongCurve26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve26 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "026-PointAlongCurve")
    oPointAlongCurve26.Inputs("Curve").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oPointAlongCurve26.Inputs("Point").Add oPointByCurves22
    oPointAlongCurve26.Inputs("TrackPoint").Add oPointFromCS10
    'oPointAlongCurve26.Inputs("Parameter").Add oParamDistMeasureLength24
    oPointAlongCurve26.Parameter("Distance") = oParamDistMeasureLength24
    oPointAlongCurve26.Parameter("TrackFlag") = 1
    oPointAlongCurve26.Evaluate

    Dim oArcByPoints27 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByPoints27 = oGCFactory.CreateEntity("ArcByPoints", pPOM, "027-ArcByPoints")
    oArcByPoints27.Inputs("StartPoint").Add oPointAlongCurve26
    oArcByPoints27.Inputs("MidPoint").Add oPointByCurves16
    oArcByPoints27.Inputs("EndPoint").Add oPointAlongCurve25
    oArcByPoints27.Parameter("PointLocation") = 1
    oArcByPoints27.Evaluate

    Dim oPointAlongCurve28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve28 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "028-PointAlongCurve")
    oPointAlongCurve28.Inputs("Curve").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oPointAlongCurve28.Inputs("Point").Add oPointAlongCurve25
    oPointAlongCurve28.Inputs("TrackPoint").Add oPointFromCS11
    'oPointAlongCurve28.Inputs("Parameter").Add oParamDistConstant3
    oPointAlongCurve28.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve28.Parameter("TrackFlag") = 1
    oPointAlongCurve28.Evaluate

    Dim oPointAlongCurve29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve29 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "029-PointAlongCurve")
    oPointAlongCurve29.Inputs("Curve").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oPointAlongCurve29.Inputs("Point").Add oPointAlongCurve26
    oPointAlongCurve29.Inputs("TrackPoint").Add oPointFromCS10
    'oPointAlongCurve29.Inputs("Parameter").Add oParamDistConstant3
    oPointAlongCurve29.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve29.Parameter("TrackFlag") = 1
    oPointAlongCurve29.Evaluate

    Dim oLineExtend30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend30 = oGCFactory.CreateEntity("LineExtend", pPOM, "030-LineExtend")
    oLineExtend30.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oLineExtend30.Inputs("Curve1").Add oArcByPoints27
    oLineExtend30.Parameter("Context1") = 2
    oLineExtend30.Parameter("Context2") = 1
    oLineExtend30.Evaluate

    Dim oLineExtend31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend31 = oGCFactory.CreateEntity("LineExtend", pPOM, "031-LineExtend")
    oLineExtend31.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oLineExtend31.Inputs("Curve1").Add oArcByPoints27
    oLineExtend31.Parameter("Context1") = 2
    oLineExtend31.Parameter("Context2") = 1
    oLineExtend31.Evaluate

    Dim oCpxStringByCurves32 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves32 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "032-CpxStringByCurves")
    oCpxStringByCurves32.Inputs("Curves").Add oLineExtend31
    oCpxStringByCurves32.Inputs("Curves").Add oArcByPoints27
    oCpxStringByCurves32.Inputs("Curves").Add oLineExtend30
    oCpxStringByCurves32.Evaluate
    
    Dim oSurfByLinearExtrusion33 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion33 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "033-SurfByLinearExtrusion")
    oSurfByLinearExtrusion33.Inputs("PlanarCrossSection").Add oCpxStringByCurves32
    oSurfByLinearExtrusion33.Inputs("ExtrusionLine").Add oLineFromCS7
    oSurfByLinearExtrusion33.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("PointX", 1) = oPointAlongCurve28.Output
    oGCMacro.Output("PointY", 1) = oPointAlongCurve29.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion33.Output
    Exit Sub
ErrorHandler:
    ' clear the error by doing nothing
End Sub
