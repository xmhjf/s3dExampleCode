VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdgeAround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select line for first axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select line for second axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Circle", "Select circle", "IJCircle OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJCircle]", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter("Side", "Side", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Inside", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Outside", 2)

    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim dDistanceForTrackPoint As Double
    Let dDistanceForTrackPoint = CDbl(pGeometricConstruction.Parameter("WeldToe")) + CDbl(pGeometricConstruction.Parameter("Offset"))
    
    Dim dDistanceForLineParallel As Double
    Let dDistanceForLineParallel = 3 + 2 * CDbl(pGeometricConstruction.Parameter("Offset"))
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oParamDistConstant1.Evaluate
    oParamDistConstant1.Public = True

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    oParamDistConstant2.Evaluate
    oParamDistConstant2.Public = True

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oParamDistConstant3.Evaluate
    oParamDistConstant3.Public = True

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Offset"))
    oParamDistConstant4.Evaluate
    oParamDistConstant4.Public = True
    
    Dim oCSByLines5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines5 = oGCFactory.CreateEntity("CSByLines", pPOM, "005-CSByLines")
    oCSByLines5.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oCSByLines5.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oCSByLines5.Parameter("AxesRoles") = 1
    oCSByLines5.Parameter("CSOrientation") = 1
    oCSByLines5.Evaluate
    oCSByLines5.Public = True

    Dim oPlaneFromCS5a As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS5a = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "006-PlaneFromCS")
    oPlaneFromCS5a.Inputs("CoordinateSystem").Add oCSByLines5
    oPlaneFromCS5a.Parameter("LookingAxis") = 3
    oPlaneFromCS5a.Parameter("Offset") = 0#
    oPlaneFromCS5a.Parameter("Range") = 5#
    oPlaneFromCS5a.Evaluate
    oPlaneFromCS5a.Public = True

    Dim oCSByLines6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines6 = oGCFactory.CreateEntity("CSByLines", pPOM, "007-CSByLines")
    oCSByLines6.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oCSByLines6.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oCSByLines6.Parameter("AxesRoles") = 1
    oCSByLines6.Parameter("CSOrientation") = 1
    oCSByLines6.Evaluate
    oCSByLines6.Public = True

    Dim oLineFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7 = oGCFactory.CreateEntity("LineFromCS", pPOM, "008-LineFromCS")
    oLineFromCS7.Inputs("CoordinateSystem").Add oCSByLines6
    oLineFromCS7.Parameter("LookingAxis") = 3
    oLineFromCS7.Parameter("Length") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
    oLineFromCS7.Evaluate
    oLineFromCS7.Public = True

    Dim oCurveByProjection7a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection7a = oGCFactory.CreateEntity("CurveByProjection", pPOM, "009-CurveByProjection")
    oCurveByProjection7a.Inputs("Curve").Add pGeometricConstruction.Inputs("Circle").Item(1)
    oCurveByProjection7a.Inputs("Surface").Add oPlaneFromCS5a
    oCurveByProjection7a.Inputs("Line").Add oLineFromCS7
    oCurveByProjection7a.Parameter("TrackFlag") = 1
    oCurveByProjection7a.Public = True
    oCurveByProjection7a.Evaluate

    Dim oPointFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS8 = oGCFactory.CreateEntity("PointFromCS", pPOM, "010-PointFromCS")
    oPointFromCS8.Inputs("CoordinateSystem").Add oCSByLines6
    oPointFromCS8.Parameter("X") = 3#
    If pGeometricConstruction.Parameter("Side") <> 2 Then
        oPointFromCS8.Parameter("Y") = dDistanceForTrackPoint
    Else
        oPointFromCS8.Parameter("Y") = -dDistanceForTrackPoint
    End If
    oPointFromCS8.Parameter("Z") = 0#
    oPointFromCS8.Evaluate
    oPointFromCS8.Public = True

    Dim oPointFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "011-PointFromCS")
    oPointFromCS9.Inputs("CoordinateSystem").Add oCSByLines6
    oPointFromCS9.Parameter("X") = 0#
    oPointFromCS9.Parameter("Y") = 0#
    oPointFromCS9.Parameter("Z") = 0#
    oPointFromCS9.Evaluate
    oPointFromCS9.Public = True

    Dim oPointFromCS10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS10 = oGCFactory.CreateEntity("PointFromCS", pPOM, "012-PointFromCS")
    oPointFromCS10.Inputs("CoordinateSystem").Add oCSByLines5
    oPointFromCS10.Parameter("X") = 3#
    If pGeometricConstruction.Parameter("Side") <> 2 Then
        oPointFromCS10.Parameter("Y") = dDistanceForTrackPoint
    Else
        oPointFromCS10.Parameter("Y") = -dDistanceForTrackPoint
    End If
    oPointFromCS10.Parameter("Z") = 0#
    oPointFromCS10.Evaluate
    oPointFromCS10.Public = True

    Dim oLineParallel11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel11 = oGCFactory.CreateEntity("LineParallel", pPOM, "013-LineParallel")
    oLineParallel11.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oLineParallel11.Inputs("Plane").Add oPlaneFromCS5a
    oLineParallel11.Inputs("TrackPoint").Add oPointFromCS8
    oLineParallel11.Parameter("Distance") = oParamDistConstant3
    oLineParallel11.Parameter("TrackFlag") = 1
    oLineParallel11.Evaluate
    oLineParallel11.Public = True

    Dim oLineParallel12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel12 = oGCFactory.CreateEntity("LineParallel", pPOM, "014-LineParallel")
    oLineParallel12.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oLineParallel12.Inputs("Plane").Add oPlaneFromCS5a
    oLineParallel12.Inputs("TrackPoint").Add oPointFromCS10
    oLineParallel12.Parameter("Distance") = oParamDistConstant3
    oLineParallel12.Parameter("TrackFlag") = 1
    oLineParallel12.Evaluate
    oLineParallel12.Public = True

    Dim oPointAtCurveExtremity13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity13 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "015-PointAtCurveExtremity")
    oPointAtCurveExtremity13.Inputs("Curve").Add oLineParallel11
    oPointAtCurveExtremity13.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity13.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity13.Evaluate
    oPointAtCurveExtremity13.Public = True

    Dim oLineParallel14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel14 = oGCFactory.CreateEntity("LineParallel", pPOM, "016-LineParallel")
    oLineParallel14.Inputs("Line").Add oLineParallel11
    oLineParallel14.Inputs("Plane").Add oPlaneFromCS5a
    oLineParallel14.Inputs("TrackPoint").Add oPointFromCS8
    oLineParallel14.Parameter("Distance") = dDistanceForLineParallel
    oLineParallel14.Parameter("TrackFlag") = 1
    oLineParallel14.Evaluate
    oLineParallel14.Public = True

    Dim oCircleConcentric15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCircleConcentric15 = oGCFactory.CreateEntity("CircleConcentric", pPOM, "017-CircleConcentric")
    oCircleConcentric15.Inputs("Circle").Add oCurveByProjection7a
    oCircleConcentric15.Inputs("TrackPoint").Add oPointFromCS8
    oCircleConcentric15.Parameter("Offset") = oParamDistConstant4
    oCircleConcentric15.Parameter("TrackFlag") = 1
    oCircleConcentric15.Evaluate
    oCircleConcentric15.Public = True

    Dim oPointAtCurveExtremity16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity16 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "018-PointAtCurveExtremity")
    oPointAtCurveExtremity16.Inputs("Curve").Add oLineParallel12
    oPointAtCurveExtremity16.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity16.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity16.Evaluate
    oPointAtCurveExtremity16.Public = True

    Dim oLineParallel17 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel17 = oGCFactory.CreateEntity("LineParallel", pPOM, "019-LineParallel")
    oLineParallel17.Inputs("Line").Add oLineParallel12
    oLineParallel17.Inputs("Plane").Add oPlaneFromCS5a
    oLineParallel17.Inputs("TrackPoint").Add oPointFromCS10
    oLineParallel17.Parameter("Distance") = dDistanceForLineParallel
    oLineParallel17.Parameter("TrackFlag") = 1
    oLineParallel17.Evaluate
    oLineParallel17.Public = True

    Dim oLineExtend18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend18 = oGCFactory.CreateEntity("LineExtend", pPOM, "020-LineExtend")
    oLineExtend18.Inputs("Line").Add oLineParallel17
    oLineExtend18.Inputs("Curve1").Add oLineParallel14
    oLineExtend18.Parameter("Context1") = 1
    oLineExtend18.Parameter("Context2") = 4
    oLineExtend18.Evaluate
    oLineExtend18.Public = True

    Dim oLineExtend19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend19 = oGCFactory.CreateEntity("LineExtend", pPOM, "021-LineExtend")
    oLineExtend19.Inputs("Line").Add oLineParallel14
    oLineExtend19.Inputs("Curve1").Add oLineExtend18
    oLineExtend19.Parameter("Context1") = 1
    oLineExtend19.Parameter("Context2") = 4
    oLineExtend19.Evaluate
    oLineExtend19.Public = True

    Dim oPointByCurves20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves20 = oGCFactory.CreateEntity("PointByCurves", pPOM, "022-PointByCurves")
    oPointByCurves20.Inputs("Curve1").Add oLineExtend18
    oPointByCurves20.Inputs("Curve2").Add oLineExtend19
    oPointByCurves20.Parameter("TrackFlag") = 1
    oPointByCurves20.Evaluate
    oPointByCurves20.Public = True

    Dim oPointByCurves21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves21 = oGCFactory.CreateEntity("PointByCurves", pPOM, "023-PointByCurves")
    oPointByCurves21.Inputs("Curve1").Add oLineParallel11
    oPointByCurves21.Inputs("Curve2").Add oCircleConcentric15
    oPointByCurves21.Inputs("TrackPoint").Add oPointByCurves20
    oPointByCurves21.Parameter("TrackFlag") = 1
    oPointByCurves21.Evaluate
    oPointByCurves21.Public = True

    Dim oLineByPoints22 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints22 = oGCFactory.CreateEntity("LineByPoints", pPOM, "024-LineByPoints")
    oLineByPoints22.Inputs("StartPoint").Add oPointFromCS9
    oLineByPoints22.Inputs("EndPoint").Add oPointByCurves20
    oLineByPoints22.Evaluate
    oLineByPoints22.Public = True

    Dim oPointByCurves23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves23 = oGCFactory.CreateEntity("PointByCurves", pPOM, "025-PointByCurves")
    oPointByCurves23.Inputs("Curve1").Add oLineParallel12
    oPointByCurves23.Inputs("Curve2").Add oCircleConcentric15
    oPointByCurves23.Inputs("TrackPoint").Add oPointByCurves20
    oPointByCurves23.Parameter("TrackFlag") = 1
    oPointByCurves23.Evaluate
    oPointByCurves23.Public = True

    Dim oLineByPoints24 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints24 = oGCFactory.CreateEntity("LineByPoints", pPOM, "026-LineByPoints")
    oLineByPoints24.Inputs("StartPoint").Add oPointByCurves21
    oLineByPoints24.Inputs("EndPoint").Add oPointAtCurveExtremity13
    oLineByPoints24.Evaluate
    oLineByPoints24.Public = True

    Dim oPointByCurves25 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves25 = oGCFactory.CreateEntity("PointByCurves", pPOM, "027-PointByCurves")
    oPointByCurves25.Inputs("Curve1").Add oCircleConcentric15
    oPointByCurves25.Inputs("Curve2").Add oLineByPoints22
    oPointByCurves25.Inputs("TrackPoint").Add oPointByCurves20
    oPointByCurves25.Parameter("TrackFlag") = 1
    oPointByCurves25.Evaluate
    oPointByCurves25.Public = True

    Dim oLineByPoints26 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints26 = oGCFactory.CreateEntity("LineByPoints", pPOM, "028-LineByPoints")
    oLineByPoints26.Inputs("StartPoint").Add oPointByCurves23
    oLineByPoints26.Inputs("EndPoint").Add oPointAtCurveExtremity16
    oLineByPoints26.Evaluate
    oLineByPoints26.Public = True

    Dim oArcByPoints27 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByPoints27 = oGCFactory.CreateEntity("ArcByPoints", pPOM, "029-ArcByPoints")
    oArcByPoints27.Inputs("StartPoint").Add oPointByCurves21
    oArcByPoints27.Inputs("MidPoint").Add oPointByCurves25
    oArcByPoints27.Inputs("EndPoint").Add oPointByCurves23
    oArcByPoints27.Parameter("PointLocation") = 1
    oArcByPoints27.Evaluate
    oArcByPoints27.Public = True

    Dim oArcByFillet28 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet28 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "030-ArcByFillet")
    oArcByFillet28.Inputs("Curve1").Add oLineByPoints26
    oArcByFillet28.Inputs("Curve2").Add oArcByPoints27
    oArcByFillet28.Parameter("Radius") = oParamDistConstant2
    oArcByFillet28.Evaluate
    oArcByFillet28.Public = True

    Dim oArcByFillet29 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet29 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "031-ArcByFillet")
    oArcByFillet29.Inputs("Curve1").Add oLineParallel11
    oArcByFillet29.Inputs("Curve2").Add oArcByPoints27
    oArcByFillet29.Parameter("Radius") = oParamDistConstant2
    oArcByFillet29.Evaluate
    oArcByFillet29.Public = True

    Dim oPointAtCurveExtremity30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity30 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "032-PointAtCurveExtremity")
    oPointAtCurveExtremity30.Inputs("Curve").Add oArcByFillet29
    oPointAtCurveExtremity30.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity30.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity30.Evaluate
    oPointAtCurveExtremity30.Public = True

    Dim oPointAtCurveExtremity31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity31 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "033-PointAtCurveExtremity")
    oPointAtCurveExtremity31.Inputs("Curve").Add oArcByFillet29
    oPointAtCurveExtremity31.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity31.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity31.Evaluate
    oPointAtCurveExtremity31.Public = True

    Dim oPointAtCurveExtremity32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity32 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "034-PointAtCurveExtremity")
    oPointAtCurveExtremity32.Inputs("Curve").Add oArcByFillet28
    oPointAtCurveExtremity32.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity32.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity32.Evaluate
    oPointAtCurveExtremity32.Public = True

    Dim oPointAtCurveExtremity33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity33 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "035-PointAtCurveExtremity")
    oPointAtCurveExtremity33.Inputs("Curve").Add oArcByFillet28
    oPointAtCurveExtremity33.Inputs("TrackPoint").Add oPointFromCS9
    oPointAtCurveExtremity33.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity33.Evaluate
    oPointAtCurveExtremity33.Public = True

    Dim oPointAlongCurve34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve34 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "036-PointAlongCurve")
    oPointAlongCurve34.Inputs("Curve").Add oLineParallel11
    oPointAlongCurve34.Inputs("Point").Add oPointAtCurveExtremity30
    oPointAlongCurve34.Inputs("TrackPoint").Add oPointAtCurveExtremity13
    oPointAlongCurve34.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve34.Parameter("TrackFlag") = 1
    oPointAlongCurve34.Evaluate
    oPointAlongCurve34.Public = True

    Dim oPointAlongCurve35 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve35 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "037-PointAlongCurve")
    oPointAlongCurve35.Inputs("Curve").Add oLineParallel12
    oPointAlongCurve35.Inputs("Point").Add oPointAtCurveExtremity33
    oPointAlongCurve35.Inputs("TrackPoint").Add oPointAtCurveExtremity16
    oPointAlongCurve35.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve35.Parameter("TrackFlag") = 1
    oPointAlongCurve35.Evaluate
    oPointAlongCurve35.Public = True

    Dim oArcByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByPoints36 = oGCFactory.CreateEntity("ArcByPoints", pPOM, "038-ArcByPoints")
    oArcByPoints36.Inputs("StartPoint").Add oPointAtCurveExtremity31
    oArcByPoints36.Inputs("MidPoint").Add oPointByCurves25
    oArcByPoints36.Inputs("EndPoint").Add oPointAtCurveExtremity32
    oArcByPoints36.Parameter("PointLocation") = 1
    oArcByPoints36.Evaluate
    oArcByPoints36.Public = True

    Dim oLineByPoints37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints37 = oGCFactory.CreateEntity("LineByPoints", pPOM, "039-LineByPoints")
    oLineByPoints37.Inputs("StartPoint").Add oPointAtCurveExtremity30
    oLineByPoints37.Inputs("EndPoint").Add oPointAtCurveExtremity13
    oLineByPoints37.Evaluate
    oLineByPoints37.Public = True

    Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "040-LineByPoints")
    oLineByPoints38.Inputs("StartPoint").Add oPointAtCurveExtremity33
    oLineByPoints38.Inputs("EndPoint").Add oPointAtCurveExtremity16
    oLineByPoints38.Evaluate
    oLineByPoints38.Public = True

    Dim oCpxStringByCurves39 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves39 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "041-CpxStringByCurves")
    oCpxStringByCurves39.Inputs("Curves").Add oLineByPoints37
    oCpxStringByCurves39.Inputs("Curves").Add oArcByFillet29
    oCpxStringByCurves39.Inputs("Curves").Add oArcByPoints36
    oCpxStringByCurves39.Inputs("Curves").Add oArcByFillet28
    oCpxStringByCurves39.Inputs("Curves").Add oLineByPoints38
    oCpxStringByCurves39.Evaluate
    oCpxStringByCurves39.Public = True

    Dim oSurfByLinearExtrusion40 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion40 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "042-SurfByLinearExtrusion")
    oSurfByLinearExtrusion40.Inputs("PlanarCrossSection").Add oCpxStringByCurves39
    oSurfByLinearExtrusion40.Inputs("ExtrusionLine").Add oLineFromCS7
    oSurfByLinearExtrusion40.Evaluate
    oSurfByLinearExtrusion40.Public = True

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion40.Output
    oGCMacro.Output("PointX", 1) = oPointAlongCurve34.Output
    oGCMacro.Output("PointY", 1) = oPointAlongCurve35.Output
End Sub
