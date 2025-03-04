VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnipedSurfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Support", "Select sketching plane (optional)", "IJPlane OR IJSurfaceBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJPlane]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CurveLeft", "Select line left", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CurveRight", "Select line right", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("MemberAxis", "Select member axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("TrimmingPoint", "Select trimming point", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", 8, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("SlopeRatio", "Slope ratio", 8, UNIT_SCALAR, 0, 0, 0, 1.01)
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", 8, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "Ribbon height", 8, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)

    Call pGeometricConstructionDefinition.AddOutput(6, "SurfaceLeft")
    Call pGeometricConstructionDefinition.AddOutput(6, "SurfaceRight")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oPointAtCurveStart2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart2 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "001-PointAtCurveStart")
    oPointAtCurveStart2.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oPointAtCurveStart2.Evaluate
    
    Dim oParamDouble4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble4 = oGCFactory.CreateEntity("ParamDouble", pPOM, "002-ParamDouble")
    oParamDouble4.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("SlopeRatio"))
    oParamDouble4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oParamDistConstant5.Evaluate

    ' project trimming point on member axis
    Dim oPointAlongCurve0 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve0 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "004-PointAlongCurve")
    oPointAlongCurve0.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oPointAlongCurve0.Inputs("Point").Add pGeometricConstruction.Inputs("TrimmingPoint").Item(1)
    oPointAlongCurve0.Parameter("Distance") = 0
    oPointAlongCurve0.Parameter("TrackFlag") = 2
    oPointAlongCurve0.Evaluate
    
    Dim oLineAtAngle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle6 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "005-LineAtAngle")
    oLineAtAngle6.Inputs("Line").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oLineAtAngle6.Inputs("Point").Add oPointAlongCurve0
    oLineAtAngle6.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oLineAtAngle6.Parameter("Angle") = -1.570796
    oLineAtAngle6.Parameter("TrackFlag") = 1
    oLineAtAngle6.Evaluate

    Dim oLineAtAngle6b As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle6b = oGCFactory.CreateEntity("LineAtAngle", pPOM, "006-LineAtAngle")
    oLineAtAngle6b.Inputs("Line").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oLineAtAngle6b.Inputs("Point").Add oPointAlongCurve0
    oLineAtAngle6b.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oLineAtAngle6b.Parameter("Angle") = 1.570796
    oLineAtAngle6b.Parameter("TrackFlag") = 1
    oLineAtAngle6b.Evaluate

    Dim oCSByPlane7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane7 = oGCFactory.CreateEntity("CSByPlane", pPOM, "007-CSByPlane")
    oCSByPlane7.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oCSByPlane7.Inputs("Point").Add oPointAtCurveStart2
    oCSByPlane7.Evaluate

    Dim oParamDistFunctionWith2Args9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args9 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "008-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args9.Inputs("Arg1").Add oParamDistConstant5
    oParamDistFunctionWith2Args9.Inputs("Arg2").Add oParamDouble4
    oParamDistFunctionWith2Args9.Parameter("Value") = 0
    oParamDistFunctionWith2Args9.Parameter("Operation") = GCDivide
    oParamDistFunctionWith2Args9.Evaluate

    Dim oLineParallel10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel10 = oGCFactory.CreateEntity("LineParallel", pPOM, "009-LineParallel")
    oLineParallel10.Inputs("Line").Add pGeometricConstruction.Inputs("CurveLeft").Item(1)
    oLineParallel10.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oLineParallel10.Inputs("TrackPoint").Add oPointAlongCurve0
    oLineParallel10.Parameter("Distance") = oParamDistConstant5
    oLineParallel10.Parameter("TrackFlag") = 2
    oLineParallel10.Evaluate

    Dim oLineParallel11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel11 = oGCFactory.CreateEntity("LineParallel", pPOM, "010-LineParallel")
    oLineParallel11.Inputs("Line").Add pGeometricConstruction.Inputs("CurveRight").Item(1)
    oLineParallel11.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oLineParallel11.Inputs("TrackPoint").Add oPointAlongCurve0
    oLineParallel11.Parameter("Distance") = oParamDistConstant5
    oLineParallel11.Parameter("TrackFlag") = 2
    oLineParallel11.Evaluate

    Dim oLineFromCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS14 = oGCFactory.CreateEntity("LineFromCS", pPOM, "011-LineFromCS")
    oLineFromCS14.Inputs("CoordinateSystem").Add oCSByPlane7
    oLineFromCS14.Parameter("LookingAxis") = 3
    oLineFromCS14.Parameter("Length") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight"))
    oLineFromCS14.Evaluate
    
    ' replace PointAtCurveEnd by PointAtCurveStart
    Dim oPointAtCurveStart15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart15 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "012-PointAtCurveStart")
    oPointAtCurveStart15.Inputs("Curve").Add oLineAtAngle6
    oPointAtCurveStart15.Evaluate

    ' replace PointAtCurveEnd by PointAtCurveStart
    Dim oPointAtCurveStart15b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart15b = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "013-PointAtCurveStart")
    oPointAtCurveStart15b.Inputs("Curve").Add oLineAtAngle6b
    oPointAtCurveStart15b.Evaluate

    Dim oLineByPoints15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints15 = oGCFactory.CreateEntity("LineByPoints", pPOM, "014-LineByPoints")
    oLineByPoints15.Inputs("StartPoint").Add oPointAtCurveStart15
    oLineByPoints15.Inputs("EndPoint").Add oPointAtCurveStart15b
    oLineByPoints15.Evaluate

    Dim oLineParallel15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel15 = oGCFactory.CreateEntity("LineParallel", pPOM, "015-LineParallel")
    oLineParallel15.Inputs("Line").Add oLineByPoints15
    oLineParallel15.Inputs("Plane").Add pGeometricConstruction.Inputs("Support").Item(1)
    oLineParallel15.Inputs("TrackPoint").Add oPointAtCurveStart2
    oLineParallel15.Parameter("Distance") = oParamDistFunctionWith2Args9
    oLineParallel15.Parameter("TrackFlag") = 1
    oLineParallel15.Evaluate

    Dim oPointByCurves12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves12 = oGCFactory.CreateEntity("PointByCurves", pPOM, "016-PointByCurves")
    oPointByCurves12.Inputs("Curve1").Add oLineParallel15
    oPointByCurves12.Inputs("Curve2").Add oLineParallel11
    oPointByCurves12.Parameter("TrackFlag") = 1
    oPointByCurves12.Evaluate

    Dim oPointByCurves13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves13 = oGCFactory.CreateEntity("PointByCurves", pPOM, "017-PointByCurves")
    oPointByCurves13.Inputs("Curve1").Add oLineParallel10
    oPointByCurves13.Inputs("Curve2").Add oLineParallel15
    oPointByCurves13.Parameter("TrackFlag") = 1
    oPointByCurves13.Evaluate
    
    Dim oPointByCurves16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves16 = oGCFactory.CreateEntity("PointByCurves", pPOM, "018-PointByCurves")
    oPointByCurves16.Inputs("Curve1").Add pGeometricConstruction.Inputs("CurveLeft").Item(1)
    oPointByCurves16.Inputs("Curve2").Add oLineByPoints15
    oPointByCurves16.Parameter("TrackFlag") = 1
    oPointByCurves16.Evaluate

    Dim oPointByCurves17 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves17 = oGCFactory.CreateEntity("PointByCurves", pPOM, "019-PointByCurves")
    oPointByCurves17.Inputs("Curve1").Add pGeometricConstruction.Inputs("CurveRight").Item(1)
    oPointByCurves17.Inputs("Curve2").Add oLineByPoints15
    oPointByCurves17.Parameter("TrackFlag") = 1
    oPointByCurves17.Evaluate
    
'''    Dim oLineByPoints18 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oLineByPoints18 = oGCFactory.CreateEntity("LineByPoints", pPOM, "000-")
'''    oLineByPoints18.Inputs("StartPoint").Add oPointByCurves17
'''    oLineByPoints18.Inputs("EndPoint").Add oPointByCurves16
'''    oLineByPoints18.Evaluate
'''
    Dim oLineByPoints19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints19 = oGCFactory.CreateEntity("LineByPoints", pPOM, "020-LineByPoints")
    oLineByPoints19.Inputs("StartPoint").Add oPointByCurves16
    oLineByPoints19.Inputs("EndPoint").Add oPointByCurves13
    oLineByPoints19.Evaluate

    Dim oLineByPoints20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints20 = oGCFactory.CreateEntity("LineByPoints", pPOM, "021-LineByPoints")
    oLineByPoints20.Inputs("StartPoint").Add oPointByCurves17
    oLineByPoints20.Inputs("EndPoint").Add oPointByCurves12
    oLineByPoints20.Evaluate

    Dim oPointAlongCurve21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve21 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "022-PointAlongCurve")
    oPointAlongCurve21.Inputs("Curve").Add oLineByPoints19
    oPointAlongCurve21.Inputs("Point").Add oPointByCurves13
    oPointAlongCurve21.Inputs("TrackPoint").Add oPointByCurves16
    oPointAlongCurve21.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oPointAlongCurve21.Parameter("TrackFlag") = 2
    oPointAlongCurve21.Evaluate

    Dim oPointAlongCurve21b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve21b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "023-PointAlongCurve")
    oPointAlongCurve21b.Inputs("Curve").Add oLineByPoints19
    oPointAlongCurve21b.Inputs("Point").Add oPointByCurves16
    oPointAlongCurve21b.Inputs("TrackPoint").Add oPointByCurves13
    oPointAlongCurve21b.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oPointAlongCurve21b.Parameter("TrackFlag") = 2
    oPointAlongCurve21b.Evaluate

    Dim oPointAlongCurve22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve22 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "024-PointAlongCurve")
    oPointAlongCurve22.Inputs("Curve").Add oLineByPoints20
    oPointAlongCurve22.Inputs("Point").Add oPointByCurves12
    oPointAlongCurve22.Inputs("TrackPoint").Add oPointByCurves17
    oPointAlongCurve22.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oPointAlongCurve22.Parameter("TrackFlag") = 2
    oPointAlongCurve22.Evaluate

    Dim oPointAlongCurve22b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve22b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "025-PointAlongCurve")
    oPointAlongCurve22b.Inputs("Curve").Add oLineByPoints20
    oPointAlongCurve22b.Inputs("Point").Add oPointByCurves17
    oPointAlongCurve22b.Inputs("TrackPoint").Add oPointByCurves12
    oPointAlongCurve22b.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oPointAlongCurve22b.Parameter("TrackFlag") = 2
    oPointAlongCurve22b.Evaluate

    Dim oLineByPoints23 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints23 = oGCFactory.CreateEntity("LineByPoints", pPOM, "026-LineByPoints")
    oLineByPoints23.Inputs("StartPoint").Add oPointAlongCurve22b
    oLineByPoints23.Inputs("EndPoint").Add oPointAlongCurve22
    oLineByPoints23.Evaluate

    Dim oLineByPoints24 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints24 = oGCFactory.CreateEntity("LineByPoints", pPOM, "027-LineByPoints")
    oLineByPoints24.Inputs("StartPoint").Add oPointAlongCurve21b
    oLineByPoints24.Inputs("EndPoint").Add oPointAlongCurve21
    oLineByPoints24.Evaluate

'''    Dim oCpxStringByCurves25 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oCpxStringByCurves25 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "000-")
'''    oCpxStringByCurves25.Inputs("Curves").Add oLineByPoints23
'''    oCpxStringByCurves25.Inputs("Curves").Add oLineByPoints18
'''    oCpxStringByCurves25.Inputs("Curves").Add oLineByPoints24
'''    oCpxStringByCurves25.Evaluate

    Dim oSurfByLinearExtrusion26 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion26 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "028-SurfByLinearExtrusion")
    oSurfByLinearExtrusion26.Inputs("PlanarCrossSection").Add oLineByPoints23
    oSurfByLinearExtrusion26.Inputs("ExtrusionLine").Add oLineFromCS14
    oSurfByLinearExtrusion26.Evaluate

    Dim oSurfByLinearExtrusion27 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion27 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "029-SurfByLinearExtrusion")
    oSurfByLinearExtrusion27.Inputs("PlanarCrossSection").Add oLineByPoints24
    oSurfByLinearExtrusion27.Inputs("ExtrusionLine").Add oLineFromCS14
    oSurfByLinearExtrusion27.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("SurfaceLeft", 1) = oSurfByLinearExtrusion27.Output
    oGCMacro.Output("SurfaceRight", 1) = oSurfByLinearExtrusion26.Output
End Sub
