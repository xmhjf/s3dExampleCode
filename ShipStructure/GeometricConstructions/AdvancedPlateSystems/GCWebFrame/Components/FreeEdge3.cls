VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdge3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService

'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "FreeEdge3"
'End Property
'
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "FreeEdge3"
'End Property

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Curve", "Select a curve", "IJCurve or IJWireBody", 1, 1)
    'Call pGeometricConstructionDefinition.AddInput("Surface", "Select a surface", "IJSurface or IJSurfaceBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point1", "Select a lower point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point2", "Select an upper point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem1", "Select a lower coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem2", "Select an upper coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a middle coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Angle1", "Lower angle", 8, 2, 78, 0, 0, 30#)
    Call pGeometricConstructionDefinition.AddParameter("Angle2", "Upper angle", 8, 2, 78, 0, 0, 30#)
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oLineByPoints1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints1 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0001-LineByPoints")
    oLineByPoints1.Inputs("StartPoint").Add pGeometricConstruction.Inputs("Point2").Item(1), "1"
    oLineByPoints1.Inputs("EndPoint").Add pGeometricConstruction.Inputs("Point1").Item(1), "1"
    oLineByPoints1.Evaluate

    Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0002-LineFromCS")
    oLineFromCS2.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1), "1"
    oLineFromCS2.Parameter("LookingAxis") = 3
    oLineFromCS2.Parameter("Length") = 2#
    oLineFromCS2.Parameter("CSOrientation") = 2
    oLineFromCS2.Parameter("LineJustification") = 1
    oLineFromCS2.Evaluate

    Dim oLineFromCS3 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS3 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0003-LineFromCS")
    oLineFromCS3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1), "2"
    oLineFromCS3.Parameter("LookingAxis") = 3
    oLineFromCS3.Parameter("Length") = 2#
    oLineFromCS3.Parameter("CSOrientation") = 2
    oLineFromCS3.Parameter("LineJustification") = 1
    oLineFromCS3.Evaluate

    Dim oLineByPointAngleLength4 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPointAngleLength4 = oGCFactory.CreateEntity("LineByPointAngleLength", pPOM, "0004-LineByPointAngleLength")
    oLineByPointAngleLength4.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1), "1"
    oLineByPointAngleLength4.Inputs("Point").Add pGeometricConstruction.Inputs("Point1").Item(1), "1"
    oLineByPointAngleLength4.Parameter("AxesRoles") = 2
    oLineByPointAngleLength4.Parameter("Angle") = pGeometricConstruction.Parameter("Angle1")
    oLineByPointAngleLength4.Parameter("Length") = 10#
    oLineByPointAngleLength4.Parameter("TrackFlag") = 1
    oLineByPointAngleLength4.Evaluate

    Dim oLineByPointAngleLength5 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPointAngleLength5 = oGCFactory.CreateEntity("LineByPointAngleLength", pPOM, "0005-LineByPointAngleLength")
    oLineByPointAngleLength5.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1), "2"
    oLineByPointAngleLength5.Inputs("Point").Add pGeometricConstruction.Inputs("Point2").Item(1), "1"
    oLineByPointAngleLength5.Parameter("AxesRoles") = 2
    oLineByPointAngleLength5.Parameter("Angle") = pGeometricConstruction.Parameter("Angle2")
    oLineByPointAngleLength5.Parameter("Length") = 10#
    oLineByPointAngleLength5.Parameter("TrackFlag") = 1
    oLineByPointAngleLength5.Evaluate

    Dim oLineFromCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS6 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0006-LineFromCS")
    oLineFromCS6.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oLineFromCS6.Parameter("LookingAxis") = 1
    oLineFromCS6.Parameter("Length") = -1
    oLineFromCS6.Parameter("CSOrientation") = 2
    oLineFromCS6.Parameter("LineJustification") = 1
    oLineFromCS6.Evaluate

    Dim oPointFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS7 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0007-PointFromCS")
    oPointFromCS7.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPointFromCS7.Parameter("X") = 0#
    oPointFromCS7.Parameter("Y") = 100.0#
    oPointFromCS7.Parameter("Z") = 0#
    oPointFromCS7.Evaluate

    Dim oCurveByCurves8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves8 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "0008a-CurveByCurves")
    oCurveByCurves8.Inputs("Curves").Add oLineFromCS2, "1"
    oCurveByCurves8.Inputs("Curves").Add oLineByPoints1, "2"
    oCurveByCurves8.Inputs("Curves").Add oLineFromCS3, "3"
    oCurveByCurves8.Parameter("ConstructionSurface") = 0
    oCurveByCurves8.Evaluate
    
    Dim oCurveByProjection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection8 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0008b-CurveByProjection")
    oCurveByProjection8.Inputs("Curve").Add oCurveByCurves8, "1"
    oCurveByProjection8.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByProjection8.Inputs("Line").Add oLineFromCS6, "1"
    oCurveByProjection8.Evaluate

    Dim oCurveAtOffset9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveAtOffset9 = oGCFactory.CreateEntity("CurveAtOffset", pPOM, "0009-CurveAtOffset")
    oCurveAtOffset9.Inputs("Curve").Add pGeometricConstruction.Inputs("Curve").Item(1), "1"
    oCurveAtOffset9.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveAtOffset9.Inputs("TrackPoint").Add oPointFromCS7, "1"
    oCurveAtOffset9.Parameter("Offset") = pGeometricConstruction.Parameter("Offset")
    oCurveAtOffset9.Parameter("Orientation") = 1
    oCurveAtOffset9.Parameter("TrackFlag") = 1
    oCurveAtOffset9.Evaluate
    
    Dim oCurveByProjection9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection9 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0021b-CurveByProjection")
    oCurveByProjection9.Inputs("Curve").Add oCurveAtOffset9, "1"
    oCurveByProjection9.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByProjection9.Inputs("Line").Add oLineFromCS6, "1"
    oCurveByProjection9.Evaluate


'    Dim oSurfAtOffset9 As SP3DGeometricConstruction.GeometricConstruction
'    Set oSurfAtOffset9 = oGCFactory.CreateEntity("SurfAtOffset", pPOM, "0009a-SurfAtOffset")
'    oSurfAtOffset9.Inputs("Surface").Add pGeometricConstruction.Inputs("Surface").Item(1), "1"
'    oSurfAtOffset9.Inputs("TrackPoint").Add oPointFromCS7, "1"
'    oSurfAtOffset9.Parameter("Offset") = pGeometricConstruction.Parameter("Offset")
'    oSurfAtOffset9.Parameter("Orientation") = 1
'    oSurfAtOffset9.Parameter("TrackFlag") = 1
'    oSurfAtOffset9.Evaluate
'
'    Dim oCurveByIntersection9 As SP3DGeometricConstruction.GeometricConstruction
'    Set oCurveByIntersection9 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0009b-CurveByIntersection")
'    oCurveByIntersection9.Inputs("Surface1").Add oSurfAtOffset9, "1"
'    oCurveByIntersection9.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
'    oCurveByIntersection9.Parameter("TrackFlag") = 1
'    oCurveByIntersection9.Parameter("ConstructionSurface") = 0
'    oCurveByIntersection9.Evaluate

    Dim oSurfByLinearExtrusion10 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion10 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0010-SurfByLinearExtrusion")
    oSurfByLinearExtrusion10.Inputs("PlanarCrossSection").Add oCurveByProjection8, "1"
    oSurfByLinearExtrusion10.Inputs("ExtrusionLine").Add oLineFromCS6, "1"
    oSurfByLinearExtrusion10.Evaluate

    Dim oPointByCurves11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves11 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0011-PointByCurves")
    oPointByCurves11.Inputs("Curve1").Add oCurveByProjection9, "1"
    oPointByCurves11.Inputs("Curve2").Add oLineByPointAngleLength4, "1"
    oPointByCurves11.Parameter("TrackFlag") = 1
    oPointByCurves11.Evaluate

    Dim oPointByCurves12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves12 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0012-PointByCurves")
    oPointByCurves12.Inputs("Curve1").Add oCurveByProjection9, "1"
    oPointByCurves12.Inputs("Curve2").Add oLineByPointAngleLength5, "1"
    oPointByCurves12.Parameter("TrackFlag") = 1
    oPointByCurves12.Evaluate

    Dim oCurveExtractPartial13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveExtractPartial13 = oGCFactory.CreateEntity("CurveExtractPartial", pPOM, "0013-CurveExtractPartial")
    oCurveExtractPartial13.Inputs("Curve").Add oCurveByProjection9, "1"
    oCurveExtractPartial13.Inputs("Splitters").Add oPointByCurves12, "1"
    oCurveExtractPartial13.Inputs("Splitters").Add oPointByCurves11, "2"
    oCurveExtractPartial13.Parameter("TrackFlag") = 1
    oCurveExtractPartial13.Parameter("ConstructionSurface") = 0
    oCurveExtractPartial13.Evaluate

    Dim oLineByPoints14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints14 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0014-LineByPoints")
    oLineByPoints14.Inputs("StartPoint").Add pGeometricConstruction.Inputs("Point1").Item(1), "1"
    oLineByPoints14.Inputs("EndPoint").Add oPointByCurves11, "1"
    oLineByPoints14.Evaluate

    Dim oLineByPoints15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints15 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0015-LineByPoints")
    oLineByPoints15.Inputs("StartPoint").Add pGeometricConstruction.Inputs("Point2").Item(1), "1"
    oLineByPoints15.Inputs("EndPoint").Add oPointByCurves12, "1"
    oLineByPoints15.Evaluate

    Dim oArcByCurvesAndPoint16 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCurvesAndPoint16 = oGCFactory.CreateEntity("ArcByCurvesAndPoint", pPOM, "0016-ArcByCurvesAndPoint")
    oArcByCurvesAndPoint16.Inputs("Curve1").Add oLineByPoints14, "1"
    oArcByCurvesAndPoint16.Inputs("Curve2").Add oCurveExtractPartial13, "1"
    oArcByCurvesAndPoint16.Inputs("Point").Add pGeometricConstruction.Inputs("Point1").Item(1), "1"
    oArcByCurvesAndPoint16.Evaluate

    Dim oArcByCurvesAndPoint17 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCurvesAndPoint17 = oGCFactory.CreateEntity("ArcByCurvesAndPoint", pPOM, "0017-ArcByCurvesAndPoint")
    oArcByCurvesAndPoint17.Inputs("Curve1").Add oLineByPoints15, "1"
    oArcByCurvesAndPoint17.Inputs("Curve2").Add oCurveExtractPartial13, "1"
    oArcByCurvesAndPoint17.Inputs("Point").Add pGeometricConstruction.Inputs("Point2").Item(1), "1"
    oArcByCurvesAndPoint17.Evaluate

    Dim oPointAtCurveEnd18 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd18 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0018-PointAtCurveEnd")
    oPointAtCurveEnd18.Inputs("Curve").Add oArcByCurvesAndPoint16, "1"
    oPointAtCurveEnd18.Evaluate

    Dim oPointAtCurveEnd19 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd19 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0019-PointAtCurveEnd")
    oPointAtCurveEnd19.Inputs("Curve").Add oArcByCurvesAndPoint17, "2"
    oPointAtCurveEnd19.Evaluate

    Dim oCurveExtractPartial20 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveExtractPartial20 = oGCFactory.CreateEntity("CurveExtractPartial", pPOM, "0020-CurveExtractPartial")
    oCurveExtractPartial20.Inputs("Curve").Add oCurveByProjection9, "1"
    oCurveExtractPartial20.Inputs("Splitters").Add oPointAtCurveEnd18, "1"
    oCurveExtractPartial20.Inputs("Splitters").Add oPointAtCurveEnd19, "2"
    oCurveExtractPartial20.Parameter("TrackFlag") = 1
    oCurveExtractPartial20.Parameter("ConstructionSurface") = 0
    oCurveExtractPartial20.Evaluate

    Dim oCurveByCurves21 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves21 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "0021a-CurveByCurves")
    oCurveByCurves21.Inputs("Curves").Add oArcByCurvesAndPoint16, "1"
    oCurveByCurves21.Inputs("Curves").Add oCurveExtractPartial20, "2"
    oCurveByCurves21.Inputs("Curves").Add oArcByCurvesAndPoint17, "3"
    oCurveByCurves21.Parameter("ConstructionSurface") = 0
    oCurveByCurves21.Evaluate
    
    Dim oCurveByProjection21 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection21 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0021b-CurveByProjection")
    oCurveByProjection21.Inputs("Curve").Add oCurveByCurves21, "1"
    oCurveByProjection21.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByProjection21.Inputs("Line").Add oLineFromCS6, "1"
    oCurveByProjection21.Evaluate

    Dim oSurfByLinearExtrusion22 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion22 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0022-SurfByLinearExtrusion")
    oSurfByLinearExtrusion22.Inputs("PlanarCrossSection").Add oCurveByProjection21, "1"
    oSurfByLinearExtrusion22.Inputs("ExtrusionLine").Add oLineFromCS6, "1"
    oSurfByLinearExtrusion22.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", "1") = oSurfByLinearExtrusion10.Output
    oGCMacro.Output("Boundary", "2") = oSurfByLinearExtrusion22.Output
End Sub
