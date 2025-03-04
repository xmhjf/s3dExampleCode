VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StraightEdgeAround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const E_FAIL = &H80004005
Const CURVES_NOT_INTERSECTING = 4
Const RADIUS_TOO_BIG = 6
Const POINTS_TOO_CLOSE = 2
Const TOLERANCE = 0.001

Private Enum Errors
    WRONG_OFFSET = 1
    OFFSET_TOO_SMALL = 2
    FILLET_TOO_BIG = 3
    ARC_IS_COMPLEMENT = 4
    CUTBACK_DISTANCE_TOO_BIG = 5
End Enum

Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    ' define inputs
    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select axis line #1", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select axis line #2", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("EdgeLine1", "Select edge line #1", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("EdgeLine2", "Select edge line #2", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Circle", "Select circle", "IJCircle OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJCircle]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("NormalLine", "Select line normal to circle", "IJLine", 1, 1)

    ' define parameters
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    ' define  errors
    Call pGeometricConstructionDefinition.AddErrorValue(WRONG_OFFSET, "WrongOffset", "Offseted circle (Offset) and offseted edge (WeldToe) do not intersect: increase/decrease Offset or decrease WeldToe")
    Call pGeometricConstructionDefinition.AddErrorValue(OFFSET_TOO_SMALL, "OffsetTooSmall", "Offseted circle (Offset) and offseted edge (WeldToe) are too close: increase Offset or decrease WeldToe")
    Call pGeometricConstructionDefinition.AddErrorValue(FILLET_TOO_BIG, "FilletTooBig", "Fillet cannot be created: decrease the Fillet, the Offset or the WeldToe")
    Call pGeometricConstructionDefinition.AddErrorValue(ARC_IS_COMPLEMENT, "ArcIsComplement", "The arc along the offseted circle (Offset) is degenerated: decrease the Fillet or increase the Offset")
    Call pGeometricConstructionDefinition.AddErrorValue(CUTBACK_DISTANCE_TOO_BIG, "CutBackDistanceTooBig", "The anchor point for trimming plane is outside the member axis: decrease the CutbackDistance")
    
    ' define outputs
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    'Call pGeometricConstructionDefinition.AddOutput(GCGTypeComplexString3d, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
    
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oParamDistConstant1.Evaluate
    oParamDistConstant1.Public = True
    
    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    oParamDistConstant2.Evaluate
    oParamDistConstant2.Public = True
    
    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Offset"))
    oParamDistConstant3.Evaluate
    oParamDistConstant3.Public = True
    
    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oParamDistConstant4.Evaluate
    oParamDistConstant4.Public = True
    
    ' compute point at node
    Dim oPointAtCurveStart6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart6 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "005-PointAtCurveStart")
    oPointAtCurveStart6.Inputs("Curve").Add pGeometricConstruction.Input("AxisLine1")
    oPointAtCurveStart6.Evaluate
    oPointAtCurveStart6.Public = True
    
    ' project it on the normal line
    Dim oPointAlongCurve6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "006-PointAlongCurve")
    oPointAlongCurve6.Inputs("Curve").Add pGeometricConstruction.Input("NormalLine")
    oPointAlongCurve6.Inputs("Point").Add oPointAtCurveStart6
    oPointAlongCurve6.Parameter("Distance") = 0
    oPointAlongCurve6.Parameter("TrackFlag") = 1
    oPointAlongCurve6.Evaluate
    oPointAlongCurve6.Public = True
    
    ' compute point at other extremity of axes lines
    Dim oPointAtCurveEnd7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd7 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "007-PointAtCurveEnd")
    oPointAtCurveEnd7.Inputs("Curve").Add pGeometricConstruction.Input("AxisLine1")
'''    oPointAtCurveEnd7.Inputs("TrackPoint").Add oPointAtCurveStart6
'''    oPointAtCurveEnd7.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAtCurveEnd7, pGeometricConstruction)
    oPointAtCurveEnd7.Public = True

    Dim oPointAtCurveEnd8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd8 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "008-PointAtCurveEnd")
    oPointAtCurveEnd8.Inputs("Curve").Add pGeometricConstruction.Input("AxisLine2")
'''    oPointAtCurveEnd8.Inputs("TrackPoint").Add oPointAtCurveStart6
'''    oPointAtCurveEnd8.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAtCurveEnd8, pGeometricConstruction)
    oPointAtCurveEnd8.Public = True

    ' compute plane big enough to project the circle
    Dim oPlaneByPointNormal9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal9 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "009-PlaneByPointNormal")
    oPlaneByPointNormal9.Inputs("Point").Add oPointAlongCurve6 'oPointAtCurveStart6
    oPlaneByPointNormal9.Inputs("Line").Add pGeometricConstruction.Input("NormalLine")
'''    Dim dRadius As Double
'''    If True Then
'''        Dim pCircle As IJCircle
'''        Set pCircle = pGeometricConstruction.Input("Circle")
'''
'''        Let dRadius = pCircle.Radius
'''    End If
'''    oPlaneByPointNormal9.Parameter("Range") = 2 * dRadius
    Dim dLength As Double
    If True Then
        Dim pLine1 As IJLine
        Set pLine1 = pGeometricConstruction.Input("AxisLine1")
        
        Dim pLine2 As IJLine
        Set pLine2 = pGeometricConstruction.Input("AxisLine2")
        
        Let dLength = pLine1.Length()
        If pLine2.Length() > dLength Then Let dLength = pLine2.Length()
    End If
    oPlaneByPointNormal9.Parameter("Range") = 2 * dLength
    oPlaneByPointNormal9.Evaluate
    oPlaneByPointNormal9.Public = True
    
    ' compute line for projection and extrusion purpose
    If True Then
        Dim oPointAlongCurve10 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve10 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "010-PointAlongCurve")
        oPointAlongCurve10.Inputs("Curve").Add pGeometricConstruction.Input("NormalLine")
        oPointAlongCurve10.Inputs("Point").Add oPointAlongCurve6 'oPointAtCurveStart6
        oPointAlongCurve10.Parameter("Distance") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
        oPointAlongCurve10.Parameter("TrackFlag") = 1
        oPointAlongCurve10.Evaluate
        oPointAlongCurve10.Public = True

        Dim oPointAlongCurve11 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "011-PointAlongCurve")
        oPointAlongCurve11.Inputs("Curve").Add pGeometricConstruction.Input("NormalLine")
        oPointAlongCurve11.Inputs("Point").Add oPointAlongCurve6 'oPointAtCurveStart6
        oPointAlongCurve11.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
        oPointAlongCurve11.Parameter("TrackFlag") = 1
        oPointAlongCurve11.Evaluate
        oPointAlongCurve11.Public = True
    End If
    
    Dim oLineByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints12 = oGCFactory.CreateEntity("LineByPoints", pPOM, "012-LineByPoints")
     oLineByPoints12.Inputs("StartPoint").Add oPointAlongCurve10
    oLineByPoints12.Inputs("EndPoint").Add oPointAlongCurve11
    oLineByPoints12.Evaluate
    oLineByPoints12.Public = True
    
    ' project EdgeLine1
    Dim oCurveByProjection13a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection13a = oGCFactory.CreateEntity("CurveByProjection", pPOM, "013-CurveByProjection")
    oCurveByProjection13a.Inputs("Curve").Add pGeometricConstruction.Input("EdgeLine1")
    oCurveByProjection13a.Inputs("Surface").Add oPlaneByPointNormal9
    oCurveByProjection13a.Inputs("Line").Add oLineByPoints12
    oCurveByProjection13a.Public = True
    oCurveByProjection13a.Evaluate
    
    ' project EdgeLine2
    Dim oCurveByProjection13b As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection13b = oGCFactory.CreateEntity("CurveByProjection", pPOM, "013-CurveByProjection")
    oCurveByProjection13b.Inputs("Curve").Add pGeometricConstruction.Input("EdgeLine2")
    oCurveByProjection13b.Inputs("Surface").Add oPlaneByPointNormal9
    oCurveByProjection13b.Inputs("Line").Add oLineByPoints12
    oCurveByProjection13b.Public = True
    oCurveByProjection13b.Evaluate
    
    Dim oLineParallel14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel14 = oGCFactory.CreateEntity("LineParallel", pPOM, "014-LineParallel")
    oLineParallel14.Inputs("Line").Add oCurveByProjection13a
    oLineParallel14.Inputs("Plane").Add oPlaneByPointNormal9
    oLineParallel14.Inputs("TrackPoint").Add oPointAtCurveEnd7 'oPointAtCurveStart6
    oLineParallel14.Parameter("Distance") = oParamDistConstant1
    oLineParallel14.Parameter("TrackFlag") = 2
    Call GCEvaluate(oLineParallel14, pGeometricConstruction)
    oLineParallel14.Public = True

    Dim oLineParallel15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel15 = oGCFactory.CreateEntity("LineParallel", pPOM, "015-LineParallel")
    oLineParallel15.Inputs("Line").Add oCurveByProjection13b
    oLineParallel15.Inputs("Plane").Add oPlaneByPointNormal9
    oLineParallel15.Inputs("TrackPoint").Add oPointAtCurveEnd8 'oPointAtCurveStart6
    oLineParallel15.Parameter("Distance") = oParamDistConstant1
    oLineParallel15.Parameter("TrackFlag") = 2
    Call GCEvaluate(oLineParallel15, pGeometricConstruction)
    oLineParallel15.Public = True

    Dim oPointAtCurveEnd16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd16 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "016-PointAtCurveEnd")
    oPointAtCurveEnd16.Inputs("Curve").Add pGeometricConstruction.Input("EdgeLine1")
    Call GCEvaluate(oPointAtCurveEnd16, pGeometricConstruction)
    oPointAtCurveEnd16.Public = True

    Dim oPointAtCurveExtremity16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity16 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "016-PointAtCurveExtremity")
    oPointAtCurveExtremity16.Inputs("Curve").Add oLineParallel14
    oPointAtCurveExtremity16.Inputs("TrackPoint").Add oPointAtCurveEnd16 'oPointAtCurveStart6
    oPointAtCurveExtremity16.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointAtCurveExtremity16, pGeometricConstruction)
    oPointAtCurveExtremity16.Public = True

    Dim oPointAtCurveEnd17 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd17 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "017-PointAtCurveEnd")
    oPointAtCurveEnd17.Inputs("Curve").Add pGeometricConstruction.Input("EdgeLine2")
    Call GCEvaluate(oPointAtCurveEnd17, pGeometricConstruction)
    oPointAtCurveEnd17.Public = True

    Dim oPointAtCurveExtremity17 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity17 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "017-PointAtCurveExtremity")
    oPointAtCurveExtremity17.Inputs("Curve").Add oLineParallel15
    oPointAtCurveExtremity17.Inputs("TrackPoint").Add oPointAtCurveEnd17 'oPointAtCurveStart6
    oPointAtCurveExtremity17.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointAtCurveExtremity17, pGeometricConstruction)
    oPointAtCurveExtremity17.Public = True
    
    Dim oCurveByProjection18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection18 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "018-CurveByProjection")
    oCurveByProjection18.Inputs("Curve").Add pGeometricConstruction.Input("Circle")
    oCurveByProjection18.Inputs("Surface").Add oPlaneByPointNormal9
    oCurveByProjection18.Inputs("Line").Add oLineByPoints12
    oCurveByProjection18.Public = True
    oCurveByProjection18.Evaluate
    
    Dim oCircleConcentric19 As SP3DGeometricConstruction.GeometricConstruction
    Set oCircleConcentric19 = oGCFactory.CreateEntity("CircleConcentric", pPOM, "019-CircleConcentric")
    oCircleConcentric19.Inputs("Circle").Add oCurveByProjection18
    'oCircleConcentric19.Inputs("TrackPoint").Add oPointAtCurveStart6
    oCircleConcentric19.Parameter("Offset") = oParamDistConstant3
    'oCircleConcentric19.Parameter("TrackFlag") = 2
    Call GCEvaluate(oCircleConcentric19, pGeometricConstruction)
    oCircleConcentric19.Public = True

    Dim pPointOfCenter As IJPoint
    If True Then
        Dim pPositionOfCenter As IJDPosition: Set pPositionOfCenter = Position_FromCircleCenter(oCircleConcentric19)
        Dim pGeometryFactory As IJGeometryFactory: Set pGeometryFactory = New GeometryFactory
        Set pPointOfCenter = pGeometryFactory.Points3d.CreateByPoint(pPOM, pPositionOfCenter.x, pPositionOfCenter.y, pPositionOfCenter.z)
    End If
   
    Dim oPointByCurves20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves20 = oGCFactory.CreateEntity("PointByCurves", pPOM, "020-PointByCurves")
    oPointByCurves20.Inputs("Curve1").Add oLineParallel14
    oPointByCurves20.Inputs("Curve2").Add oCircleConcentric19
    oPointByCurves20.Inputs("TrackPoint").Add oPointAtCurveEnd7
    oPointByCurves20.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointByCurves20, pGeometricConstruction, CURVES_NOT_INTERSECTING, WRONG_OFFSET)
    oPointByCurves20.Public = True

    Dim oPointByCurves21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves21 = oGCFactory.CreateEntity("PointByCurves", pPOM, "021-PointByCurves")
    oPointByCurves21.Inputs("Curve1").Add oLineParallel15
    oPointByCurves21.Inputs("Curve2").Add oCircleConcentric19
    oPointByCurves21.Inputs("TrackPoint").Add oPointAtCurveEnd8
    oPointByCurves21.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointByCurves21, pGeometricConstruction, CURVES_NOT_INTERSECTING, WRONG_OFFSET)
    oPointByCurves21.Public = True

    Dim oLineByPoints22 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints22 = oGCFactory.CreateEntity("LineByPoints", pPOM, "022-LineByPoints")
    oLineByPoints22.Inputs("StartPoint").Add oPointByCurves20
    oLineByPoints22.Inputs("EndPoint").Add oPointAtCurveExtremity16
    Call GCEvaluate(oLineByPoints22, pGeometricConstruction)
    oLineByPoints22.Public = True

    Dim oLineByPoints23 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints23 = oGCFactory.CreateEntity("LineByPoints", pPOM, "023-LineByPoints")
    oLineByPoints23.Inputs("StartPoint").Add oPointByCurves21
    oLineByPoints23.Inputs("EndPoint").Add oPointAtCurveExtremity17
    Call GCEvaluate(oLineByPoints23, pGeometricConstruction)
    oLineByPoints23.Public = True

    Dim oArcByCenter24 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCenter24 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "024-ArcByCenter")
    oArcByCenter24.Inputs("Center").Add pPointOfCenter 'oPointAtCurveStart6
    oArcByCenter24.Inputs("StartPoint").Add oPointByCurves20
    oArcByCenter24.Inputs("EndPoint").Add oPointByCurves21
    oArcByCenter24.Inputs("NormalLine").Add pGeometricConstruction.Input("NormalLine")
    Call GCEvaluate(oArcByCenter24, pGeometricConstruction, POINTS_TOO_CLOSE, OFFSET_TOO_SMALL)
    oArcByCenter24.Public = True

    Dim oArcByFillet25 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet25 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "025-ArcByFillet")
    oArcByFillet25.Inputs("Curve1").Add oLineByPoints22
    oArcByFillet25.Inputs("Curve2").Add oArcByCenter24
    oArcByFillet25.Parameter("Radius") = oParamDistConstant2
    Call GCEvaluate(oArcByFillet25, pGeometricConstruction, RADIUS_TOO_BIG, FILLET_TOO_BIG)
    oArcByFillet25.Public = True

    Dim oArcByFillet26 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet26 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "026-ArcByFillet")
    oArcByFillet26.Inputs("Curve1").Add oLineByPoints23
    oArcByFillet26.Inputs("Curve2").Add oArcByCenter24
    oArcByFillet26.Parameter("Radius") = oParamDistConstant2
    Call GCEvaluate(oArcByFillet26, pGeometricConstruction, RADIUS_TOO_BIG, FILLET_TOO_BIG)
    oArcByFillet26.Public = True

    Dim oPointAtCurveExtremity27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity27 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "027-PointAtCurveExtremity")
    oPointAtCurveExtremity27.Inputs("Curve").Add oArcByFillet25
    oPointAtCurveExtremity27.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAtCurveExtremity27.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointAtCurveExtremity27, pGeometricConstruction)
    oPointAtCurveExtremity27.Public = True

    Dim oPointAtCurveExtremity28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity28 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "028-PointAtCurveExtremity")
    oPointAtCurveExtremity28.Inputs("Curve").Add oArcByFillet25
    oPointAtCurveExtremity28.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAtCurveExtremity28.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAtCurveExtremity28, pGeometricConstruction)
    oPointAtCurveExtremity28.Public = True

    Dim oPointAtCurveExtremity29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity29 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "029-PointAtCurveExtremity")
    oPointAtCurveExtremity29.Inputs("Curve").Add oArcByFillet26
    oPointAtCurveExtremity29.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAtCurveExtremity29.Parameter("TrackFlag") = 1
    Call GCEvaluate(oPointAtCurveExtremity29, pGeometricConstruction)
    oPointAtCurveExtremity29.Public = True

    Dim oPointAtCurveExtremity30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity30 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "030-PointAtCurveExtremity")
    oPointAtCurveExtremity30.Inputs("Curve").Add oArcByFillet26
    oPointAtCurveExtremity30.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAtCurveExtremity30.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAtCurveExtremity30, pGeometricConstruction)
    oPointAtCurveExtremity30.Public = True

    Dim oLineByPoints31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints31 = oGCFactory.CreateEntity("LineByPoints", pPOM, "031-LineByPoints")
    oLineByPoints31.Inputs("StartPoint").Add oPointAtCurveExtremity28
    oLineByPoints31.Inputs("EndPoint").Add oPointAtCurveExtremity16
    Call GCEvaluate(oLineByPoints31, pGeometricConstruction)
    oLineByPoints31.Public = True
    
    Dim oPointAlongCurve31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve31 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "032-PointAlongCurve")
    oPointAlongCurve31.Inputs("Curve").Add oLineByPoints31
    oPointAlongCurve31.Inputs("Point").Add oPointAtCurveEnd7
    oPointAlongCurve31.Parameter("Distance") = 0
    oPointAlongCurve31.Parameter("TrackFlag") = 1
    oPointAlongCurve31.Evaluate
    oPointAlongCurve31.Public = True

    Dim oLineByPoints31s As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints31s = oGCFactory.CreateEntity("LineByPoints", pPOM, "033-LineByPoints")
    oLineByPoints31s.Inputs("StartPoint").Add oPointAtCurveExtremity28
    oLineByPoints31s.Inputs("EndPoint").Add oPointAlongCurve31
    Call GCEvaluate(oLineByPoints31s, pGeometricConstruction)
    oLineByPoints31s.Public = True

    Dim oPointAlongCurve32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve32 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "034-PointAlongCurve")
    oPointAlongCurve32.Inputs("Curve").Add oLineParallel14
    oPointAlongCurve32.Inputs("Point").Add oPointAtCurveExtremity28
    oPointAlongCurve32.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAlongCurve32.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve32.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAlongCurve32, pGeometricConstruction)
    oPointAlongCurve32.Public = True

    If Not IsPositionInsideCurve(Position_FromPoint(oPointAlongCurve32.Output), oLineParallel14.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)
    
    Dim oLineByPoints33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints33 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints33.Inputs("StartPoint").Add oPointAtCurveExtremity30
    oLineByPoints33.Inputs("EndPoint").Add oPointAtCurveExtremity17
    Call GCEvaluate(oLineByPoints33, pGeometricConstruction)
    oLineByPoints33.Public = True

    Dim oPointAlongCurve33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve33 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "036-PointAlongCurve")
    oPointAlongCurve33.Inputs("Curve").Add oLineByPoints33
    oPointAlongCurve33.Inputs("Point").Add oPointAtCurveEnd8
    oPointAlongCurve33.Parameter("Distance") = 0
    oPointAlongCurve33.Parameter("TrackFlag") = 1
    oPointAlongCurve33.Evaluate
    oPointAlongCurve33.Public = True

    Dim oLineByPoints33s As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints33s = oGCFactory.CreateEntity("LineByPoints", pPOM, "037-LineByPoints")
    oLineByPoints33s.Inputs("StartPoint").Add oPointAtCurveExtremity30
    oLineByPoints33s.Inputs("EndPoint").Add oPointAlongCurve33
    Call GCEvaluate(oLineByPoints33s, pGeometricConstruction)
    oLineByPoints33s.Public = True
    
    Dim oPointAlongCurve34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve34 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "038-PointAlongCurve")
    oPointAlongCurve34.Inputs("Curve").Add oLineParallel15
    oPointAlongCurve34.Inputs("Point").Add oPointAtCurveExtremity30
    oPointAlongCurve34.Inputs("TrackPoint").Add pPointOfCenter 'oPointAtCurveStart6
    oPointAlongCurve34.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve34.Parameter("TrackFlag") = 2
    Call GCEvaluate(oPointAlongCurve34, pGeometricConstruction)
    oPointAlongCurve34.Public = True

    If Not IsPositionInsideCurve(Position_FromPoint(oPointAlongCurve34.Output), oLineParallel15.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)

    Dim oArcByCenter35 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCenter35 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "035-ArcByCenter")
    oArcByCenter35.Inputs("Center").Add pPointOfCenter 'oPointAtCurveStart6
    oArcByCenter35.Inputs("StartPoint").Add oPointAtCurveExtremity27
    oArcByCenter35.Inputs("EndPoint").Add oPointAtCurveExtremity29
    oArcByCenter35.Inputs("NormalLine").Add pGeometricConstruction.Input("NormalLine")
    Call GCEvaluate(oArcByCenter35, pGeometricConstruction, POINTS_TOO_CLOSE, OFFSET_TOO_SMALL)
    oArcByCenter35.Public = True

    ' when the fillets are too big, the small arc cannot be created. The complement arc is created unstead
    Dim pArc1 As IJArc: Set pArc1 = oArcByCenter24
    Dim pArc2 As IJArc: Set pArc2 = oArcByCenter35
    If pArc2.SweepAngle > pArc1.SweepAngle Then Call GCProcessError(pGeometricConstruction, , ARC_IS_COMPLEMENT)

    Dim oCpxStringByCurves36 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves36 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "036-CpxStringByCurves")
    oCpxStringByCurves36.Inputs("Curves").Add oLineByPoints31s
    oCpxStringByCurves36.Inputs("Curves").Add oArcByFillet25
    oCpxStringByCurves36.Inputs("Curves").Add oArcByCenter35
    oCpxStringByCurves36.Inputs("Curves").Add oArcByFillet26
    oCpxStringByCurves36.Inputs("Curves").Add oLineByPoints33s
    Call GCEvaluate(oCpxStringByCurves36, pGeometricConstruction)
    oCpxStringByCurves36.Public = True

    Dim oSurfByLinearExtrusion37 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion37 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "037-SurfByLinearExtrusion")
    oSurfByLinearExtrusion37.Inputs("PlanarCrossSection").Add oCpxStringByCurves36
    oSurfByLinearExtrusion37.Inputs("ExtrusionLine").Add oLineByPoints12
    Call GCEvaluate(oSurfByLinearExtrusion37, pGeometricConstruction)
    oSurfByLinearExtrusion37.Public = True

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion37.Output
'    oGCMacro.Output("Boundary", 1) = oCpxStringByCurves36.Output
    oGCMacro.Output("PointX", 1) = oPointAlongCurve32.Output
    oGCMacro.Output("PointY", 1) = oPointAlongCurve34.Output
End Sub
