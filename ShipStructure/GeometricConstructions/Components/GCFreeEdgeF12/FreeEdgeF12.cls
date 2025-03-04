VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdgeF12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Enum Errors
    CUTBACK_DISTANCE_TOO_BIG = 1
End Enum

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select line for first axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select line for second axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("ArmLength", "ArmLength", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    ' define  errors
    Call pGeometricConstructionDefinition.AddErrorValue(CUTBACK_DISTANCE_TOO_BIG, "CutBackDistanceTooBig", "The anchor point for trimming plane is outside the member axis: decrease the CutbackDistance")
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oGC1 As IJGeometricConstruction
    Set oGC1 = oGCFactory.CreateEntity("CSByLines", pPOM, "001-CSByLines")
    oGC1.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oGC1.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oGC1.Parameter("AxesRoles") = 1
    oGC1.Parameter("CSOrientation") = 1
    oGC1.Evaluate
    oGC1.Public = True
    
    Dim oGC2 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC2 = oGCFactory.CreateEntity("CSByLines", pPOM, "002-CSByLines")
    oGC2.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oGC2.Inputs("AxisLine2").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oGC2.Parameter("AxesRoles") = 1
    oGC2.Parameter("CSOrientation") = 1
    oGC2.Evaluate
    oGC2.Public = True
    
    Dim oGC3 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC3 = oGCFactory.CreateEntity("PointFromCS", pPOM, "003-PointFromCS")
    oGC3.Inputs("CoordinateSystem").Add oGC1
    oGC3.Parameter("X") = 3.5
    oGC3.Parameter("Y") = 0.5
    oGC3.Parameter("Z") = 0#
    oGC3.Evaluate
    oGC3.Public = True

    Dim oGC4 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC4 = oGCFactory.CreateEntity("PointFromCS", pPOM, "004-PointFromCS")
    oGC4.Inputs("CoordinateSystem").Add oGC1
    oGC4.Parameter("X") = 3#
    oGC4.Parameter("Y") = 3#
    oGC4.Parameter("Z") = 0#
    oGC4.Evaluate
    oGC4.Public = True
    
    Dim oGC5 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5 = oGCFactory.CreateEntity("LineOrientate", pPOM, "005-LineOrientate")
    oGC5.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine1").Item(1)
    oGC5.Inputs("CoordinateSystem").Add oGC1
    oGC5.Evaluate
    oGC5.Public = True
       
    Dim oGC6 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC6 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "006-PlaneFromCS")
    oGC6.Inputs("CoordinateSystem").Add oGC1
    oGC6.Parameter("LookingAxis") = 3
    oGC6.Parameter("Offset") = 0#
    oGC6.Parameter("Range") = -10#
    oGC6.Evaluate
    oGC6.Public = True
    
    Dim oGC7 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC7 = oGCFactory.CreateEntity("LineOrientate", pPOM, "007-LineOrientate")
    oGC7.Inputs("Line").Add pGeometricConstruction.Inputs("AxisLine2").Item(1)
    oGC7.Inputs("CoordinateSystem").Add oGC2
    oGC7.Evaluate
    oGC7.Public = True
    
    Dim oGC8 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC8 = oGCFactory.CreateEntity("PointFromCS", pPOM, "008-PointFromCS")
    oGC8.Inputs("CoordinateSystem").Add oGC2
    oGC8.Parameter("X") = 3.5
    oGC8.Parameter("Y") = 0.5
    oGC8.Parameter("Z") = 0#
    oGC8.Evaluate
    oGC8.Public = True
    
    Dim oGC9 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC9 = oGCFactory.CreateEntity("LineFromCS", pPOM, "009-LineFromCS")
    oGC9.Inputs("CoordinateSystem").Add oGC2
    oGC9.Parameter("LookingAxis") = 3
    oGC9.Parameter("Length") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
    oGC9.Evaluate
    oGC9.Public = True
    
    Dim oCurveByProjection9a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection9a = oGCFactory.CreateEntity("CurveByProjection", pPOM, "010-CurveByProjection")
    oCurveByProjection9a.Inputs("Curve").Add oGC5
    oCurveByProjection9a.Inputs("Surface").Add oGC6
    oCurveByProjection9a.Inputs("Line").Add oGC9
    oCurveByProjection9a.Parameter("TrackFlag") = 1
    oCurveByProjection9a.Public = True
    oCurveByProjection9a.Evaluate
    
    Dim oCurveOrientate9b As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate9b = oGCFactory.CreateEntity("CurveOrientate", pPOM, "011-CurveOrientate")
    oCurveOrientate9b.Inputs("Curve").Add oCurveByProjection9a
    oCurveOrientate9b.Inputs("CoordinateSystem").Add oGC1
    oCurveOrientate9b.Public = True
    oCurveOrientate9b.Evaluate
    
    Dim oCurveByProjection9c As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection9c = oGCFactory.CreateEntity("CurveByProjection", pPOM, "012-CurveByProjection")
    oCurveByProjection9c.Inputs("Curve").Add oGC7
    oCurveByProjection9c.Inputs("Surface").Add oGC6
    oCurveByProjection9c.Inputs("Line").Add oGC9
    oCurveByProjection9c.Parameter("TrackFlag") = 1
    oCurveByProjection9c.Public = True
    oCurveByProjection9c.Evaluate
    
    Dim oCurveOrientate9d As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate9d = oGCFactory.CreateEntity("CurveOrientate", pPOM, "013-CurveOrientate")
    oCurveOrientate9d.Inputs("Curve").Add oCurveByProjection9c
    oCurveOrientate9d.Inputs("CoordinateSystem").Add oGC2
    oCurveOrientate9d.Public = True
    oCurveOrientate9d.Evaluate
    
    Dim oGC10a As SP3DGeometricConstruction.GeometricConstruction
    Set oGC10a = oGCFactory.CreateEntity("LineExtend", pPOM, "014a-LineExtend")
    oGC10a.Inputs("Line").Add oCurveOrientate9b
    oGC10a.Inputs("Curve1").Add oCurveOrientate9d
    oGC10a.Parameter("Context1") = 1
    oGC10a.Parameter("Context2") = 4
    oGC10a.Evaluate
    oGC10a.Public = True
    
    Dim oGC10b As SP3DGeometricConstruction.GeometricConstruction
    Set oGC10b = oGCFactory.CreateEntity("LineExtend", pPOM, "014b-LineExtend")
    oGC10b.Inputs("Line").Add oCurveOrientate9d
    oGC10b.Inputs("Curve1").Add oCurveOrientate9b
    oGC10b.Parameter("Context1") = 1
    oGC10b.Parameter("Context2") = 4
    oGC10b.Evaluate
    oGC10b.Public = True

    Dim oGC10 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC10 = oGCFactory.CreateEntity("LineAtBissector", pPOM, "014-LineAtBissector")
    oGC10.Inputs("Line1").Add oGC10a ' oCurveOrientate9b
    oGC10.Inputs("Line2").Add oGC10b ' oCurveOrientate9d
    oGC10.Inputs("TrackPoint").Add oGC4
    oGC10.Parameter("SweepAngle") = 1
    oGC10.Parameter("TrackFlag") = 1
    oGC10.Evaluate
    oGC10.Public = True
    
    Dim oGC11 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC11 = oGCFactory.CreateEntity("LineParallel", pPOM, "015-LineParallel")
    oGC11.Inputs("Line").Add oCurveOrientate9b
    oGC11.Inputs("Plane").Add oGC6
    oGC11.Inputs("TrackPoint").Add oGC3
    oGC11.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oGC11.Parameter("TrackFlag") = 1
    oGC11.Evaluate
    oGC11.Public = True
    
    Dim oGC12 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC12 = oGCFactory.CreateEntity("LineParallel", pPOM, "016-LineParallel")
    oGC12.Inputs("Line").Add oCurveOrientate9d
    oGC12.Inputs("Plane").Add oGC6
    oGC12.Inputs("TrackPoint").Add oGC8
    oGC12.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oGC12.Parameter("TrackFlag") = 1
    oGC12.Evaluate
    oGC12.Public = True
    
    Dim oGC13 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC13 = oGCFactory.CreateEntity("LineExtend", pPOM, "017-LineExtend")
    oGC13.Inputs("Line").Add oGC11
    oGC13.Inputs("Curve1").Add oGC12
    oGC13.Parameter("Context1") = 1
    oGC13.Parameter("Context2") = 4
    oGC13.Evaluate
    oGC13.Public = True
    
    Dim oGC14 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC14 = oGCFactory.CreateEntity("LineExtend", pPOM, "018-LineExtend")
    oGC14.Inputs("Line").Add oGC12
    oGC14.Inputs("Curve1").Add oGC11
    oGC14.Parameter("Context1") = 1
    oGC14.Parameter("Context2") = 4
    oGC14.Evaluate
    oGC14.Public = True
    
    Dim oGC15 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC15 = oGCFactory.CreateEntity("PointByCurves", pPOM, "019-PointByCurves")
    oGC15.Inputs("Curve1").Add oGC13
    oGC15.Inputs("Curve2").Add oGC14
    oGC15.Parameter("TrackFlag") = 1
    oGC15.Evaluate
    oGC15.Public = True
    
    Dim oGC16 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC16 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "020-PointAlongCurve")
    oGC16.Inputs("Curve").Add oGC10
    oGC16.Inputs("Point").Add oGC15
    oGC16.Inputs("TrackPoint").Add oGC4
    oGC16.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("ArmLength"))
    oGC16.Parameter("TrackFlag") = 1
    oGC16.Evaluate
    oGC16.Public = True
    
    Dim oGC17 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC17 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "021-LineAtAngle")
    oGC17.Inputs("Line").Add oGC10
    oGC17.Inputs("Point").Add oGC16
    oGC17.Inputs("Plane").Add oGC6
    oGC17.Inputs("TrackPoint").Add oGC3
    oGC17.Parameter("Angle") = 1.570796
    oGC17.Parameter("TrackFlag") = 1
    oGC17.Evaluate
    oGC17.Public = True
    
    Dim oGC18 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC18 = oGCFactory.CreateEntity("LineExtend", pPOM, "022-LineExtend")
    oGC18.Inputs("Line").Add oGC17
    oGC18.Inputs("Curve1").Add oGC11
    oGC18.Parameter("Context1") = 1
    oGC18.Parameter("Context2") = 4
    oGC18.Evaluate
    oGC18.Public = True
    
    Dim oGC19 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC19 = oGCFactory.CreateEntity("LineExtend", pPOM, "023-LineExtend")
    oGC19.Inputs("Line").Add oGC11
    oGC19.Inputs("Curve1").Add oGC18
    oGC19.Parameter("Context1") = 1
    oGC19.Parameter("Context2") = 4
    oGC19.Evaluate
    oGC19.Public = True
    
    Dim oGC20 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC20 = oGCFactory.CreateEntity("LineExtend", pPOM, "024-LineExtend")
    oGC20.Inputs("Line").Add oGC12
    oGC20.Inputs("Curve1").Add oGC17
    oGC20.Parameter("Context1") = 1
    oGC20.Parameter("Context2") = 1
    oGC20.Evaluate
    oGC20.Public = True
    
    Dim oGC21 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC21 = oGCFactory.CreateEntity("LineExtend", pPOM, "025-LineExtend")
    oGC21.Inputs("Line").Add oGC18
    oGC21.Inputs("Curve1").Add oGC12
    oGC21.Parameter("Context1") = 4
    oGC21.Parameter("Context2") = 1
    oGC21.Evaluate
    oGC21.Public = True
   
    Dim oGC22 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC22 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "026-ArcByFillet")
    oGC22.Inputs("Curve1").Add oGC19
    oGC22.Inputs("Curve2").Add oGC21
    oGC22.Parameter("Radius") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    oGC22.Evaluate
    oGC22.Public = True
    
    Dim oGC23 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC23 = oGCFactory.CreateEntity("LineExtend", pPOM, "027-LineExtend")
    oGC23.Inputs("Line").Add oGC19
    oGC23.Inputs("Curve1").Add oGC22
    oGC23.Parameter("Context1") = 1
    oGC23.Parameter("Context2") = 1
    oGC23.Evaluate
    oGC23.Public = True
    
    Dim oGC24 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC24 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "028-ArcByFillet")
    oGC24.Inputs("Curve1").Add oGC21
    oGC24.Inputs("Curve2").Add oGC20
    oGC24.Parameter("Radius") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    oGC24.Evaluate
    oGC24.Public = True
    
    Dim oGC25 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC25 = oGCFactory.CreateEntity("LineExtend", pPOM, "029-LineExtend")
    oGC25.Inputs("Line").Add oGC20
    oGC25.Inputs("Curve1").Add oGC24
    oGC25.Parameter("Context1") = 1
    oGC25.Parameter("Context2") = 1
    oGC25.Evaluate
    oGC25.Public = True
    
    Dim oGC26 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC26 = oGCFactory.CreateEntity("LineExtend", pPOM, "030-LineExtend")
    oGC26.Inputs("Line").Add oGC21
    oGC26.Inputs("Curve1").Add oGC24
    oGC26.Inputs("Curve2").Add oGC22
    oGC26.Parameter("Context1") = 4
    oGC26.Parameter("Context2") = 1
    oGC26.Evaluate
    oGC26.Public = True

    Dim oGC27 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC27 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "031-PointAtCurveStart")
    oGC27.Inputs("Curve").Add oGC23
    oGC27.Evaluate
    
    Dim oGC28 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC28 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "032-PointAtCurveStart")
    oGC28.Inputs("Curve").Add oGC25
    oGC28.Evaluate
    oGC28.Public = True
    
    Dim oGC29 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC29 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "033-PointAlongCurve")
    oGC29.Inputs("Curve").Add oGC11
    oGC29.Inputs("Point").Add oGC27
    oGC29.Inputs("TrackPoint").Add oGC3
    oGC29.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oGC29.Parameter("TrackFlag") = 1
    oGC29.Evaluate
    oGC29.Public = True
    
    If Not IsPositionInsideCurve(Position_FromPoint(oGC29.Output), oGC11.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)

    Dim oGC30 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC30 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "034-PointAlongCurve")
    oGC30.Inputs("Curve").Add oGC12
    oGC30.Inputs("Point").Add oGC28
    oGC30.Inputs("TrackPoint").Add oGC8
    oGC30.Parameter("Distance") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oGC30.Parameter("TrackFlag") = 1
    oGC30.Evaluate
    oGC30.Public = True

    If Not IsPositionInsideCurve(Position_FromPoint(oGC30.Output), oGC12.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)
    
    Dim oGC31 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC31 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "035-CpxStringByCurves")
    oGC31.Inputs("Curves").Add oGC23
    oGC31.Inputs("Curves").Add oGC22
    oGC31.Inputs("Curves").Add oGC26
    oGC31.Inputs("Curves").Add oGC24
    oGC31.Inputs("Curves").Add oGC25
    oGC31.Evaluate
    oGC31.Public = True
    
    Dim oGC32 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC32 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "036-SurfByLinearExtrusion")
    oGC32.Inputs("PlanarCrossSection").Add oGC31
    oGC32.Inputs("ExtrusionLine").Add oGC9
    oGC32.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    oGCMacro.Output("Boundary", 1) = oGC32.Output
    oGCMacro.Output("PointY", 1) = oGC30.Output
    oGCMacro.Output("PointX", 1) = oGC29.Output
End Sub
