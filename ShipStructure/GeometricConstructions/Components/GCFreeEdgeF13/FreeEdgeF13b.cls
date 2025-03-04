VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdgeF13b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Const RADIUS_TOO_BIG = 6
Private Enum Errors
    FILLET_TOO_BIG = 1
    CUTBACK_DISTANCE_TOO_BIG = 2
End Enum
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("AxisLine1", "Select line for first axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("AxisLine2", "Select line for second axis", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Fillet", "Fillet", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("CutBack2", "CutBack2", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter("Side", "Side", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Inside", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Outside", 2)
   
    ' define  errors
    Call pGeometricConstructionDefinition.AddErrorValue(FILLET_TOO_BIG, "FilletTooBig", "Fillet cannot be created: decrease the Fillet")
    Call pGeometricConstructionDefinition.AddErrorValue(CUTBACK_DISTANCE_TOO_BIG, "CutBackDistanceTooBig", "The anchor point for trimming plane is outside the member axis: decrease the CutbackDistance")
    
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim oGC1 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC1 = oGCFactory.CreateEntity("CSByLines", pPOM, "001-CSByLines")
    oGC1.Inputs("AxisLine1").Add pGeometricConstruction.Input("AxisLine1")
    oGC1.Inputs("AxisLine2").Add pGeometricConstruction.Input("AxisLine2")
    oGC1.Parameter("AxesRoles") = 1
    oGC1.Parameter("CSOrientation") = 1
    oGC1.Evaluate
    oGC1.Public = True
    
    Dim oGC2 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC2 = oGCFactory.CreateEntity("CSByLines", pPOM, "002-CSByLines")
    oGC2.Inputs("AxisLine1").Add pGeometricConstruction.Input("AxisLine2")
    oGC2.Inputs("AxisLine2").Add pGeometricConstruction.Input("AxisLine1")
    oGC2.Parameter("AxesRoles") = 1
    oGC2.Parameter("CSOrientation") = 1
    oGC2.Evaluate
    oGC2.Public = True
    
    Dim oGC3 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oGC3.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oGC3.Evaluate
    oGC3.Public = True
    
    Dim oGC4 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oGC4.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Fillet"))
    oGC4.Evaluate
    oGC4.Public = True
    
    Dim oGC5 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oGC5.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oGC5.Evaluate
    oGC5.Public = True
    
    Dim oGC5b As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5b = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "006-ParamDistConstant")
    oGC5b.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("CutBack2"))
    oGC5b.Evaluate
    oGC5b.Public = True
    
    Dim oGC6 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC6 = oGCFactory.CreateEntity("PointFromCS", pPOM, "007-PointFromCS")
    oGC6.Inputs("CoordinateSystem").Add oGC1
    oGC6.Parameter("X") = 6#
    If pGeometricConstruction.Parameter("Side") <> 2 Then
        oGC6.Parameter("Y") = 0.1
    Else
        oGC6.Parameter("Y") = -0.1
    End If
    oGC6.Parameter("Z") = 0#
    oGC6.Evaluate
    oGC6.Public = True
    
    Dim oGC7 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC7 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "008-PlaneFromCS")
    oGC7.Inputs("CoordinateSystem").Add oGC1
    oGC7.Parameter("LookingAxis") = 3
    oGC7.Parameter("Offset") = 0#
    oGC7.Parameter("Range") = -10#
    oGC7.Evaluate
    oGC7.Public = True
    
    Dim oGC8 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC8 = oGCFactory.CreateEntity("LineFromCS", pPOM, "009-LineFromCS")
    oGC8.Inputs("CoordinateSystem").Add oGC1
    oGC8.Parameter("LookingAxis") = 3
    oGC8.Parameter("Length") = -CDbl(pGeometricConstruction.Parameter("RibbonHeight")) / 2
    oGC8.Evaluate
    oGC8.Public = True
    
    Dim oGC9 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "010-PointFromCS")
    oGC9.Inputs("CoordinateSystem").Add oGC2
    oGC9.Parameter("X") = 6#
    If pGeometricConstruction.Parameter("Side") <> 2 Then
        oGC9.Parameter("Y") = 0.1
    Else
        oGC9.Parameter("Y") = -0.1
    End If
    oGC9.Parameter("Z") = 0#
    oGC9.Evaluate
    oGC9.Public = True
    
    Dim oCurveByProjection7a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection7a = oGCFactory.CreateEntity("CurveByProjection", pPOM, "011-CurveByProjection")
    oCurveByProjection7a.Inputs("Curve").Add pGeometricConstruction.Input("AxisLine1")
    oCurveByProjection7a.Inputs("Surface").Add oGC7
    oCurveByProjection7a.Inputs("Line").Add oGC8
    oCurveByProjection7a.Parameter("TrackFlag") = 1
    oCurveByProjection7a.Public = True
    oCurveByProjection7a.Evaluate
   
    Dim oCurveOrientate7a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate7a = oGCFactory.CreateEntity("CurveOrientate", pPOM, "012-CurveOrientate")
    oCurveOrientate7a.Inputs("Curve").Add oCurveByProjection7a
    oCurveOrientate7a.Inputs("CoordinateSystem").Add oGC1
    oCurveOrientate7a.Public = True
    oCurveOrientate7a.Evaluate
    
    Dim oCurveByProjection7b As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection7b = oGCFactory.CreateEntity("CurveByProjection", pPOM, "013-CurveByProjection")
    oCurveByProjection7b.Inputs("Curve").Add pGeometricConstruction.Input("AxisLine2")
    oCurveByProjection7b.Inputs("Surface").Add oGC7
    oCurveByProjection7b.Inputs("Line").Add oGC8
    oCurveByProjection7b.Parameter("TrackFlag") = 1
    oCurveByProjection7b.Public = True
    oCurveByProjection7b.Evaluate
   
    Dim oCurveOrientate7b As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate7b = oGCFactory.CreateEntity("CurveOrientate", pPOM, "014-CurveOrientate")
    oCurveOrientate7b.Inputs("Curve").Add oCurveByProjection7b
    oCurveOrientate7b.Inputs("CoordinateSystem").Add oGC2
    oCurveOrientate7b.Public = True
    oCurveOrientate7b.Evaluate
    
    Dim oGC10 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC10 = oGCFactory.CreateEntity("LineParallel", pPOM, "015-LineParallel")
    oGC10.Inputs("Line").Add oCurveOrientate7a
    oGC10.Inputs("Plane").Add oGC7
    oGC10.Inputs("TrackPoint").Add oGC6
    oGC10.Parameter("Distance") = oGC3
    oGC10.Parameter("TrackFlag") = 1
    oGC10.Evaluate
    oGC10.Public = True
    
    Dim oGC11 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC11 = oGCFactory.CreateEntity("LineParallel", pPOM, "016-LineParallel")
    oGC11.Inputs("Line").Add oCurveOrientate7b
    oGC11.Inputs("Plane").Add oGC7
    oGC11.Inputs("TrackPoint").Add oGC9
    oGC11.Parameter("Distance") = oGC3
    oGC11.Parameter("TrackFlag") = 1
    oGC11.Evaluate
    oGC11.Public = True
    
    Dim oGC12 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC12 = oGCFactory.CreateEntity("LineExtend", pPOM, "017-LineExtend")
    oGC12.Inputs("Line").Add oGC11
    oGC12.Inputs("Curve1").Add oGC10
    oGC12.Parameter("Context1") = 1
    oGC12.Parameter("Context2") = 1
    oGC12.Evaluate
    oGC12.Public = True
    
    Dim oGC13 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC13 = oGCFactory.CreateEntity("LineExtend", pPOM, "018-LineExtend")
    oGC13.Inputs("Line").Add oGC10
    oGC13.Inputs("Curve1").Add oGC11
    oGC13.Parameter("Context1") = 1
    oGC13.Parameter("Context2") = 1
    oGC13.Evaluate
    oGC13.Public = True
    
    Dim oGC14 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC14 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "019-ArcByFillet")
    oGC14.Inputs("Curve1").Add oGC12
    oGC14.Inputs("Curve2").Add oGC13
    oGC14.Parameter("Radius") = oGC4
    Call GCEvaluate(oGC14, pGeometricConstruction, -1, FILLET_TOO_BIG)
    oGC14.Public = True
    
    Dim oGC15a As SP3DGeometricConstruction.GeometricConstruction
    Set oGC15a = oGCFactory.CreateEntity("LineExtend", pPOM, "020-LineExtend")
    oGC15a.Inputs("Line").Add oGC13
    oGC15a.Inputs("Curve1").Add oGC14
    oGC15a.Parameter("Context1") = 1
    oGC15a.Parameter("Context2") = 1
    oGC15a.Evaluate
    oGC15a.Public = True
    
    Dim oGC16a As SP3DGeometricConstruction.GeometricConstruction
    Set oGC16a = oGCFactory.CreateEntity("LineExtend", pPOM, "021-LineExtend")
    oGC16a.Inputs("Line").Add oGC12
    oGC16a.Inputs("Curve1").Add oGC14
    oGC16a.Parameter("Context1") = 1
    oGC16a.Parameter("Context2") = 1
    oGC16a.Evaluate
    oGC16a.Public = True
    
    Dim oGC15 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC15 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "022-PointAtCurveStart")
    oGC15.Inputs("Curve").Add oGC15a
    oGC15.Evaluate
    oGC15.Public = True
    
    Dim oGC16 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC16 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "023-PointAtCurveStart")
    oGC16.Inputs("Curve").Add oGC16a
    oGC16.Evaluate
    oGC16.Public = True
    
    Dim oGC21 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC21 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "024-PointAlongCurve")
    oGC21.Inputs("Curve").Add oGC15a
    oGC21.Inputs("Point").Add oGC15
    oGC21.Inputs("TrackPoint").Add oGC6
    oGC21.Parameter("Distance") = oGC5
    oGC21.Parameter("TrackFlag") = 1
    oGC21.Evaluate
    oGC21.Public = True
    
    If Not IsPositionInsideCurve(Position_FromPoint(oGC21.Output), oGC15a.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)
    
    Dim oGC22 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC22 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "025-PointAlongCurve")
    oGC22.Inputs("Curve").Add oGC16a
    oGC22.Inputs("Point").Add oGC16
    oGC22.Inputs("TrackPoint").Add oGC9
    oGC22.Parameter("Distance") = oGC5b
    oGC22.Parameter("TrackFlag") = 1
    oGC22.Evaluate
    oGC22.Public = True
    
    If Not IsPositionInsideCurve(Position_FromPoint(oGC22.Output), oGC16a.Output) Then Call GCProcessError(pGeometricConstruction, , CUTBACK_DISTANCE_TOO_BIG)
    
    Dim oGC23 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC23 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "026-CpxStringByCurves")
    oGC23.Inputs("Curves").Add oGC15a
    oGC23.Inputs("Curves").Add oGC14
    oGC23.Inputs("Curves").Add oGC16a
    oGC23.Evaluate
    oGC23.Public = True
    
    Dim oGC24 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC24 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "027-SurfByLinearExtrusion")
    oGC24.Inputs("PlanarCrossSection").Add oGC23
    oGC24.Inputs("ExtrusionLine").Add oGC8
    oGC24.Evaluate
    oGC24.Public = True
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    oGCMacro.Output("PointY", 1) = oGC22.Output
    oGCMacro.Output("PointX", 1) = oGC21.Output
    oGCMacro.Output("Boundary", 1) = oGC24.Output
End Sub

