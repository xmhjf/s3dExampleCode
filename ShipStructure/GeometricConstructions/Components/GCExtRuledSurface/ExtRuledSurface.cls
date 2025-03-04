VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtRuledSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Curve1", "Select a first curve", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Curve2", "Select a seond curve", "IJCurve OR IJWireBody", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointS1")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointE1")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointS2")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointE2")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oGC1 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC1 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "001-PointAtCurveMiddle")
    oGC1.Inputs("Curve").Add pGeometricConstruction.Inputs("Curve1").Item(1)
    oGC1.Evaluate
    oGC1.Public = True
    
    Dim oGC2 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC2 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "002-PointAtCurveMiddle")
    oGC2.Inputs("Curve").Add pGeometricConstruction.Inputs("Curve2").Item(1)
    oGC2.Evaluate
    oGC2.Public = True
        
    Dim oGC3 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC3 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "003-PointAtCurveEnd")
    oGC3.Inputs("Curve").Add pGeometricConstruction.Inputs("Curve1").Item(1)
    oGC3.Evaluate
    oGC3.Public = True
        
    Dim oGC4bis As SP3DGeometricConstruction.GeometricConstruction
    Set oGC4bis = oGCFactory.CreateEntity("CSByPoints", pPOM, "004-CSByPoints")
    oGC4bis.Inputs("OriginPoint").Add oGC1
    oGC4bis.Inputs("AxisPoint1").Add oGC3
    oGC4bis.Inputs("AxisPoint2").Add oGC2
    oGC4bis.Parameter("AxesRoles") = 1
    oGC4bis.Parameter("CSOrientation") = 1
    oGC4bis.Evaluate
    oGC4bis.Public = True
    
    Dim oGC5bis As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5bis = oGCFactory.CreateEntity("LineOrientate", pPOM, "005-LineOrientate")
    oGC5bis.Inputs("Line").Add pGeometricConstruction.Inputs("Curve1").Item(1)
    oGC5bis.Inputs("CoordinateSystem").Add oGC4bis
    oGC5bis.Evaluate
    oGC5bis.Public = True

    Dim oGC6bis As SP3DGeometricConstruction.GeometricConstruction
    Set oGC6bis = oGCFactory.CreateEntity("LineOrientate", pPOM, "006-LineOrientate")
    oGC6bis.Inputs("Line").Add pGeometricConstruction.Inputs("Curve2").Item(1)
    oGC6bis.Inputs("CoordinateSystem").Add oGC4bis
    oGC6bis.Evaluate
    oGC6bis.Public = True

    Dim oGC7bis As SP3DGeometricConstruction.GeometricConstruction
    Set oGC7bis = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "007-PointAtCurveStart")
    oGC7bis.Inputs("Curve").Add oGC5bis
    oGC7bis.Evaluate
    oGC7bis.Public = True
    
    Dim oGC8bis As SP3DGeometricConstruction.GeometricConstruction
    Set oGC8bis = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "008-PointAtCurveEnd")
    oGC8bis.Inputs("Curve").Add oGC5bis
    oGC8bis.Evaluate
    oGC8bis.Public = True
    
    Dim oGC3c As SP3DGeometricConstruction.GeometricConstruction
    Set oGC3c = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "009-PointAtCurveStart")
    oGC3c.Inputs("Curve").Add oGC6bis
    oGC3c.Evaluate
    oGC3c.Public = True
    
    Dim oGC4 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC4 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "010-PointAtCurveEnd")
    oGC4.Inputs("Curve").Add oGC6bis
    oGC4.Evaluate
    oGC4.Public = True
       
    Dim oGC5 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "011-ParamDistConstant")
    oGC5.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oGC5.Evaluate
    oGC5.Public = True
    
    Dim oGC6 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "012-PointAlongCurve")
    oGC6.Inputs("Curve").Add oGC5bis
    oGC6.Inputs("Point").Add oGC7bis
    oGC6.Inputs("TrackPoint").Add oGC8bis
    oGC6.Parameter("Distance") = oGC5
    oGC6.Parameter("TrackFlag") = 2
    oGC6.Evaluate
    oGC6.Public = True
    
    Dim oGC7 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC7 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "013-PointAlongCurve")
    oGC7.Inputs("Curve").Add oGC5bis
    oGC7.Inputs("Point").Add oGC8bis
    oGC7.Inputs("TrackPoint").Add oGC7bis
    oGC7.Parameter("Distance") = oGC5
    oGC7.Parameter("TrackFlag") = 2
    oGC7.Evaluate
    oGC7.Public = True
    
    Dim oGC8 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC8 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "014-PointAlongCurve")
    oGC8.Inputs("Curve").Add oGC6bis
    oGC8.Inputs("Point").Add oGC3c
    oGC8.Inputs("TrackPoint").Add oGC4
    oGC8.Parameter("Distance") = oGC5
    oGC8.Parameter("TrackFlag") = 2
    oGC8.Evaluate
    oGC8.Public = True
    
    Dim oGC9 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC9 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "015-PointAlongCurve")
    oGC9.Inputs("Curve").Add oGC6bis
    oGC9.Inputs("Point").Add oGC4
    oGC9.Inputs("TrackPoint").Add oGC3c
    oGC9.Parameter("Distance") = oGC5
    oGC9.Parameter("TrackFlag") = 2
    oGC9.Evaluate
    oGC9.Public = True
    
    Dim oGC10 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC10 = oGCFactory.CreateEntity("LineByPoints", pPOM, "016-LineByPoints")
    oGC10.Inputs("StartPoint").Add oGC7
    oGC10.Inputs("EndPoint").Add oGC9
    oGC10.Evaluate
    oGC10.Public = True
    
    Dim oGC11 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "017-PointAlongCurve")
    oGC11.Inputs("Curve").Add oGC10
    oGC11.Inputs("Point").Add oGC7
    oGC11.Inputs("TrackPoint").Add oGC9
    oGC11.Parameter("Distance") = oGC5
    oGC11.Parameter("TrackFlag") = 2
    oGC11.Evaluate
    oGC11.Public = True
    
    Dim oGC12 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC12 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "018-PointAlongCurve")
    oGC12.Inputs("Curve").Add oGC10
    oGC12.Inputs("Point").Add oGC9
    oGC12.Inputs("TrackPoint").Add oGC7
    oGC12.Parameter("Distance") = oGC5
    oGC12.Parameter("TrackFlag") = 2
    oGC12.Evaluate
    oGC12.Public = True
    
    Dim oGC13 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC13 = oGCFactory.CreateEntity("LineByPoints", pPOM, "019-LineByPoints")
    oGC13.Inputs("StartPoint").Add oGC6
    oGC13.Inputs("EndPoint").Add oGC8
    oGC13.Evaluate
    oGC13.Public = True
        
    Dim oGC15 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC15 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "020-PointAlongCurve")
    oGC15.Inputs("Curve").Add oGC13
    oGC15.Inputs("Point").Add oGC6
    oGC15.Inputs("TrackPoint").Add oGC8
    oGC15.Parameter("Distance") = oGC5
    oGC15.Parameter("TrackFlag") = 2
    oGC15.Evaluate
    oGC15.Public = True
    
    Dim oGC14 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC14 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "021-PointAlongCurve")
    oGC14.Inputs("Curve").Add oGC13
    oGC14.Inputs("Point").Add oGC8
    oGC14.Inputs("TrackPoint").Add oGC6
    oGC14.Parameter("Distance") = oGC5
    oGC14.Parameter("TrackFlag") = 2
    oGC14.Evaluate
    oGC14.Public = True

    Dim oGC16 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC16 = oGCFactory.CreateEntity("LineByPoints", pPOM, "022-LineByPoints")
    oGC16.Inputs("StartPoint").Add oGC14
    oGC16.Inputs("EndPoint").Add oGC12
    oGC16.Evaluate
    oGC16.Public = True
    
    Dim oGC17 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC17 = oGCFactory.CreateEntity("LineByPoints", pPOM, "023-LineByPoints")
    oGC17.Inputs("StartPoint").Add oGC15
    oGC17.Inputs("EndPoint").Add oGC11
    oGC17.Evaluate
    oGC17.Public = True
    
    Dim oGC18 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC18 = oGCFactory.CreateEntity("SurfByRuling", pPOM, "024-SurfByRuling")
    oGC18.Inputs("Curve1").Add oGC17
    oGC18.Inputs("Curve2").Add oGC16
    oGC18.Evaluate
    oGC18.Public = True
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    oGCMacro.Output("PointS1", 1) = oGC7bis.Output
    oGCMacro.Output("PointE1", 1) = oGC8bis.Output
    oGCMacro.Output("PointS2", 1) = oGC3c.Output
    oGCMacro.Output("PointE2", 1) = oGC4.Output
    oGCMacro.Output("Support", 1) = oGC18.Output
End Sub
