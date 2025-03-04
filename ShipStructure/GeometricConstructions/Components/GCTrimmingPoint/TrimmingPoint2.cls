VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrimmingPoint2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberAxis", "Select a curve", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point1", "Select a first point", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point2", "Select a second point", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("SlopeRatio", "SlopeRatio", GCDouble, UNIT_SCALAR, 0, 0, 0, 1.01)
    
    Call pGeometricConstructionDefinition.AddOutput(1, "Point")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim oGC1 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC1 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "001-PointAlongCurve")
    oGC1.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oGC1.Inputs("Point").Add pGeometricConstruction.Inputs("Point1").Item(1)
    oGC1.Parameter("Distance") = 0#
    oGC1.Parameter("TrackFlag") = 1
    oGC1.Evaluate
    oGC1.Public = True
    
    Dim oGC2 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC2 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "002-PointAlongCurve")
    oGC2.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oGC2.Inputs("Point").Add pGeometricConstruction.Inputs("Point2").Item(1)
    oGC2.Parameter("Distance") = 0#
    oGC2.Parameter("TrackFlag") = 1
    oGC2.Evaluate
    oGC2.Public = True
    
    Dim oGC3 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC3 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "003-PointAtCurveStart")
    oGC3.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oGC3.Evaluate
    oGC3.Public = True
    
    Dim oGC4 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC4 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "004-ParamDistMeasureParallel")
    oGC4.Inputs("Graphics").Add oGC3
    oGC4.Inputs("Graphics").Add oGC1
    oGC4.Evaluate
    oGC4.Public = True
    
    Dim oGC5 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC5 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "005-ParamDistMeasureParallel")
    oGC5.Inputs("Graphics").Add oGC3
    oGC5.Inputs("Graphics").Add oGC2
    oGC5.Evaluate
    oGC5.Public = True
    
    Dim oGC6 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC6 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "006-ParamDistFunctionWithNArgs")
    oGC6.Inputs("Args").Add oGC5
    oGC6.Inputs("Args").Add oGC4
    oGC6.Parameter("Operation") = 2
    oGC6.Evaluate
    oGC6.Public = True
    
    Dim oParamDouble6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble6 = oGCFactory.CreateEntity("ParamDouble", pPOM, "007-ParamDouble")
    oParamDouble6.Parameter("Value") = pGeometricConstruction.Parameter("SlopeRatio")
    oParamDouble6.Evaluate
  
    Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "008-ParamDistConstant")
    oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe")
    oParamDistConstant6.Evaluate
    
    Dim oParamDistFunctionWith2Args6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args6 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "009-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args6.Inputs("Arg1").Add oParamDistConstant6
    oParamDistFunctionWith2Args6.Inputs("Arg2").Add oParamDouble6
    oParamDistFunctionWith2Args6.Parameter("Value") = 0
    oParamDistFunctionWith2Args6.Parameter("Operation") = GCDivide
    oParamDistFunctionWith2Args6.Evaluate
    
    Dim oGC7 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC7 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "010-PointAlongCurve")
    oGC7.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oGC7.Parameter("Distance") = oGC6
    oGC7.Evaluate
    oGC7.Public = True
    
    Dim oGC8 As SP3DGeometricConstruction.GeometricConstruction
    Set oGC8 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "011-PointAlongCurve")
    oGC8.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis").Item(1)
    oGC8.Inputs("Point").Add oGC7
    oGC8.Parameter("Distance") = oParamDistFunctionWith2Args6
    oGC8.Evaluate
    oGC8.Public = True
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    oGCMacro.Output("Point", 1) = oGC8.Output
End Sub

