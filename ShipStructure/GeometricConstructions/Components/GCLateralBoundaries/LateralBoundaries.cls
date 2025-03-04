VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LateralBoundaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem1", "Select coordinate system 1", "{249077B1-DB8C-11D2-8360-0800367F3D03}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("PointS1", "Select start point 1", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("PointE1", "Select end point 1", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem2", "Select coordinate system 2", "{249077B1-DB8C-11D2-8360-0800367F3D03}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("PointS2", "Select start point 2", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("PointE2", "Select end point 2", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oPointFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS1 = oGCFactory.CreateEntity("PointFromCS", pPOM, "001-PointFromCS")
    oPointFromCS1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1)
    oPointFromCS1.Inputs("Point").Add pGeometricConstruction.Inputs("PointS1").Item(1)
    oPointFromCS1.Parameter("X") = 0#
    oPointFromCS1.Parameter("Y") = 0#
    oPointFromCS1.Parameter("Z") = 0.2
    oPointFromCS1.Evaluate

    Dim oPointFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS2 = oGCFactory.CreateEntity("PointFromCS", pPOM, "002-PointFromCS")
    oPointFromCS2.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1)
    oPointFromCS2.Inputs("Point").Add pGeometricConstruction.Inputs("PointS1").Item(1)
    oPointFromCS2.Parameter("X") = 0#
    oPointFromCS2.Parameter("Y") = 0#
    oPointFromCS2.Parameter("Z") = -0.2
    oPointFromCS2.Evaluate

    Dim oPointFromCS3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS3 = oGCFactory.CreateEntity("PointFromCS", pPOM, "003-PointFromCS")
    oPointFromCS3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1)
    oPointFromCS3.Inputs("Point").Add pGeometricConstruction.Inputs("PointE1").Item(1)
    oPointFromCS3.Parameter("X") = 0#
    oPointFromCS3.Parameter("Y") = 0#
    oPointFromCS3.Parameter("Z") = 0.2
    oPointFromCS3.Evaluate

    Dim oPointFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS4 = oGCFactory.CreateEntity("PointFromCS", pPOM, "004-PointFromCS")
    oPointFromCS4.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1)
    oPointFromCS4.Inputs("Point").Add pGeometricConstruction.Inputs("PointE1").Item(1)
    oPointFromCS4.Parameter("X") = 0#
    oPointFromCS4.Parameter("Y") = 0#
    oPointFromCS4.Parameter("Z") = -0.2
    oPointFromCS4.Evaluate

    Dim oPointFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS5 = oGCFactory.CreateEntity("PointFromCS", pPOM, "005-PointFromCS")
    oPointFromCS5.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1)
    oPointFromCS5.Inputs("Point").Add pGeometricConstruction.Inputs("PointS2").Item(1)
    oPointFromCS5.Parameter("X") = 0#
    oPointFromCS5.Parameter("Y") = 0#
    oPointFromCS5.Parameter("Z") = 0.2
    oPointFromCS5.Evaluate

    Dim oPointFromCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS6 = oGCFactory.CreateEntity("PointFromCS", pPOM, "006-PointFromCS")
    oPointFromCS6.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1)
    oPointFromCS6.Inputs("Point").Add pGeometricConstruction.Inputs("PointS2").Item(1)
    oPointFromCS6.Parameter("X") = 0#
    oPointFromCS6.Parameter("Y") = 0#
    oPointFromCS6.Parameter("Z") = -0.2
    oPointFromCS6.Evaluate

    Dim oPointFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS7 = oGCFactory.CreateEntity("PointFromCS", pPOM, "007-PointFromCS")
    oPointFromCS7.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1)
    oPointFromCS7.Inputs("Point").Add pGeometricConstruction.Inputs("PointE2").Item(1)
    oPointFromCS7.Parameter("X") = 0#
    oPointFromCS7.Parameter("Y") = 0#
    oPointFromCS7.Parameter("Z") = 0.2
    oPointFromCS7.Evaluate

    Dim oPointFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS8 = oGCFactory.CreateEntity("PointFromCS", pPOM, "008-PointFromCS")
    oPointFromCS8.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1)
    oPointFromCS8.Inputs("Point").Add pGeometricConstruction.Inputs("PointE2").Item(1)
    oPointFromCS8.Parameter("X") = 0#
    oPointFromCS8.Parameter("Y") = 0#
    oPointFromCS8.Parameter("Z") = -0.2
    oPointFromCS8.Evaluate

    Dim oLineByPoints9 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints9 = oGCFactory.CreateEntity("LineByPoints", pPOM, "009-LineByPoints")
    oLineByPoints9.Inputs("StartPoint").Add oPointFromCS2
    oLineByPoints9.Inputs("EndPoint").Add oPointFromCS1
    oLineByPoints9.Evaluate

    Dim oLineByPoints10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints10 = oGCFactory.CreateEntity("LineByPoints", pPOM, "010-LineByPoints")
    oLineByPoints10.Inputs("StartPoint").Add oPointFromCS4
    oLineByPoints10.Inputs("EndPoint").Add oPointFromCS3
    oLineByPoints10.Evaluate

    Dim oLineByPoints11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints11 = oGCFactory.CreateEntity("LineByPoints", pPOM, "011-LineByPoints")
    oLineByPoints11.Inputs("StartPoint").Add oPointFromCS6
    oLineByPoints11.Inputs("EndPoint").Add oPointFromCS5
    oLineByPoints11.Evaluate

    Dim oLineByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints12 = oGCFactory.CreateEntity("LineByPoints", pPOM, "012-LineByPoints")
    oLineByPoints12.Inputs("StartPoint").Add oPointFromCS8
    oLineByPoints12.Inputs("EndPoint").Add oPointFromCS7
    oLineByPoints12.Evaluate

    Dim oExtRuledSurface13 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface13 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "013-ExtRuledSurface")
    oExtRuledSurface13.Inputs("Curve1").Add oLineByPoints11
    oExtRuledSurface13.Inputs("Curve2").Add oLineByPoints9
    oExtRuledSurface13.Parameter("Extension") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oExtRuledSurface13.Evaluate

    Dim oExtRuledSurface14 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface14 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "014-ExtRuledSurface")
    oExtRuledSurface14.Inputs("Curve1").Add oLineByPoints12
    oExtRuledSurface14.Inputs("Curve2").Add oLineByPoints10
    oExtRuledSurface14.Parameter("Extension") = CDbl(pGeometricConstruction.Parameter("Extension"))
    oExtRuledSurface14.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", 1) = oExtRuledSurface14.Output("Support", 1)
    oGCMacro.Output("Boundary", 2) = oExtRuledSurface13.Output("Support", 1)
End Sub
