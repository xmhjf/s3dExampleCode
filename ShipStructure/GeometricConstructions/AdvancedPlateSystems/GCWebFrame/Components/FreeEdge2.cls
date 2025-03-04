VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService

'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "FreeEdge2"
'End Property
'
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "FreeEdge2"
'End Property

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Curve", "Select a curve", "IJCurve or IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oPointFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS1 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0001-PointFromCS")
    oPointFromCS1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPointFromCS1.Parameter("X") = 0#
    oPointFromCS1.Parameter("Y") = 100.0#
    oPointFromCS1.Parameter("Z") = 0#
    oPointFromCS1.Evaluate

    Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0002-LineFromCS")
    oLineFromCS2.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oLineFromCS2.Parameter("LookingAxis") = 1
    oLineFromCS2.Parameter("Length") = -1
    oLineFromCS2.Parameter("CSOrientation") = 2
    oLineFromCS2.Parameter("LineJustification") = 1
    oLineFromCS2.Evaluate

    Dim oCurveAtOffset3 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveAtOffset3 = oGCFactory.CreateEntity("CurveAtOffset", pPOM, "0003-CurveAtOffset")
    oCurveAtOffset3.Inputs("Curve").Add pGeometricConstruction.Inputs("Curve").Item(1), "1"
    oCurveAtOffset3.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveAtOffset3.Inputs("TrackPoint").Add oPointFromCS1, "2"
    oCurveAtOffset3.Parameter("Offset") = pGeometricConstruction.Parameter("Offset")
    oCurveAtOffset3.Parameter("Orientation") = 1
    oCurveAtOffset3.Parameter("TrackFlag") = 1
    oCurveAtOffset3.Evaluate

    Dim oSurfByLinearExtrusion4 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion4 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0004-SurfByLinearExtrusion")
    oSurfByLinearExtrusion4.Inputs("PlanarCrossSection").Add oCurveAtOffset3, "1"
    oSurfByLinearExtrusion4.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
    oSurfByLinearExtrusion4.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", "1") = oSurfByLinearExtrusion4.Output

End Sub
