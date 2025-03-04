VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEdge1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements IJGCTypeInfo
Implements IJGeometricConstructionDefinitionService

'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "FreeEdge1"
'End Property
'
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "FreeEdge1"
'End Property
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point1", "Select a lower point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point2", "Select an upper point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem1", "Select a lower coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem2", "Select an upper coordinate system", "IJDCoordinateSystem", 1, 1)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS1 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0001-LineFromCS")
    oLineFromCS1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1), "1"
    oLineFromCS1.Parameter("LookingAxis") = 1
    oLineFromCS1.Parameter("Length") = -1
    oLineFromCS1.Parameter("CSOrientation") = 2
    oLineFromCS1.Parameter("LineJustification") = 1
    oLineFromCS1.Evaluate

    Dim oLineByPoints2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints2 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0002-LineByPoints")
    oLineByPoints2.Inputs("StartPoint").Add pGeometricConstruction.Inputs("Point2").Item(1), "1"
    oLineByPoints2.Inputs("EndPoint").Add pGeometricConstruction.Inputs("Point1").Item(1), "1"
    oLineByPoints2.Evaluate

    Dim oLineFromCS3 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS3 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0003-LineFromCS")
    oLineFromCS3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem2").Item(1), "1"
    oLineFromCS3.Parameter("LookingAxis") = 3
    oLineFromCS3.Parameter("Length") = 2#
    oLineFromCS3.Parameter("CSOrientation") = 2
    oLineFromCS3.Parameter("LineJustification") = 1
    oLineFromCS3.Evaluate

    Dim oLineFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS4 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0004-LineFromCS")
    oLineFromCS4.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem1").Item(1), "2"
    oLineFromCS4.Parameter("LookingAxis") = 3
    oLineFromCS4.Parameter("Length") = 2#
    oLineFromCS4.Parameter("CSOrientation") = 2
    oLineFromCS4.Parameter("LineJustification") = 1
    oLineFromCS4.Evaluate

    Dim oCpxStringByCurves5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves5 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0005a-CpxStringByCurves")
    oCpxStringByCurves5.Inputs("Curves").Add oLineFromCS3, "1"
    oCpxStringByCurves5.Inputs("Curves").Add oLineByPoints2, "2"
    oCpxStringByCurves5.Inputs("Curves").Add oLineFromCS4, "3"
    oCpxStringByCurves5.Evaluate
    
    Dim oCurveByProjection5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection5 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0005b-CurveByProjection")
    oCurveByProjection5.Inputs("Curve").Add oCpxStringByCurves5, "1"
    oCurveByProjection5.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByProjection5.Inputs("Line").Add oLineFromCS1, "1"
    oCurveByProjection5.Evaluate

    Dim oSurfByLinearExtrusion6 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion6 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0006-SurfByLinearExtrusion")
    oSurfByLinearExtrusion6.Inputs("PlanarCrossSection").Add oCurveByProjection5, "1"
    oSurfByLinearExtrusion6.Inputs("ExtrusionLine").Add oLineFromCS1, "1"
    oSurfByLinearExtrusion6.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", "1") = oSurfByLinearExtrusion6.Output
End Sub
