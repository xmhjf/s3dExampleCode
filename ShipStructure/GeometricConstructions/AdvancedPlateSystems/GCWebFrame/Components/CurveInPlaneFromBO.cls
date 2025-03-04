VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurveInPlaneFromBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
'Implements IJGCTypeInfo
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "CurveInPlaneFromBO"
'End Property
'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "Curve by intersection or projection of a FCBO (plate/stiffener/member system) in a plane"
'End Property
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Connectable", "Select a plane/plate/stiffener/member system", "IJPlane" + _
                                                                                                " OR IJPlateSystem AND [GCCMNSTRDefinitions.Filters,IsRootPlateSystem]" + _
                                                                                                " OR IJStiffenerSystem AND [GCCMNSTRDefinitions.Filters,IsRootStiffenerSystem]" + _
                                                                                                " OR ISPSMemberSystem", 1, 1, "")
    Call pGeometricConstructionDefinition.AddInput("Plane", "Select a plane", "IJPlane OR IJSurfaceBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJPlane]", 1, 1)
    Call pGeometricConstructionDefinition.AddControlledInput("Port")
    
    Call pGeometricConstructionDefinition.AddOutput(7, "Curve")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    pGeometricConstruction.ControlledInputs("Port").Clear

    Dim oConnectable As Object: Set oConnectable = pGeometricConstruction.Input("Connectable")
    Dim oResultingCurve As SP3DGeometricConstruction.GeometricConstruction: Set oResultingCurve = Nothing
    If TypeOf oConnectable Is IJPlateSystem Then
        Dim oFacePortExtractor01 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor01 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0001-FacePortExtractor0")
        oFacePortExtractor01.Inputs("Connectable").Add pGeometricConstruction.Inputs("Connectable").Item(1), "1"
        oFacePortExtractor01.Parameter("Offset") = 0#
        oFacePortExtractor01.Parameter("GeometrySelector") = 4
        oFacePortExtractor01.Evaluate
        pGeometricConstruction.ControlledInputs("Port").AddElements oFacePortExtractor01.ControlledInputs("Port")
        
        Dim oCurveByIntersection2 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection2 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0002-CurveByIntersection")
        oCurveByIntersection2.Inputs("Surface1").Add pGeometricConstruction.Inputs("Plane").Item(1), "1"
        oCurveByIntersection2.Inputs("Surface2").Add oFacePortExtractor01, "1"
        oCurveByIntersection2.Parameter("TrackFlag") = 1
        oCurveByIntersection2.Parameter("ConstructionSurface") = 0
        oCurveByIntersection2.Evaluate
    
        Set oResultingCurve = oCurveByIntersection2
        
    ElseIf TypeOf oConnectable Is IJStiffenerSystem Or TypeOf oConnectable Is ISPSMemberSystem Then
        Dim oAxisPortExtractor22 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor22 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0002-AxisPortExtractor2")
        oAxisPortExtractor22.Inputs("Connectable").Add pGeometricConstruction.Inputs("Connectable").Item(1), "11"
        oAxisPortExtractor22.Parameter("GeometrySelector") = 4
        oAxisPortExtractor22.Evaluate
        pGeometricConstruction.ControlledInputs("Port").AddElements oAxisPortExtractor22.ControlledInputs("Port")
        
        Dim oCSByPlane1 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByPlane1 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0001-CSByPlane")
        oCSByPlane1.Inputs("Plane").Add pGeometricConstruction.Inputs("Plane").Item(1), "3"
        oCSByPlane1.Evaluate
    
        Dim oLineFromCS3 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS3 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0003-LineFromCS")
        oLineFromCS3.Inputs("CoordinateSystem").Add oCSByPlane1, "1"
        oLineFromCS3.Parameter("LookingAxis") = 3
        oLineFromCS3.Parameter("Length") = 5#
        oLineFromCS3.Parameter("CSOrientation") = 1
        oLineFromCS3.Parameter("LineJustification") = 1
        oLineFromCS3.Evaluate
    
        Dim oCurveByProjection4 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByProjection4 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0004-CurveByProjection")
        oCurveByProjection4.Inputs("Curve").Add oAxisPortExtractor22, "1"
        oCurveByProjection4.Inputs("Surface").Add pGeometricConstruction.Inputs("Plane").Item(1), "2"
        oCurveByProjection4.Inputs("Line").Add oLineFromCS3, "1"
        oCurveByProjection4.Parameter("TrackFlag") = 1
        oCurveByProjection4.Parameter("ConstructionSurface") = 0
        oCurveByProjection4.Evaluate
        
        Set oResultingCurve = oCurveByProjection4
    Else
        Dim oCurveByIntersection5 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection5 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0005-CurveByIntersection")
        oCurveByIntersection5.Inputs("Surface1").Add pGeometricConstruction.Inputs("Plane").Item(1), "1"
        oCurveByIntersection5.Inputs("Surface2").Add pGeometricConstruction.Inputs("Connectable").Item(1), "1"
        oCurveByIntersection5.Parameter("TrackFlag") = 1
        oCurveByIntersection5.Parameter("ConstructionSurface") = 0
        oCurveByIntersection5.Evaluate
        
        Set oResultingCurve = oCurveByIntersection5
    End If
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Curve", 1) = oResultingCurve.Output
End Sub
