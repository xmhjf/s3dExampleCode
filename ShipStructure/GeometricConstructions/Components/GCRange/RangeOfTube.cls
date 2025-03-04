VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeOfTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("PrimaryMember", "Select a MemberPart", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddInput("Tube", "Select a Tube", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a Tube", "IJStructCoordinateSystem", 0, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_5_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_6_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_7_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_8_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "Lines")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypePoint3d, "Points")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryMember").Item(1), "1"
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port_7_1").Clear
    pGeometricConstruction.ControlledInputs("Port_7_1").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Tube").Item(1), "1"
    oLineAxisPortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("Port_8_1").Clear
    pGeometricConstruction.ControlledInputs("Port_8_1").Add oLineAxisPortExtractor2.ControlledInputs("Port").Item("1")
   
    Dim oCS As Object
    If pGeometricConstruction.Inputs("CoordinateSystem").Count = 0 Then
        Dim oCSByLines3a As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines3a = oGCFactory.CreateEntity("CSByLines", pPOM, "003a-CSByLines")
        oCSByLines3a.Inputs("AxisLine1").Add oLineAxisPortExtractor1
        oCSByLines3a.Inputs("AxisLine2").Add oLineAxisPortExtractor2
        oCSByLines3a.Parameter("AxesRoles") = 1
        oCSByLines3a.Parameter("CSOrientation") = 1
        oCSByLines3a.Parameter("TrackFlag") = 1
        oCSByLines3a.Evaluate
    
        Dim oLineFromCS3b As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS3b = oGCFactory.CreateEntity("LineFromCS", pPOM, "003b-LineFromCS")
        oLineFromCS3b.Inputs("CoordinateSystem").Add oCSByLines3a
        oLineFromCS3b.Parameter("LookingAxis") = 2
        oLineFromCS3b.Parameter("Length") = 1#
        oLineFromCS3b.Parameter("CSOrientation") = 1
        oLineFromCS3b.Parameter("LineJustification") = 1
        oLineFromCS3b.Evaluate
        
        Dim oCSByLines3c As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines3c = oGCFactory.CreateEntity("CSByLines", pPOM, "003c-CSByLines")
        oCSByLines3c.Inputs("AxisLine1").Add oLineFromCS3b
        oCSByLines3c.Inputs("AxisLine2").Add oLineAxisPortExtractor1
        oCSByLines3c.Parameter("AxesRoles") = 3
        oCSByLines3c.Parameter("CSOrientation") = 1
        oCSByLines3c.Parameter("TrackFlag") = 1
        oCSByLines3c.Evaluate
        
        Set oCS = oCSByLines3c
    Else
        Set oCS = pGeometricConstruction.Input("CoordinateSystem")
    End If

    Dim oFacePortExtractor4 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor4 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "004-FacePortExtractor")
    oFacePortExtractor4.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryMember").Item(1), "1"
    oFacePortExtractor4.Inputs("CoordinateSystem").Add oCS
    oFacePortExtractor4.Parameter("GeometrySelector") = 4
    oFacePortExtractor4.Parameter("FacesContext") = 2
    oFacePortExtractor4.Parameter("LookingAxis") = 3
    oFacePortExtractor4.Parameter("IntersectingPlane") = 0
    oFacePortExtractor4.Parameter("SurfaceType") = 1
    oFacePortExtractor4.Parameter("TrackFlag") = 2
    oFacePortExtractor4.Parameter("Offset") = 0#
    oFacePortExtractor4.Evaluate
    pGeometricConstruction.ControlledInputs("Port_5_1").Clear
    pGeometricConstruction.ControlledInputs("Port_5_1").Add oFacePortExtractor4.ControlledInputs("Port").Item("1")
    
    Dim oPointByCurveAndSurface5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface5 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "005-PointByCurveAndSurface")
    oPointByCurveAndSurface5.Inputs("Curve1").Add oLineAxisPortExtractor2
    oPointByCurveAndSurface5.Inputs("Surface2").Add oFacePortExtractor4
    oPointByCurveAndSurface5.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface5.Evaluate

    Dim oCSByCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS6 = oGCFactory.CreateEntity("CSByCS", pPOM, "006-CSByCS")
    oCSByCS6.Inputs("CoordinateSystem").Add oCS
    oCSByCS6.Inputs("Origin").Add oPointByCurveAndSurface5
    oCSByCS6.Parameter("AxesRoles") = 1
    oCSByCS6.Parameter("CSOrientation") = 1
    oCSByCS6.Parameter("TrackFlag") = 1
    oCSByCS6.Evaluate

    Dim oLineFromCS7x As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7x = oGCFactory.CreateEntity("LineFromCS", pPOM, "007x-LineFromCS")
    oLineFromCS7x.Inputs("CoordinateSystem").Add oCSByCS6
    oLineFromCS7x.Parameter("LookingAxis") = 1
    oLineFromCS7x.Parameter("Length") = -100#
    oLineFromCS7x.Parameter("CSOrientation") = 1
    oLineFromCS7x.Parameter("LineJustification") = 1
    oLineFromCS7x.Evaluate
    
    Dim oLineFromCS7y As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7y = oGCFactory.CreateEntity("LineFromCS", pPOM, "007y-LineFromCS")
    oLineFromCS7y.Inputs("CoordinateSystem").Add oCSByCS6
    oLineFromCS7y.Parameter("LookingAxis") = 2
    oLineFromCS7y.Parameter("Length") = -100#
    oLineFromCS7y.Parameter("CSOrientation") = 1
    oLineFromCS7y.Parameter("LineJustification") = 1
    oLineFromCS7y.Evaluate

    Dim oPlaneFromCS8z As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS8z = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "008z-PlaneFromCS")
    oPlaneFromCS8z.Inputs("CoordinateSystem").Add oCSByCS6
    oPlaneFromCS8z.Parameter("LookingAxis") = 3
    oPlaneFromCS8z.Parameter("Range") = 100#
    oPlaneFromCS8z.Evaluate

    Dim oPointFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "009-PointFromCS")
    oPointFromCS9.Inputs("CoordinateSystem").Add oCSByCS6
    oPointFromCS9.Parameter("X") = -10#
    oPointFromCS9.Parameter("Y") = -10#
    oPointFromCS9.Parameter("Z") = 0#
    oPointFromCS9.Evaluate
    
    Dim oLineAxisPortExtractor11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor11 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "011-LineAxisPortExtractor")
    oLineAxisPortExtractor11.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Tube").Item(1), "1"
    oLineAxisPortExtractor11.Evaluate
    pGeometricConstruction.ControlledInputs("Port_8_1").Clear
    pGeometricConstruction.ControlledInputs("Port_8_1").Add oLineAxisPortExtractor11.ControlledInputs("Port").Item("1")
    
    Dim oPointAtCurveMiddle12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "012-PointAtCurveMiddle")
    oPointAtCurveMiddle12.Inputs("Curve").Add oLineAxisPortExtractor11
    oPointAtCurveMiddle12.Evaluate

    Dim oCSFromMember13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember13 = oGCFactory.CreateEntity("CSFromMember", pPOM, "013-CSFromMember")
    oCSFromMember13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Tube").Item(1), "1"
    oCSFromMember13.Inputs("Point").Add oPointAtCurveMiddle12
    oCSFromMember13.Evaluate

    Dim oFacePortExtractor14 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor14 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "014-FacePortExtractor")
    oFacePortExtractor14.Inputs("Connectable").Add pGeometricConstruction.Inputs("Tube").Item(1), "1"
    oFacePortExtractor14.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor14.Parameter("GeometrySelector") = 4
    oFacePortExtractor14.Parameter("FacesContext") = 2 '4
    oFacePortExtractor14.Parameter("LookingAxis") = 2
    oFacePortExtractor14.Parameter("IntersectingPlane") = 0
    oFacePortExtractor14.Parameter("SurfaceType") = 1
    oFacePortExtractor14.Parameter("TrackFlag") = 1
    oFacePortExtractor14.Parameter("Offset") = 0#
    oFacePortExtractor14.Evaluate
    pGeometricConstruction.ControlledInputs("Port_6_1").Clear
    pGeometricConstruction.ControlledInputs("Port_6_1").Add oFacePortExtractor14.ControlledInputs("Port").Item("1")
    
    Dim oPlaneFromCS16x As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS16x = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "016x-PlaneFromCS")
    oPlaneFromCS16x.Inputs("CoordinateSystem").Add oCSByCS6
    oPlaneFromCS16x.Parameter("LookingAxis") = 1
    oPlaneFromCS16x.Parameter("Range") = 100#
    oPlaneFromCS16x.Evaluate

    Dim oPlaneFromCS16y As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS16y = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "016y-PlaneFromCS")
    oPlaneFromCS16y.Inputs("CoordinateSystem").Add oCSByCS6
    oPlaneFromCS16y.Parameter("LookingAxis") = 2
    oPlaneFromCS16y.Parameter("Range") = 100#
    oPlaneFromCS16y.Evaluate

    Dim oPointBySurfaces21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointBySurfaces21 = oGCFactory.CreateEntity("PointBySurfaces", pPOM, "021-PointBySurfaces")
    oPointBySurfaces21.Inputs("Surface1").Add oPlaneFromCS16y
    oPointBySurfaces21.Inputs("Surface2").Add oFacePortExtractor14
    oPointBySurfaces21.Inputs("Surface3").Add oPlaneFromCS8z
    oPointBySurfaces21.Inputs("TrackPoint").Add oPointFromCS9
    oPointBySurfaces21.Parameter("TrackFlag") = 1
    oPointBySurfaces21.Evaluate
    
    Dim oPointBySurfaces22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointBySurfaces22 = oGCFactory.CreateEntity("PointBySurfaces", pPOM, "022-PointBySurfaces")
    oPointBySurfaces22.Inputs("Surface1").Add oPlaneFromCS16y
    oPointBySurfaces22.Inputs("Surface2").Add oFacePortExtractor14
    oPointBySurfaces22.Inputs("Surface3").Add oPlaneFromCS8z
    oPointBySurfaces22.Inputs("TrackPoint").Add oPointFromCS9
    oPointBySurfaces22.Parameter("TrackFlag") = 2
    oPointBySurfaces22.Evaluate

    Dim oPointBySurfaces23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointBySurfaces23 = oGCFactory.CreateEntity("PointBySurfaces", pPOM, "023-PointBySurfaces")
    oPointBySurfaces23.Inputs("Surface1").Add oPlaneFromCS16x
    oPointBySurfaces23.Inputs("Surface2").Add oFacePortExtractor14
    oPointBySurfaces23.Inputs("Surface3").Add oPlaneFromCS8z
    oPointBySurfaces23.Inputs("TrackPoint").Add oPointFromCS9
    oPointBySurfaces23.Parameter("TrackFlag") = 1
    oPointBySurfaces23.Evaluate

    Dim oPointBySurfaces24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointBySurfaces24 = oGCFactory.CreateEntity("PointBySurfaces", pPOM, "024-PointBySurfaces")
    oPointBySurfaces24.Inputs("Surface1").Add oPlaneFromCS16x
    oPointBySurfaces24.Inputs("Surface2").Add oFacePortExtractor14
    oPointBySurfaces24.Inputs("Surface3").Add oPlaneFromCS8z
    oPointBySurfaces24.Inputs("TrackPoint").Add oPointFromCS9
    oPointBySurfaces24.Parameter("TrackFlag") = 2
    oPointBySurfaces24.Evaluate
                           
    Dim oLineAtAngle31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle31 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "031-LineAtAngle")
    oLineAtAngle31.Inputs("Line").Add oLineFromCS7y
    oLineAtAngle31.Inputs("Point").Add oPointBySurfaces21
    oLineAtAngle31.Parameter("Angle") = 0#
    oLineAtAngle31.Parameter("TrackFlag") = 1
    oLineAtAngle31.Evaluate

    Dim oLineAtAngle32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle32 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "032-LineAtAngle")
    oLineAtAngle32.Inputs("Line").Add oLineFromCS7y
    oLineAtAngle32.Inputs("Point").Add oPointBySurfaces22
    oLineAtAngle32.Parameter("Angle") = 0#
    oLineAtAngle32.Parameter("TrackFlag") = 1
    oLineAtAngle32.Evaluate

    Dim oLineAtAngle33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle33 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "033-LineAtAngle")
    oLineAtAngle33.Inputs("Line").Add oLineFromCS7x
    oLineAtAngle33.Inputs("Point").Add oPointBySurfaces23
    oLineAtAngle33.Parameter("Angle") = 0#
    oLineAtAngle33.Parameter("TrackFlag") = 1
    oLineAtAngle33.Evaluate

    Dim oLineAtAngle34 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle34 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "034-LineAtAngle")
    oLineAtAngle34.Inputs("Line").Add oLineFromCS7x
    oLineAtAngle34.Inputs("Point").Add oPointBySurfaces24
    oLineAtAngle34.Parameter("Angle") = 0#
    oLineAtAngle34.Parameter("TrackFlag") = 1
    oLineAtAngle34.Evaluate

    Dim oLineExtend35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend35 = oGCFactory.CreateEntity("LineExtend", pPOM, "035-LineExtend")
    oLineExtend35.Inputs("Line").Add oLineAtAngle31
    oLineExtend35.Inputs("Curve1").Add oLineAtAngle34
    oLineExtend35.Inputs("Curve2").Add oLineAtAngle33
    oLineExtend35.Parameter("Context1") = 1
    oLineExtend35.Parameter("Context2") = 3
    oLineExtend35.Evaluate

    Dim oLineExtend36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend36 = oGCFactory.CreateEntity("LineExtend", pPOM, "036-LineExtend")
    oLineExtend36.Inputs("Line").Add oLineAtAngle32
    oLineExtend36.Inputs("Curve1").Add oLineAtAngle33
    oLineExtend36.Inputs("Curve2").Add oLineAtAngle34
    oLineExtend36.Parameter("Context1") = 1
    oLineExtend36.Parameter("Context2") = 3
    oLineExtend36.Evaluate
    
    Dim oLineExtend37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend37 = oGCFactory.CreateEntity("LineExtend", pPOM, "037-LineExtend")
    oLineExtend37.Inputs("Line").Add oLineAtAngle33
    oLineExtend37.Inputs("Curve1").Add oLineExtend35
    oLineExtend37.Inputs("Curve2").Add oLineAtAngle32
    oLineExtend37.Parameter("Context1") = 1
    oLineExtend37.Parameter("Context2") = 3
    oLineExtend37.Evaluate
    
    Dim oLineExtend38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend38 = oGCFactory.CreateEntity("LineExtend", pPOM, "038-LineExtend")
    oLineExtend38.Inputs("Line").Add oLineAtAngle34
    oLineExtend38.Inputs("Curve1").Add oLineAtAngle31
    oLineExtend38.Inputs("Curve2").Add oLineAtAngle32
    oLineExtend38.Parameter("Context1") = 1
    oLineExtend38.Parameter("Context2") = 3
    oLineExtend38.Evaluate
    
    Dim oPointByCurves41 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves41 = oGCFactory.CreateEntity("PointByCurves", pPOM, "041-PointByCurves")
    oPointByCurves41.Inputs("Curve1").Add oLineExtend35
    oPointByCurves41.Inputs("Curve2").Add oLineExtend37
    oPointByCurves41.Evaluate

    Dim oPointByCurves42 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves42 = oGCFactory.CreateEntity("PointByCurves", pPOM, "042-PointByCurves")
    oPointByCurves42.Inputs("Curve1").Add oLineExtend36
    oPointByCurves42.Inputs("Curve2").Add oLineExtend38
    oPointByCurves42.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("CoordinateSystem", 1) = oCSByCS6.Output
    oGCMacro.Output("Lines", "XHigh") = oLineExtend35.Output
    oGCMacro.Output("Lines", "XLow") = oLineExtend36.Output
    oGCMacro.Output("Lines", "YLow") = oLineExtend37.Output
    oGCMacro.Output("Lines", "YHigh") = oLineExtend38.Output
    oGCMacro.Output("Points", "Low") = oPointByCurves41.Output
    oGCMacro.Output("Points", "High") = oPointByCurves42.Output
End Sub
