VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PointAtMinDistEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
'Implements IJGCTypeInfo
'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "PointAtMinDistEx"
'End Property
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "PointAtMinDistEx"
'End Property

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select main reference for sketching plane (plate/profile system, plane)", "IJPlane", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Connectable1", "Select a middle root plate/profile system or surface", "IJStiffenerSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootStiffenerSystem] OR IJPlateSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootPlateSystem] OR [GCSHPSTRDefinitions.ShipFilters,IsPlaneNotPlate] OR [GCSHPSTRDefinitions.ShipFilters,IsSurfaceNotPlate]", 1, 1, "IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    Call pGeometricConstructionDefinition.AddInput("Connectable2", "Select a upper or lower root plate/profile system or surface", "IJStiffenerSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootStiffenerSystem] OR IJPlateSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootPlateSystem] OR [GCSHPSTRDefinitions.ShipFilters,IsPlaneNotPlate] OR [GCSHPSTRDefinitions.ShipFilters,IsSurfaceNotPlate]", 1, 1, "IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("TracePoint", "Select trace point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("SidePoint", "Select a side  point", "IJPoint", 1, 1)
    
    Call pGeometricConstructionDefinition.AddControlledInput("Port1")
    Call GeometricConstructionDefinition_CopyControlledInputsEx(pGeometricConstructionDefinition, "GCSHPSTRDefinitions.PointAtMinimumDistanceDef", "")
    
    Call pGeometricConstructionDefinition.AddOutput(1, "Point")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundary")
    
    Call pGeometricConstructionDefinition.AddParameter("BoundaryOffset", "BoundaryOffset", 8, 1, 59, 0, 0, 0.25)
    Call pGeometricConstructionDefinition.AddParameter("PointEvaluationMethod", "Method", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("PointEvaluationMethod", "At girth then along normal", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("PointEvaluationMethod", "At intersection of offset curves", 2)
    Call pGeometricConstructionDefinition.AddParameter("GirthDistance", "Girth", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("NormalDistance", "Normal", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("OffsetDistance1", "Offset 1", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("OffsetDistance2", "Offset 2", 8, 1, 59, 0, 0, 1#)
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim dBoundaryOffset As Double: Let dBoundaryOffset = pGeometricConstruction.Parameter("BoundaryOffset")
    
    Dim oCurveInPlaneFromBO1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveInPlaneFromBO1 = oGCFactory.CreateEntity("CurveInPlaneFromBO", pPOM, "0001-CurveInPlaneFromBO")
    oCurveInPlaneFromBO1.Inputs("Connectable").Add pGeometricConstruction.Inputs("Connectable2").Item(1), "1"
    oCurveInPlaneFromBO1.Inputs("Plane").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveInPlaneFromBO1.Evaluate

    Dim oPointFromCS2a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS2a = oGCFactory.CreateEntity("PointFromCS", pPOM, "0002a-PointFromCS")
    oPointFromCS2a.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPointFromCS2a.Inputs("Point").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oPointFromCS2a.Parameter("X") = 0#
    oPointFromCS2a.Parameter("Y") = 10#
    oPointFromCS2a.Parameter("Z") = 0#
    oPointFromCS2a.Evaluate
    
    Dim oVectorTangentToCurve2b As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve2b = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0002b-VectorTangentToCurve")
    oVectorTangentToCurve2b.Inputs("Curve").Add oCurveInPlaneFromBO1.Output("Curve"), "1"
    oVectorTangentToCurve2b.Inputs("Point").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oVectorTangentToCurve2b.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oVectorTangentToCurve2b.Inputs("TrackPoint").Add oPointFromCS2a, "1"
    oVectorTangentToCurve2b.Parameter("Range") = 10#
    oVectorTangentToCurve2b.Parameter("Orientation") = 1
    oVectorTangentToCurve2b.Parameter("TrackFlag") = 1
    oVectorTangentToCurve2b.Evaluate
    
    Dim oCSByCS3  As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS3 = oGCFactory.CreateEntity("CSByCS", pPOM, "0003-CSByCS")
    oCSByCS3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oCSByCS3.Inputs("Origin").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oCSByCS3.Evaluate

    Dim oLineFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS4 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0004-LineFromCS")
    oLineFromCS4.Inputs("CoordinateSystem").Add oCSByCS3, "1"
    oLineFromCS4.Parameter("LookingAxis") = 1
    oLineFromCS4.Parameter("Length") = 1
    oLineFromCS4.Parameter("CSOrientation") = 1#
    oLineFromCS4.Evaluate

    Dim oCSByLines5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines5 = oGCFactory.CreateEntity("CSByLines", pPOM, "0005-CSByLines")
    oCSByLines5.Inputs("AxisLine1").Add oVectorTangentToCurve2b, "1"
    oCSByLines5.Inputs("AxisLine2").Add oLineFromCS4, "1"
    oCSByLines5.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("SidePoint").Item(1), "1"
    oCSByLines5.Parameter("AxesRoles") = 4
    oCSByLines5.Parameter("CSOrientation") = 1
    oCSByLines5.Evaluate
    
    Dim oPointFromCS5a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS5a = oGCFactory.CreateEntity("PointFromCS", pPOM, "0005a-PointFromCS")
    oPointFromCS5a.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oPointFromCS5a.Inputs("Point").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oPointFromCS5a.Parameter("X") = 0#
    oPointFromCS5a.Parameter("Y") = 0.1
    oPointFromCS5a.Parameter("Z") = 1#
    oPointFromCS5a.Evaluate
    
    Dim oRootSystemOrSurface1 As Object: Set oRootSystemOrSurface1 = pGeometricConstruction.Inputs("Connectable2").Item(1)
    Dim oBoundary As Object: Set oBoundary = Nothing
    If dBoundaryOffset <> 0# Then
        If TypeOf oRootSystemOrSurface1 Is IJPlateSystem _
        Or TypeOf oRootSystemOrSurface1 Is IJStiffenerSystem Then
            Dim oFacePortGenerator6 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortGenerator6 = oGCFactory.CreateEntity("FacePortGenerator", pPOM, "0006-FacePortGenerator")
            oFacePortGenerator6.Inputs("RootPlateSystem").Add oRootSystemOrSurface1, "1"
            oFacePortGenerator6.Inputs("ReferencePoint").Add oPointFromCS5a, "1"
            oFacePortGenerator6.Parameter("OutputRadius") = 0#
            oFacePortGenerator6.Parameter("GeometrySelector") = 4#
            oFacePortGenerator6.Evaluate
            Set oRootSystemOrSurface1 = oFacePortGenerator6
        End If
        
        Dim oSurfAtOffset6 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfAtOffset6 = oGCFactory.CreateEntity("SurfAtOffset", pPOM, "0006-SurfAtOffset")
        oSurfAtOffset6.Inputs("Surface").Add oRootSystemOrSurface1, "1"
        oSurfAtOffset6.Inputs("TrackPoint").Add oPointFromCS5a, "1"
        oSurfAtOffset6.Parameter("Offset") = dBoundaryOffset
        oSurfAtOffset6.Parameter("TrackFlag") = 1
        oSurfAtOffset6.Evaluate
        Set oRootSystemOrSurface1 = oSurfAtOffset6
        Set oBoundary = oRootSystemOrSurface1
    End If

    Dim oPointAtMinimumDistance8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtMinimumDistance8 = oGCFactory.CreateEntity("PointAtMinimumDistance", pPOM, "0008-PointAtMinimumDistance")
    oPointAtMinimumDistance8.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oPointAtMinimumDistance8.Inputs("RootPlateSystemOrSurface1").Add oRootSystemOrSurface1, "1"
    oPointAtMinimumDistance8.Inputs("RootPlateSystemOrSurface2").Add pGeometricConstruction.Inputs("Connectable1").Item(1), "1"
    oPointAtMinimumDistance8.Inputs("SketchingPlane").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "2"
    oPointAtMinimumDistance8.Parameter("PointEvaluationMethod") = pGeometricConstruction.Parameter("PointEvaluationMethod")
    oPointAtMinimumDistance8.Parameter("EvaluationMethodGirthDistance") = pGeometricConstruction.Parameter("GirthDistance")
    oPointAtMinimumDistance8.Parameter("EvaluationMethodNormalDistance") = pGeometricConstruction.Parameter("NormalDistance")
    oPointAtMinimumDistance8.Parameter("EvaluationMethodOffsetDistance1") = pGeometricConstruction.Parameter("OffsetDistance1")
    oPointAtMinimumDistance8.Parameter("EvaluationMethodOffsetDistance2") = pGeometricConstruction.Parameter("OffsetDistance2")
    oPointAtMinimumDistance8.Parameter("2DCreated") = 0
    If WebFrameLocation = 2 Then  'outside
        oPointAtMinimumDistance8.Parameter("TrackFlagForIntersection") = 3 ' +y
    Else
        oPointAtMinimumDistance8.Parameter("TrackFlagForIntersection") = 4 ' -y
    End If
    
    oPointAtMinimumDistance8.Parameter("TrackFlagForGirthOrientation") = 3
    oPointAtMinimumDistance8.Parameter("TrackFlagForNormalOrientation") = 5
    oPointAtMinimumDistance8.Parameter("TrackFlagForOffsetOrientation1") = 3
    oPointAtMinimumDistance8.Parameter("TrackFlagForOffsetOrientation2") = 5
    oPointAtMinimumDistance8.Evaluate
    Call GeometricConstruction_PropagateControlledInputs(pGeometricConstruction, oPointAtMinimumDistance8, "")
    
    Dim oPointFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0009PointFromCS")
    oPointFromCS9.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oPointFromCS9.Inputs("Point").Add oPointAtMinimumDistance8, "1"
    oPointFromCS9.Parameter("X") = 0#
    oPointFromCS9.Parameter("Y") = 0#
    oPointFromCS9.Parameter("Z") = 1#
    oPointFromCS9.Evaluate

    Dim oCSByPoints10 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints10 = oGCFactory.CreateEntity("CSByPoints", pPOM, "0010-CSByPoints")
    oCSByPoints10.Inputs("OriginPoint").Add oPointAtMinimumDistance8, "1"
    oCSByPoints10.Inputs("AxisPoint1").Add oPointFromCS9, "1"
    oCSByPoints10.Inputs("AxisPoint2").Add pGeometricConstruction.Inputs("TracePoint")(1), "1"
    oCSByPoints10.Parameter("AxesRoles") = 5
    oCSByPoints10.Parameter("CSOrientation") = 1
    oCSByPoints10.Evaluate
    
    Dim oLineFromCS12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS12 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0012-LineFromCS")
    oLineFromCS12.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oLineFromCS12.Parameter("LookingAxis") = 1
    oLineFromCS12.Parameter("Length") = -0.5
    oLineFromCS12.Parameter("CSOrientation") = 1
    oLineFromCS12.Parameter("LineJustification") = 1
    oLineFromCS12.Evaluate
    
    If oBoundary Is Nothing Then
        Dim oSurfByNonLinearExtrusion13 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByNonLinearExtrusion13 = oGCFactory.CreateEntity("SurfByNonLinearExtrusion", pPOM, "0013-SurfByNonLinearExtrusion")
        oSurfByNonLinearExtrusion13.Inputs("PlanarCrossSection").Add oLineFromCS12, "1"
        oSurfByNonLinearExtrusion13.Inputs("PlanarExtrusionPath").Add oCurveInPlaneFromBO1.Output("Curve"), "1"
        oSurfByNonLinearExtrusion13.Evaluate
        Set oBoundary = oSurfByNonLinearExtrusion13
    End If
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Point", 1) = oPointAtMinimumDistance8.Output
    oGCMacro.Output("CoordinateSystem", 1) = oCSByPoints10.Output
    oGCMacro.Output("Boundary", 1) = oBoundary
End Sub
Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, bIsOnRibbonBar As Boolean)
    Select Case sName
        Case "PointEvaluationMethod": bIsOnRibbonBar = True
        Case "GirthDistance", "NormalDistance": bIsOnRibbonBar = pGC.Parameter("PointEvaluationMethod") = 1
        Case "OffsetDistance1", "OffsetDistance2": bIsOnRibbonBar = pGC.Parameter("PointEvaluationMethod") = 2
    End Select
End Sub
