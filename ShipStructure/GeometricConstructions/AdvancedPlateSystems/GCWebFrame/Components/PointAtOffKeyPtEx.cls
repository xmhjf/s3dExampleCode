VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PointAtOffKeyPtEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
'Implements IJGCTypeInfo
'Private Property Get IJGCTypeInfo_Description() As String
'    IJGCTypeInfo_Description = "PointAtOffKeyPtEx"
'End Property
'Private Property Get IJGCTypeInfo_Name() As String
'    IJGCTypeInfo_Name = "PointAtOffKeyPtEx"
'End Property
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1, "IJSurface IJSurfaceBody")
    Call pGeometricConstructionDefinition.AddInput("Connectable", "Select an upper profile system or member system", "IJStiffenerSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootStiffenerSystem] OR ISPSMemberSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRolledMemberSystem]", 1, 1, "IJGeometry IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("TracePoint", "Select trace point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("SidePoint", "Select side point", "IJPoint", 1, 1)
    
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1")
    
    Call pGeometricConstructionDefinition.AddOutput(1, "Point")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundary")
    
    Call pGeometricConstructionDefinition.AddParameter("Reference", "Reference", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Reference", "Web-end", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Reference", "Flange-end", 2)
    
    Call pGeometricConstructionDefinition.AddParameter("VWeb", "Along", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("UWeb", "Normal", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("UFlange", "Along", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("VFlange", "Normal", 8, 1, 59, 0, 0, 0#)
    
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oRootPlateSystem As Object
    If True Then
        Dim pDesignChild As IJDesignChild: Set pDesignChild = pGeometricConstruction.Inputs("Connectable").Item(1)
        Set oRootPlateSystem = pDesignChild.GetParent
    End If
    
    Dim oTubeFromDetailedProfile1 As SP3DGeometricConstruction.GeometricConstruction
    Set oTubeFromDetailedProfile1 = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "0001-TubeFromDetailedProfile")
    oTubeFromDetailedProfile1.Inputs("RootProfileSystem").Add pGeometricConstruction.Inputs("Connectable").Item(1), "1"
    oTubeFromDetailedProfile1.Inputs("ReferencePoint").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oTubeFromDetailedProfile1.Parameter("Length") = 2#
    oTubeFromDetailedProfile1.Evaluate
    pGeometricConstruction.ControlledInputs("LeafPlateSystem").Clear
    pGeometricConstruction.ControlledInputs("LeafPlateSystem").Add oTubeFromDetailedProfile1.ControlledInputs("LeafPlateSystem").Item(1)
    
    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0002-FacePortExtractor0")
    oFacePortExtractor02.Inputs("Connectable").Add oRootPlateSystem, "1"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 4
    oFacePortExtractor02.Evaluate
    pGeometricConstruction.ControlledInputs("Port1").Clear
    pGeometricConstruction.ControlledInputs("Port1").Add oFacePortExtractor02.ControlledInputs("Port").Item("1")
    
    Dim oPointFromCS3a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS3a = oGCFactory.CreateEntity("PointFromCS", pPOM, "0003a-PointFromCS")
    oPointFromCS3a.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPointFromCS3a.Inputs("Point").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oPointFromCS3a.Parameter("X") = 0#
    oPointFromCS3a.Parameter("Y") = 1#
    oPointFromCS3a.Parameter("Z") = 0#
    oPointFromCS3a.Evaluate

    Dim oVectorNormalToSurface3 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface3 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "0003-VectorNormalToSurface")
    oVectorNormalToSurface3.Inputs("Surface").Add oFacePortExtractor02, "2"
    oVectorNormalToSurface3.Inputs("Point").Add oPointFromCS3a, "1"
    oVectorNormalToSurface3.Parameter("Range") = 10#
    oVectorNormalToSurface3.Parameter("Orientation") = 1
    oVectorNormalToSurface3.Parameter("TrackFlag") = 1
    oVectorNormalToSurface3.Evaluate

    Dim oPointAtCurveEnd4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd4 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0004-PointAtCurveEnd")
    oPointAtCurveEnd4.Inputs("Curve").Add oVectorNormalToSurface3, "2"
    oPointAtCurveEnd4.Evaluate

    Dim oCSByPoints5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints5 = oGCFactory.CreateEntity("CSByPoints", pPOM, "0005-CSByPoints")
    oCSByPoints5.Inputs("OriginPoint").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oCSByPoints5.Inputs("AxisPoint1").Add oPointAtCurveEnd4, "1"
    oCSByPoints5.Inputs("AxisPoint2").Add pGeometricConstruction.Inputs("SidePoint").Item(1), "1"
    oCSByPoints5.Parameter("AxesRoles") = 5
    oCSByPoints5.Parameter("CSOrientation") = 1
    oCSByPoints5.Evaluate

    Dim oSurfaceBodyExtractor6 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfaceBodyExtractor6 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "0006-SurfaceBodyExtractor")
    oSurfaceBodyExtractor6.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile1, "1"
    oSurfaceBodyExtractor6.Inputs("CoordinateSystem").Add oCSByPoints5, "1"
    oSurfaceBodyExtractor6.Parameter("FacesContext") = 2
    oSurfaceBodyExtractor6.Parameter("LookingAxis") = 2
    oSurfaceBodyExtractor6.Parameter("IntersectingPlane") = 2
    oSurfaceBodyExtractor6.Parameter("SurfaceType") = 1
    oSurfaceBodyExtractor6.Parameter("TrackFlag") = 2
    oSurfaceBodyExtractor6.Parameter("Offset") = 0#
    oSurfaceBodyExtractor6.Evaluate

    Dim oSurfaceBodyExtractor7 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfaceBodyExtractor7 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "0007-SurfaceBodyExtractor")
    oSurfaceBodyExtractor7.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile1, "1"
    oSurfaceBodyExtractor7.Inputs("CoordinateSystem").Add oCSByPoints5, "1"
    oSurfaceBodyExtractor7.Inputs("Surface").Add oSurfaceBodyExtractor6, "1"
    oSurfaceBodyExtractor7.Parameter("FacesContext") = 4
    oSurfaceBodyExtractor7.Parameter("LookingAxis") = 3
    oSurfaceBodyExtractor7.Parameter("IntersectingPlane") = 2
    oSurfaceBodyExtractor7.Parameter("SurfaceType") = 1
    oSurfaceBodyExtractor7.Parameter("TrackFlag") = 2
    oSurfaceBodyExtractor7.Parameter("Offset") = 0#
    oSurfaceBodyExtractor7.Evaluate

    Dim oCurveByIntersection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection8 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0008-CurveByIntersection")
    oCurveByIntersection8.Inputs("Surface1").Add oSurfaceBodyExtractor6, "1"
    oCurveByIntersection8.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByIntersection8.Parameter("TrackFlag") = 1
    oCurveByIntersection8.Parameter("ConstructionSurface") = 0
    oCurveByIntersection8.Evaluate

    Dim oCurveByIntersection9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection9 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0009-CurveByIntersection")
    oCurveByIntersection9.Inputs("Surface1").Add oSurfaceBodyExtractor7, "2"
    oCurveByIntersection9.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByIntersection9.Parameter("TrackFlag") = 1
    oCurveByIntersection9.Parameter("ConstructionSurface") = 0
    oCurveByIntersection9.Evaluate

    Dim oPointAtCurveExtremity10a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity10a = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0010a-PointAtCurveExtremity")
    oPointAtCurveExtremity10a.Inputs("Curve").Add oCurveByIntersection8, "1"
    oPointAtCurveExtremity10a.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oPointAtCurveExtremity10a.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity10a.Evaluate

    Dim oPointAtCurveExtremity10b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity10b = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0010b-PointAtCurveExtremity")
    oPointAtCurveExtremity10b.Inputs("Curve").Add oCurveByIntersection8, "1"
    oPointAtCurveExtremity10b.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oPointAtCurveExtremity10b.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity10b.Evaluate

    Dim oVectorTangentToCurve10c As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve10c = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "00010c-VectorTangentToCurve")
    oVectorTangentToCurve10c.Inputs("Curve").Add oCurveByIntersection8, "1"
    oVectorTangentToCurve10c.Inputs("Point").Add oPointAtCurveExtremity10b, "1"
    oVectorTangentToCurve10c.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    'oVectorTangentToCurve10c.Inputs("TrackPoint").Add oPointFromCS2a, "1"
    oVectorTangentToCurve10c.Parameter("Range") = 1#
    oVectorTangentToCurve10c.Parameter("Orientation") = 3
    oVectorTangentToCurve10c.Parameter("TrackFlag") = 1
    oVectorTangentToCurve10c.Evaluate
    
    Dim oPointAtCurveExtremity11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity11 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0011-PointAtCurveExtremity")
    oPointAtCurveExtremity11.Inputs("Curve").Add oCurveByIntersection9, "2"
    oPointAtCurveExtremity11.Inputs("TrackPoint").Add oPointAtCurveExtremity10a, "2"
    oPointAtCurveExtremity11.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity11.Evaluate

    Dim oCSByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints12 = oGCFactory.CreateEntity("CSByPoints", pPOM, "0012-CSByPoints")
    oCSByPoints12.Inputs("OriginPoint").Add oPointAtCurveExtremity10a, "2"
    oCSByPoints12.Inputs("AxisPoint1").Add oPointAtCurveExtremity11, "1"
    oCSByPoints12.Inputs("AxisPoint2").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oCSByPoints12.Parameter("AxesRoles") = 5
    oCSByPoints12.Parameter("CSOrientation") = 1
    oCSByPoints12.Evaluate

    Dim oCSByPoints13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints13 = oGCFactory.CreateEntity("CSByPoints", pPOM, "0013-CSByPoints")
    oCSByPoints13.Inputs("OriginPoint").Add oPointAtCurveExtremity11, "1"
    oCSByPoints13.Inputs("AxisPoint1").Add oPointAtCurveExtremity10a, "1"
    oCSByPoints13.Inputs("AxisPoint2").Add pGeometricConstruction.Inputs("TracePoint").Item(1), "1"
    oCSByPoints13.Parameter("AxesRoles") = 5
    oCSByPoints13.Parameter("CSOrientation") = 1
    oCSByPoints13.Evaluate

    If pGeometricConstruction.Parameter("Reference") = 1 Then
        Dim oPointFromCS15 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS15 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0015-PointFromCS")
        oPointFromCS15.Inputs("CoordinateSystem").Add oCSByPoints12, "1"
        oPointFromCS15.Parameter("X") = 0#
        oPointFromCS15.Parameter("Y") = pGeometricConstruction.Parameter("VWeb")
        oPointFromCS15.Parameter("Z") = pGeometricConstruction.Parameter("UWeb")
        oPointFromCS15.Evaluate
    
        Dim oCSByCS16 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByCS16 = oGCFactory.CreateEntity("CSByCS", pPOM, "0016-CSByCS")
        oCSByCS16.Inputs("CoordinateSystem").Add oCSByPoints12, "1"
        oCSByCS16.Inputs("Origin").Add oPointFromCS15, "2"
        oCSByCS16.Parameter("AxesRoles") = 1
        oCSByCS16.Parameter("CSOrientation") = 1
        oCSByCS16.Parameter("TrackFlag") = 1
        oCSByCS16.Evaluate
    Else
        Dim oPointFromCS14 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS14 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0014-PointFromCS")
        oPointFromCS14.Inputs("CoordinateSystem").Add oCSByPoints13, "2"
        oPointFromCS14.Parameter("X") = 0#
        oPointFromCS14.Parameter("Y") = pGeometricConstruction.Parameter("VFlange")
        oPointFromCS14.Parameter("Z") = pGeometricConstruction.Parameter("UFlange")
        oPointFromCS14.Evaluate
    
        Dim oCSByCS17 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByCS17 = oGCFactory.CreateEntity("CSByCS", pPOM, "0017-CSByCS")
        oCSByCS17.Inputs("CoordinateSystem").Add oCSByPoints12, "1"
        oCSByCS17.Inputs("Origin").Add oPointFromCS14, "2"
        oCSByCS17.Parameter("AxesRoles") = 1
        oCSByCS17.Parameter("CSOrientation") = 1
        oCSByCS17.Parameter("TrackFlag") = 1
        oCSByCS17.Evaluate
    End If
    
    Dim oLineFromCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS18 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0018-LineFromCS")
    oLineFromCS18.Inputs("CoordinateSystem").Add oCSByPoints5, "1"
    oLineFromCS18.Parameter("LookingAxis") = 1
    oLineFromCS18.Parameter("Length") = -0.5
    oLineFromCS18.Parameter("CSOrientation") = 1
    oLineFromCS18.Parameter("LineJustification") = 1
    oLineFromCS18.Evaluate
    
    Dim oCurveByCurves19 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves19 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "0019-CurveByCurves")
    oCurveByCurves19.Inputs("Curves").Add oVectorTangentToCurve10c, "1"
    oCurveByCurves19.Inputs("Curves").Add oCurveByIntersection8, "2"
    oCurveByCurves19.Inputs("Curves").Add oCurveByIntersection9, "3"
    oCurveByCurves19.Evaluate
    
    Dim oSurfByLinearExtrusion20 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion20 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0020-SurfByLinearExtrusion")
    oSurfByLinearExtrusion20.Inputs("PlanarCrossSection").Add oCurveByCurves19, "1"
    oSurfByLinearExtrusion20.Inputs("ExtrusionLine").Add oLineFromCS18, "1"
    oSurfByLinearExtrusion20.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    If pGeometricConstruction.Parameter("Reference") = 1 Then
        oGCMacro.Output("Point", 1) = oPointFromCS15.Output
        oGCMacro.Output("CoordinateSystem", 1) = oCSByCS16.Output
    Else
        oGCMacro.Output("Point", 1) = oPointFromCS14.Output
        oGCMacro.Output("CoordinateSystem", 1) = oCSByCS17.Output
    End If
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion20.Output

End Sub
Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, bIsOnRibbonBar As Boolean)
    Select Case sName
        Case "Reference": bIsOnRibbonBar = True
        Case "UWeb", "VWeb": bIsOnRibbonBar = pGC.Parameter("Reference") = 1
        Case "UFlange", "VFlange": bIsOnRibbonBar = pGC.Parameter("Reference") = 2
    End Select
End Sub
