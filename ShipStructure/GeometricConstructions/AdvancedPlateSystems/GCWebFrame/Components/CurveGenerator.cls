VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurveGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1, "IJSurface IJSurfaceBody")
    Call pGeometricConstructionDefinition.AddInput("MiddleBoundary", "Select middle boundary", "IJSurface OR IJSurfaceBody", 1, 1, "IJSurface IJSurfaceBody")
    Call pGeometricConstructionDefinition.AddInput("LowerCurve", "Select lower curve", "IJCurve OR IJWireBody", 1, 1, "IJCurve IJWireBody")
    Call pGeometricConstructionDefinition.AddInput("UpperCurve", "Select upper curve", "IJCurve OR IJWireBody", 1, 1, "IJCurve IJWireBody")
    Call pGeometricConstructionDefinition.AddInput("MiddlePoint", "Select middle point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("LowerPoint", "Select lower point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("UpperPoint", "Select upper point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("RootPlateSystem", "Select root plate system", "IJPlateSystem AND [GCSHPSTRDefinitions.ShipFilters,IsRootPlateSystem]", 1, 1, "IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select coordinate system", "IJDCoordinateSystem", 1, 1)
    
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem", "IJPlate IJDPlateMoldedConventions")
    
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddOutput(7, "Curve")
    
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oCurveByIntersection1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection1 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0001-CurveByIntersection")
    oCurveByIntersection1.Inputs("Surface1").Add pGeometricConstruction.Inputs("MiddleBoundary").Item(1), "1"
    oCurveByIntersection1.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByIntersection1.Parameter("TrackFlag") = 1
    oCurveByIntersection1.Parameter("ConstructionSurface") = 0
    oCurveByIntersection1.Evaluate

    If Not TypeOf pGeometricConstruction.Inputs("MiddleBoundary").Item(1) Is IJPlane Then
        Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0002-LineFromCS")
        oLineFromCS2.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
        oLineFromCS2.Parameter("LookingAxis") = 2
        oLineFromCS2.Parameter("Length") = 5#
        oLineFromCS2.Parameter("CSOrientation") = 1
        oLineFromCS2.Parameter("LineJustification") = 1
        oLineFromCS2.Evaluate
    
'        Dim oPointFromCS3 As SP3DGeometricConstruction.GeometricConstruction
'        Set oPointFromCS3 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0003-PointFromCS")
'        oPointFromCS3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
'        oPointFromCS3.Parameter("X") = 0#
'        oPointFromCS3.Parameter("Y") = 1#
'        oPointFromCS3.Parameter("Z") = 0#
'        oPointFromCS3.Evaluate
    
        Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0004-ParamDistConstant")
        oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
        oParamDistConstant4.Evaluate
    
        Dim oParamDouble5 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDouble5 = oGCFactory.CreateEntity("ParamDouble", pPOM, "0005-ParamDouble")
        oParamDouble5.Parameter("Value") = 3#
        oParamDouble5.Evaluate
    
        Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0006-ParamDistConstant")
        oParamDistConstant6.Parameter("Value") = 1.5
        oParamDistConstant6.Evaluate
    
        Dim oCurveExtractPartial7 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveExtractPartial7 = oGCFactory.CreateEntity("CurveExtractPartial", pPOM, "0007-CurveExtractPartial")
        oCurveExtractPartial7.Inputs("Curve").Add oCurveByIntersection1, "1"
        oCurveExtractPartial7.Inputs("Splitters").Add pGeometricConstruction.Inputs("LowerPoint").Item(1), "1"
        oCurveExtractPartial7.Inputs("Splitters").Add pGeometricConstruction.Inputs("UpperPoint").Item(1), "2"
        oCurveExtractPartial7.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveExtractPartial7.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "2"
        oCurveExtractPartial7.Parameter("TrackFlag") = 1
        oCurveExtractPartial7.Parameter("ConstructionSurface") = 0
        oCurveExtractPartial7.Evaluate
    
        Dim oParamDistMeasureLength8 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureLength8 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "0008-ParamDistMeasureLength")
        oParamDistMeasureLength8.Inputs("Curve").Add oCurveExtractPartial7, "1"
        oParamDistMeasureLength8.Parameter("Value") = 20.6232273778236
        oParamDistMeasureLength8.Evaluate
    
        Dim oParamDistFunctionWith2Args9 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args9 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0009-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args9.Inputs("Arg1").Add oParamDistMeasureLength8, "1"
        oParamDistFunctionWith2Args9.Inputs("Arg2").Add oParamDouble5, "1"
        oParamDistFunctionWith2Args9.Parameter("Value") = 6.87440912594122
        oParamDistFunctionWith2Args9.Parameter("Operation") = 4
        oParamDistFunctionWith2Args9.Evaluate
    
        Dim oCurveAtOffset10 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveAtOffset10 = oGCFactory.CreateEntity("CurveAtOffset", pPOM, "0010-CurveAtOffset")
        oCurveAtOffset10.Inputs("Curve").Add oCurveExtractPartial7, "1"
        oCurveAtOffset10.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveAtOffset10.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oCurveAtOffset10.Parameter("Offset") = oParamDistConstant4
        oCurveAtOffset10.Parameter("Orientation") = 1
        If WebFrameLocation = 2 Then  'outside
            oCurveAtOffset10.Parameter("TrackFlag") = 2
        Else
            oCurveAtOffset10.Parameter("TrackFlag") = 1
        End If
        oCurveAtOffset10.Evaluate
        
        Dim oPointAtCurveExtremity11 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity11 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0011-PointAtCurveExtremity")
        oPointAtCurveExtremity11.Inputs("Curve").Add oCurveAtOffset10, "1"
        oPointAtCurveExtremity11.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("LowerPoint").Item(1), "1"
        oPointAtCurveExtremity11.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity11.Evaluate
    
        Dim oPointAtCurveExtremity12 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity12 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0012-PointAtCurveExtremity")
        oPointAtCurveExtremity12.Inputs("Curve").Add oCurveAtOffset10, "1"
        oPointAtCurveExtremity12.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("UpperPoint").Item(1), "1"
        oPointAtCurveExtremity12.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity12.Evaluate
    
        Dim oVectorTangentToCurve13 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve13 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0013-VectorTangentToCurve")
        oVectorTangentToCurve13.Inputs("Curve").Add oCurveAtOffset10, "1"
        oVectorTangentToCurve13.Inputs("Point").Add oPointAtCurveExtremity11, "1"
        oVectorTangentToCurve13.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oVectorTangentToCurve13.Parameter("Range") = 100#
        oVectorTangentToCurve13.Parameter("Orientation") = 3
        oVectorTangentToCurve13.Parameter("TrackFlag") = 1
        oVectorTangentToCurve13.Evaluate
    
        Dim oVectorTangentToCurve14 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve14 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0014-VectorTangentToCurve")
        oVectorTangentToCurve14.Inputs("Curve").Add oCurveAtOffset10, "1"
        oVectorTangentToCurve14.Inputs("Point").Add oPointAtCurveExtremity11, "1"
        oVectorTangentToCurve14.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oVectorTangentToCurve14.Parameter("Range") = 100#
        oVectorTangentToCurve14.Parameter("Orientation") = 2
        oVectorTangentToCurve14.Parameter("TrackFlag") = 1
        oVectorTangentToCurve14.Evaluate
    
        Dim oVectorTangentToCurve15 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve15 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0015-VectorTangentToCurve")
        oVectorTangentToCurve15.Inputs("Curve").Add oCurveAtOffset10, "1"
        oVectorTangentToCurve15.Inputs("Point").Add oPointAtCurveExtremity12, "1"
        oVectorTangentToCurve15.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oVectorTangentToCurve15.Parameter("Range") = 100#
        oVectorTangentToCurve15.Parameter("Orientation") = 2
        oVectorTangentToCurve15.Parameter("TrackFlag") = 1
        oVectorTangentToCurve15.Evaluate
    
        Dim oVectorTangentToCurve16 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve16 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0016-VectorTangentToCurve")
        oVectorTangentToCurve16.Inputs("Curve").Add oCurveAtOffset10, "1"
        oVectorTangentToCurve16.Inputs("Point").Add oPointAtCurveExtremity12, "1"
        oVectorTangentToCurve16.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oVectorTangentToCurve16.Parameter("Range") = 100#
        oVectorTangentToCurve16.Parameter("Orientation") = 3
        oVectorTangentToCurve16.Parameter("TrackFlag") = 1
        oVectorTangentToCurve16.Evaluate
    
        Dim oCpxStringByCurves17 As SP3DGeometricConstruction.GeometricConstruction
        Set oCpxStringByCurves17 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0017-CpxStringByCurves")
        oCpxStringByCurves17.Inputs("Curves").Add oVectorTangentToCurve13, "1"
        oCpxStringByCurves17.Inputs("Curves").Add oVectorTangentToCurve14, "2"
        oCpxStringByCurves17.Evaluate
    
        Dim oCpxStringByCurves18 As SP3DGeometricConstruction.GeometricConstruction
        Set oCpxStringByCurves18 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0018-CpxStringByCurves")
        oCpxStringByCurves18.Inputs("Curves").Add oVectorTangentToCurve16, "1"
        oCpxStringByCurves18.Inputs("Curves").Add oVectorTangentToCurve15, "2"
        oCpxStringByCurves18.Evaluate
    
        Dim oPointByCurves19 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves19 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0019-PointByCurves")
        oPointByCurves19.Inputs("Curve1").Add oCpxStringByCurves17, "1"
        oPointByCurves19.Inputs("Curve2").Add pGeometricConstruction.Inputs("LowerCurve").Item(1), "1"
        oPointByCurves19.Inputs("TrackPoint").Add oPointAtCurveExtremity11, "1"
        oPointByCurves19.Parameter("TrackFlag") = 1
        oPointByCurves19.Evaluate
    
        Dim oPointByCurves20 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves20 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0020-PointByCurves")
        oPointByCurves20.Inputs("Curve1").Add pGeometricConstruction.Inputs("UpperCurve").Item(1), "1"
        oPointByCurves20.Inputs("Curve2").Add oCpxStringByCurves18, "1"
        oPointByCurves20.Inputs("TrackPoint").Add oPointAtCurveExtremity12, "1"
        oPointByCurves20.Parameter("TrackFlag") = 1
        oPointByCurves20.Evaluate
    
        Dim oPointByProjectOnCurve21 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByProjectOnCurve21 = oGCFactory.CreateEntity("PointByProjectOnCurve", pPOM, "0021-PointByProjectOnCurve")
        oPointByProjectOnCurve21.Inputs("Point").Add oPointByCurves19, "1"
        oPointByProjectOnCurve21.Inputs("Curve").Add oCurveByIntersection1, "1"
        oPointByProjectOnCurve21.Evaluate
    
        Dim oPointByProjectOnCurve22 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByProjectOnCurve22 = oGCFactory.CreateEntity("PointByProjectOnCurve", pPOM, "0022-PointByProjectOnCurve")
        oPointByProjectOnCurve22.Inputs("Point").Add oPointByCurves20, "2"
        oPointByProjectOnCurve22.Inputs("Curve").Add oCurveByIntersection1, "1"
        oPointByProjectOnCurve22.Evaluate
    
        Dim oPointAlongCurve23 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve23 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0023-PointAlongCurve")
        oPointAlongCurve23.Inputs("Curve").Add oCurveByIntersection1, "2"
        oPointAlongCurve23.Inputs("Point").Add oPointByProjectOnCurve21, "1"
        oPointAlongCurve23.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oPointAlongCurve23.Parameter("Distance") = oParamDistFunctionWith2Args9
        oPointAlongCurve23.Parameter("TrackFlag") = 2
        On Error Resume Next
        oPointAlongCurve23.Evaluate
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set oPointAlongCurve23 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0023a-PointAtCurveExtremity")
            oPointAlongCurve23.Inputs("Curve").Add oCurveByIntersection1, "2"
            oPointAlongCurve23.Inputs("TrackPoint").Add oPointByProjectOnCurve21, "1"
            oPointAlongCurve23.Evaluate
        End If
        On Error GoTo 0
    
        Dim oPointAlongCurve24 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve24 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0024-PointAlongCurve")
        oPointAlongCurve24.Inputs("Curve").Add oCurveByIntersection1, "1"
        oPointAlongCurve24.Inputs("Point").Add oPointByProjectOnCurve22, "2"
        oPointAlongCurve24.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oPointAlongCurve24.Parameter("Distance") = oParamDistFunctionWith2Args9
        oPointAlongCurve24.Parameter("TrackFlag") = 2
        On Error Resume Next
        oPointAlongCurve24.Evaluate
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set oPointAlongCurve24 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0024a-PointAtCurveExtremity")
            oPointAlongCurve24.Inputs("Curve").Add oCurveByIntersection1, "2"
            oPointAlongCurve24.Inputs("TrackPoint").Add oPointByProjectOnCurve22, "1"
            oPointAlongCurve24.Evaluate
        End If
        On Error GoTo 0
    
        Dim oLineByPoints25 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints25 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0025-LineByPoints")
        oLineByPoints25.Inputs("StartPoint").Add oPointAlongCurve23, "3"
        oLineByPoints25.Inputs("EndPoint").Add oPointAlongCurve24, "1"
        oLineByPoints25.Evaluate
    
        Dim oLineFromCS26 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS26 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0026-LineFromCS")
        oLineFromCS26.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
        oLineFromCS26.Inputs("ReferencePoint").Add oPointAlongCurve23, "2"
        oLineFromCS26.Parameter("LookingAxis") = 1
        oLineFromCS26.Parameter("Length") = 1#
        oLineFromCS26.Parameter("CSOrientation") = 1
        oLineFromCS26.Parameter("LineJustification") = 1
        oLineFromCS26.Evaluate
    
        Dim oCSByLines27 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines27 = oGCFactory.CreateEntity("CSByLines", pPOM, "0027-CSByLines")
        oCSByLines27.Inputs("AxisLine1").Add oLineFromCS26, "1"
        oCSByLines27.Inputs("AxisLine2").Add oLineByPoints25, "1"
        oCSByLines27.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oCSByLines27.Parameter("AxesRoles") = 6
        oCSByLines27.Parameter("CSOrientation") = 1
        oCSByLines27.Parameter("TrackFlag") = 1
        oCSByLines27.Evaluate
        
        Dim oCurveExtractPartial28 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveExtractPartial28 = oGCFactory.CreateEntity("CurveExtractPartial", pPOM, "0028-CurveExtractPartial")
        oCurveExtractPartial28.Inputs("Curve").Add oCurveByIntersection1, "1"
        oCurveExtractPartial28.Inputs("Splitters").Add oPointAlongCurve23, "1"
        oCurveExtractPartial28.Inputs("Splitters").Add oPointAlongCurve24, "2"
        oCurveExtractPartial28.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveExtractPartial28.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "2"
        oCurveExtractPartial28.Parameter("TrackFlag") = 1
        oCurveExtractPartial28.Parameter("ConstructionSurface") = 0
        oCurveExtractPartial28.Evaluate
        
        Dim oPointsAsRangeBox29 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointsAsRangeBox29 = oGCFactory.CreateEntity("PointsAsRangeBox", pPOM, "0029-PointsAsRangeBox")
        oPointsAsRangeBox29.Inputs("CoordinateSystem").Add oCSByLines27, "1"
        oPointsAsRangeBox29.Inputs("Geometries").Add oCurveExtractPartial28, "1"
        oPointsAsRangeBox29.Evaluate
    
        Dim oPointFromCS30 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS30 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0030-PointFromCS")
        oPointFromCS30.Inputs("CoordinateSystem").Add oCSByLines27, "1"
        oPointFromCS30.Inputs("Point").Add oPointsAsRangeBox29.Output("Points", "Low"), "1"
        oPointFromCS30.Parameter("X") = -0.5
        oPointFromCS30.Parameter("Y") = -1#
        oPointFromCS30.Parameter("Z") = 0#
        oPointFromCS30.Evaluate

        Dim oPointFromCS31 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS31 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0031-PointFromCS")
        oPointFromCS31.Inputs("CoordinateSystem").Add oCSByLines27, "1"
        oPointFromCS31.Inputs("Point").Add oPointsAsRangeBox29.Output("Points", "High"), "1"
        oPointFromCS31.Parameter("X") = 0.5
        oPointFromCS31.Parameter("Y") = 1#
        oPointFromCS31.Parameter("Z") = 0#
        oPointFromCS31.Evaluate

        Dim oSurfAsBox32 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfAsBox32 = oGCFactory.CreateEntity("SurfAsBox", pPOM, "0032-SurfAsBox")
        oSurfAsBox32.Inputs("CoordinateSystem").Add oCSByLines27, "1"
        oSurfAsBox32.Inputs("PointLow").Add oPointFromCS30, "1"
        oSurfAsBox32.Inputs("PointHigh").Add oPointFromCS31, "1"
        oSurfAsBox32.Evaluate
        
        Dim oFacePortGenerator33 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortGenerator33 = oGCFactory.CreateEntity("FacePortGenerator", pPOM, "0033-FacePortGenerator")
        oFacePortGenerator33.Inputs("RootPlateSystem").Add pGeometricConstruction.Inputs("RootPlateSystem").Item(1), "1"
        oFacePortGenerator33.Inputs("ReferencePoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oFacePortGenerator33.Inputs("SplittingTool").Add oSurfAsBox32, "1"
        oFacePortGenerator33.Parameter("OutputRadius") = 1#
        oFacePortGenerator33.Parameter("GeometrySelector") = 4
        oFacePortGenerator33.Evaluate
        pGeometricConstruction.ControlledInputs("LeafPlateSystem").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem").Add oFacePortGenerator33.ControlledInputs("LeafPlateSystem").Item(1)
    Else
        Set oFacePortGenerator33 = oGCFactory.CreateEntity("FacePortGenerator", pPOM, "0033-FacePortGenerator")
        oFacePortGenerator33.Inputs("RootPlateSystem").Add pGeometricConstruction.Inputs("RootPlateSystem").Item(1), "1"
        oFacePortGenerator33.Inputs("ReferencePoint").Add pGeometricConstruction.Inputs("MiddlePoint").Item(1), "1"
        oFacePortGenerator33.Parameter("OutputRadius") = 1#
        oFacePortGenerator33.Parameter("GeometrySelector") = 4
        oFacePortGenerator33.Evaluate
        pGeometricConstruction.ControlledInputs("LeafPlateSystem").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem").Add oFacePortGenerator33.ControlledInputs("LeafPlateSystem").Item(1)
    End If
    
    Dim oCurveByIntersection37 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection37 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0037-CurveByIntersection")
    oCurveByIntersection37.Inputs("Surface1").Add oFacePortGenerator33, "1"
    oCurveByIntersection37.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCurveByIntersection37.Parameter("TrackFlag") = 1
    oCurveByIntersection37.Parameter("ConstructionSurface") = 0
    oCurveByIntersection37.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Curve", 1) = oCurveByIntersection37.Output
End Sub
