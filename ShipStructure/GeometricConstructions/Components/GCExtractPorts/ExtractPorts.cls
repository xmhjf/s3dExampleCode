VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtractPorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a rolled Member or a designed Member or a Stiffener system", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCCMNSTRDefinitions.Filters,IsRootStiffenerSystem]", 1, 1, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddInput("ZAxis", "Select line normal to the sketching plane near the nodal point", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("XAxis", "Select axis line", "IJLine OR IJWireBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJLine]", 0, 1)
    
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem", "IJContactNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem", "IJPlate IJDPlateMoldedConventions")

    Call pGeometricConstructionDefinition.AddParameter("Support", "Support", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Down", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Support", "Up", 2)
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "MemberAxis")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "CurveRight")
    Call pGeometricConstructionDefinition.AddOutput(GCGTypeLine3d, "CurveLeft")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oLineAxisPortExtractor0 As SP3DGeometricConstruction.GeometricConstruction
    Dim oMemberPart As Object: Set oMemberPart = pGeometricConstruction.Inputs("MemberPart").Item(1)
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
        Set oLineAxisPortExtractor0 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
        oLineAxisPortExtractor0.Inputs("MemberPart").Add oMemberPart
        oLineAxisPortExtractor0.Evaluate
    Else
        Set oLineAxisPortExtractor0 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "001-AxisPortExtractor")
        oLineAxisPortExtractor0.Input("Connectable") = oMemberPart
        oLineAxisPortExtractor0.Parameter("GeometrySelector") = 4 ' GCStable
        oLineAxisPortExtractor0.Evaluate
    End If
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("AxisPort"), oLineAxisPortExtractor0.ControlledInputs("Port"))

    Dim oLineAxisPortExtractor0a As Object
    If pGeometricConstruction.Inputs("XAxis").Count > 0 Then
        Set oLineAxisPortExtractor0a = pGeometricConstruction.Inputs("XAxis").Item(1)
    Else
        Set oLineAxisPortExtractor0a = oLineAxisPortExtractor0
    End If

'    Dim oParamDistMeasureLength1 As SP3DGeometricConstruction.GeometricConstruction
'    Set oParamDistMeasureLength1 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM)
'    oParamDistMeasureLength1.Inputs("Curve").Add oLineAxisPortExtractor0a ' pGeometricConstruction.Inputs("XAxis").Item(1)
'    oParamDistMeasureLength1.Parameter("Value") = 0 ' pGeometricConstruction.Parameter("GCParamDistMeasureLength-1-0003")
'    oParamDistMeasureLength1.Public = True
'    oParamDistMeasureLength1.Evaluate

    Dim oCSByLines2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines2 = oGCFactory.CreateEntity("CSByLines", pPOM, "002-CSByLines")
    oCSByLines2.Inputs("AxisLine1").Add pGeometricConstruction.Inputs("ZAxis").Item(1)
    oCSByLines2.Inputs("AxisLine2").Add oLineAxisPortExtractor0a ' pGeometricConstruction.Inputs("XAxis").Item(1)
    oCSByLines2.Parameter("AxesRoles") = 3
    oCSByLines2.Parameter("CSOrientation") = 1
    oCSByLines2.Public = True
    oCSByLines2.Evaluate
    
    ' extract y axis
    Dim oLineFromCS2a As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS2a = oGCFactory.CreateEntity("LineFromCS", pPOM, "003-LineFromCS")
    oLineFromCS2a.Inputs("CoordinateSystem").Add oCSByLines2
    oLineFromCS2a.Parameter("LookingAxis") = 2
    oLineFromCS2a.Parameter("Length") = 1
    oLineFromCS2a.Evaluate
    
    ' make sure the x-axis of the CS is the line axis
    Dim oCSByLines2b As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines2b = oGCFactory.CreateEntity("CSByLines", pPOM, "004-CSByLines")
    oCSByLines2b.Inputs("AxisLine1").Add oLineAxisPortExtractor0a
    oCSByLines2b.Inputs("AxisLine2").Add oLineFromCS2a
    oCSByLines2b.Parameter("AxesRoles") = 1
    oCSByLines2b.Parameter("CSOrientation") = 1
    oCSByLines2b.Public = True
    oCSByLines2b.Evaluate
    
    ' use origin of CS to position closer and further extremities of the line axis
    Dim oPointFromCS2c As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS2c = oGCFactory.CreateEntity("PointFromCS", pPOM, "005-PointFromCS")
    oPointFromCS2c.Inputs("CoordinateSystem").Add oCSByLines2b
    oPointFromCS2c.Parameter("X") = 0
    oPointFromCS2c.Parameter("Y") = 0
    oPointFromCS2c.Parameter("Z") = 0
    oPointFromCS2c.Evaluate
    
    ' find the closer extremity of the line axis
    Dim oPointAtCurveExtremity2d As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity2d = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "006-PointAtCurveExtremity")
    oPointAtCurveExtremity2d.Inputs("Curve").Add oLineAxisPortExtractor0a
    oPointAtCurveExtremity2d.Inputs("TrackPoint").Add oPointFromCS2c
    oPointAtCurveExtremity2d.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity2d.Evaluate
    
    ' find the further extremity of the line axis
    Dim oPointAtCurveExtremity2e As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity2e = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "007-PointAtCurveExtremity")
    oPointAtCurveExtremity2e.Inputs("Curve").Add oLineAxisPortExtractor0a
    oPointAtCurveExtremity2e.Inputs("TrackPoint").Add oPointAtCurveExtremity2d
    oPointAtCurveExtremity2e.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity2e.Evaluate
    
    ' position a point along the curve from the closer extremity
    Dim oPointAlongCurve2f As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve2f = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "008-PointAlongCurve")
    oPointAlongCurve2f.Inputs("Curve").Add oLineAxisPortExtractor0a
    oPointAlongCurve2f.Inputs("Point").Add oPointAtCurveExtremity2d
    oPointAlongCurve2f.Inputs("TrackPoint").Add oPointAtCurveExtremity2e
    oPointAlongCurve2f.Parameter("Distance") = 0.1
    oPointAlongCurve2f.Parameter("TrackFlag") = 1
    oPointAlongCurve2f.Evaluate
    
    ' position a point above
    Dim oPointFromCS2g As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS2g = oGCFactory.CreateEntity("PointFromCS", pPOM, "009-PointFromCS")
    oPointFromCS2g.Inputs("CoordinateSystem").Add oCSByLines2b
    oPointFromCS2g.Inputs("Point").Add oPointAlongCurve2f
    oPointFromCS2g.Parameter("X") = 0
    oPointFromCS2g.Parameter("Y") = 0
    oPointFromCS2g.Parameter("Z") = 1
    oPointFromCS2g.Evaluate

    ' define the new z-axis
    Dim oLineByPoints2h As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints2h = oGCFactory.CreateEntity("LineByPoints", pPOM, "010-LineByPoints")
    oLineByPoints2h.Inputs("StartPoint").Add oPointAlongCurve2f
    oLineByPoints2h.Inputs("EndPoint").Add oPointFromCS2g
    oLineByPoints2h.Evaluate

    ' define a CS, whose z-axis will intersect the ports
    Dim oCSByLines2i As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines2i = oGCFactory.CreateEntity("CSByLines", pPOM, "011-CSByLines")
    oCSByLines2i.Inputs("AxisLine1").Add oLineByPoints2h
    oCSByLines2i.Inputs("AxisLine2").Add oLineAxisPortExtractor0a ' pGeometricConstruction.Inputs("XAxis").Item(1)
    oCSByLines2i.Parameter("AxesRoles") = 3
    oCSByLines2i.Parameter("CSOrientation") = 1
    oCSByLines2i.Public = True
    oCSByLines2i.Evaluate

'    Dim oPointAtCurveMiddle3 As SP3DGeometricConstruction.GeometricConstruction
'    Set oPointAtCurveMiddle3 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM)
'    oPointAtCurveMiddle3.Inputs("Curve").Add oLineAxisPortExtractor0a
'    oPointAtCurveMiddle3.Inputs("TrackPoint").Add oPointAtCurveExtremity2
'    oPointAtCurveMiddle3.Public = True
'    oPointAtCurveMiddle3.Evaluate

'    Dim oPlaneFromCS4 As SP3DGeometricConstruction.GeometricConstruction
'    Set oPlaneFromCS4 = oGCFactory.CreateEntity("PlaneFromCS", pPOM)
'    oPlaneFromCS4.Inputs("CoordinateSystem").Add oCSByLines2
'    oPlaneFromCS4.Parameter("LookingAxis") = 3
'    oPlaneFromCS4.Parameter("Offset") = 0#
'    oPlaneFromCS4.Parameter("Range") = 0.1
'    oPlaneFromCS4.Public = True
'    oPlaneFromCS4.Evaluate

    Dim oFacePortExtractor5 As SP3DGeometricConstruction.GeometricConstruction
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
        pGeometricConstruction.ControlledInputs("RootPlateSystem").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem").Clear
        
        Set oFacePortExtractor5 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "021-FacePortExtractor")
        oFacePortExtractor5.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oFacePortExtractor5.Inputs("CoordinateSystem").Add oCSByLines2i
        oFacePortExtractor5.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort")
        oFacePortExtractor5.Parameter("GeometrySelector") = 4
        oFacePortExtractor5.Parameter("FacesContext") = 2
        oFacePortExtractor5.Parameter("LookingAxis") = 3
        oFacePortExtractor5.Parameter("SurfaceType") = 2 '1
        oFacePortExtractor5.Parameter("TrackFlag") = CInt(pGeometricConstruction.Parameter("Support"))
        oFacePortExtractor5.Parameter("Offset") = 0#
        oFacePortExtractor5.Public = True
        oFacePortExtractor5.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort"), oFacePortExtractor5.ControlledInputs("Port"))
    
        Dim oEdgePortExtractor8 As SP3DGeometricConstruction.GeometricConstruction
        Set oEdgePortExtractor8 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "022-EdgePortExtractor")
        oEdgePortExtractor8.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oEdgePortExtractor8.Inputs("FacePort").Add oFacePortExtractor5
        oEdgePortExtractor8.Inputs("CoordinateSystem").Add oCSByLines2i
        oEdgePortExtractor8.Parameter("GeometrySelector") = 4
    '    oEdgePortExtractor8.Parameter("FacesContext") = 2
        oEdgePortExtractor8.Parameter("LookingAxis") = 2
        oEdgePortExtractor8.Parameter("TrackFlag") = 2
        oEdgePortExtractor8.Public = True
        oEdgePortExtractor8.Evaluate
    
        Dim oEdgePortExtractor9 As SP3DGeometricConstruction.GeometricConstruction
        Set oEdgePortExtractor9 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "023-EdgePortExtractor")
        oEdgePortExtractor9.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oEdgePortExtractor9.Inputs("FacePort").Add oFacePortExtractor5
        oEdgePortExtractor9.Inputs("CoordinateSystem").Add oCSByLines2i
        oEdgePortExtractor9.Parameter("GeometrySelector") = 4
    '    oEdgePortExtractor9.Parameter("FacesContext") = 2
        oEdgePortExtractor9.Parameter("LookingAxis") = 2
        oEdgePortExtractor9.Parameter("TrackFlag") = 1
        oEdgePortExtractor9.Public = True
        oEdgePortExtractor9.Evaluate
   Else
        pGeometricConstruction.ControlledInputs("FacePort").Clear
        
        Dim oPointFromCS5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS5 = oGCFactory.CreateEntity("PointFromCS", pPOM, "021-PointFromCS")
        oPointFromCS5.Inputs("CoordinateSystem").Add oCSByLines2i
        oPointFromCS5.Parameter("X") = 0#
        oPointFromCS5.Parameter("Y") = 0#
        oPointFromCS5.Parameter("Z") = 0#
        oPointFromCS5.Evaluate
       
        Dim oPointAtCurveExtremity5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity5 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "022-PointAtCurveExtremity")
        oPointAtCurveExtremity5.Inputs("Curve").Add oLineAxisPortExtractor0a
        oPointAtCurveExtremity5.Inputs("TrackPoint").Add oPointFromCS5
        oPointAtCurveExtremity5.Parameter("TrackFlag") = 1 ' near
        oPointAtCurveExtremity5.Evaluate
        
        Dim oTubeFromDetailedProfile6 As SP3DGeometricConstruction.GeometricConstruction
        Set oTubeFromDetailedProfile6 = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "023-TubeFromDetailedProfile")
        oTubeFromDetailedProfile6.Inputs("RootProfileSystem").Add oMemberPart
        oTubeFromDetailedProfile6.Inputs("ReferencePoint").Add oPointAtCurveExtremity5
        oTubeFromDetailedProfile6.Parameter("Length") = 0#
        oTubeFromDetailedProfile6.Evaluate
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("RootPlateSystem"), oTubeFromDetailedProfile6.ControlledInputs("RootPlateSystem"))
        Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("LeafPlateSystem"), oTubeFromDetailedProfile6.ControlledInputs("LeafPlateSystem"))

        Dim oSurfaceBodyExtractor7 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfaceBodyExtractor7 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "024-SurfaceBodyExtractor")
        oSurfaceBodyExtractor7.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile6
        oSurfaceBodyExtractor7.Inputs("CoordinateSystem").Add oCSByLines2i
        oSurfaceBodyExtractor7.Parameter("FacesContext") = 2
        oSurfaceBodyExtractor7.Parameter("LookingAxis") = 3
        oSurfaceBodyExtractor7.Parameter("IntersectingPlane") = 0
        oSurfaceBodyExtractor7.Parameter("SurfaceType") = 1
        oSurfaceBodyExtractor7.Parameter("TrackFlag") = CInt(pGeometricConstruction.Parameter("Support")) '1
        oSurfaceBodyExtractor7.Parameter("Offset") = 0#
        oSurfaceBodyExtractor7.Evaluate
        
        ' face of solid body always orented outside
        Dim pSurfaceBody As IJSurfaceBody: Set pSurfaceBody = oSurfaceBodyExtractor7
        Call pSurfaceBody.ReverseTopologyOrientation
        
        ' compute distance betweeen mid point and surface
        Dim oPointAtSurfaceCenter5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtSurfaceCenter5 = oGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "025-PointAtSurfaceCenter")
        oPointAtSurfaceCenter5.Inputs("Surface").Add oSurfaceBodyExtractor7
        oPointAtSurfaceCenter5.Evaluate

        Dim oPointAlongCurve5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve5 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "026-PointAlongCurve")
        oPointAlongCurve5.Inputs("Curve").Add oLineAxisPortExtractor0a
        oPointAlongCurve5.Inputs("Point").Add oPointAtSurfaceCenter5
        oPointAlongCurve5.Parameter("Distance") = 0
        oPointAlongCurve5.Parameter("TrackFlag") = 1
        oPointAlongCurve5.Evaluate

        Dim oParamDistMeasureParallel7 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel7 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "027-ParamDistMeasureParallel")
        oParamDistMeasureParallel7.Inputs("Graphics").Add oSurfaceBodyExtractor7
        oParamDistMeasureParallel7.Inputs("Graphics").Add oPointAlongCurve5
        oParamDistMeasureParallel7.Evaluate
        
        ' offset the surface to go through the mid point
        If CDbl(oParamDistMeasureParallel7.Parameter("Value")) > EPSILON Then
            Dim oSurfAtOffset7 As SP3DGeometricConstruction.GeometricConstruction
            Set oSurfAtOffset7 = oGCFactory.CreateEntity("SurfAtOffset", pPOM, "028-SurfAtOffset7")
            oSurfAtOffset7.Inputs("Surface").Add oSurfaceBodyExtractor7
            oSurfAtOffset7.Inputs("TrackPoint").Add oPointAlongCurve5
            oSurfAtOffset7.Parameter("Offset") = CDbl(oParamDistMeasureParallel7.Parameter("Value"))
            oSurfAtOffset7.Parameter("TrackFlag") = 1
            oSurfAtOffset7.Evaluate
        
            Set oSurfaceBodyExtractor7 = oSurfAtOffset7
        End If
        
        Dim oPointFromCS8 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS8 = oGCFactory.CreateEntity("PointFromCS", pPOM, "029-PointFromCS")
        oPointFromCS8.Inputs("CoordinateSystem").Add oCSByLines2i
        oPointFromCS8.Inputs("Point").Add oPointAlongCurve5
        oPointFromCS8.Parameter("X") = 0#
        oPointFromCS8.Parameter("Y") = 1#
        oPointFromCS8.Parameter("Z") = 0#
        oPointFromCS8.Evaluate
    
        Dim oEdgeFromCurveOrSurf8 As SP3DGeometricConstruction.GeometricConstruction
        Set oEdgeFromCurveOrSurf8 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "030-EdgeFromCurveOrSurf")
        oEdgeFromCurveOrSurf8.Inputs("Geometry").Add oSurfaceBodyExtractor7
        oEdgeFromCurveOrSurf8.Inputs("TrackPoint").Add oPointFromCS8
        oEdgeFromCurveOrSurf8.Parameter("TrackFlag") = 1
        oEdgeFromCurveOrSurf8.Evaluate
        
        Dim oPointFromCS9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS9 = oGCFactory.CreateEntity("PointFromCS", pPOM, "031-PointFromCS")
        oPointFromCS9.Inputs("CoordinateSystem").Add oCSByLines2i
        oPointFromCS9.Inputs("Point").Add oPointAlongCurve5
        oPointFromCS9.Parameter("X") = 0#
        oPointFromCS9.Parameter("Y") = -1#
        oPointFromCS9.Parameter("Z") = 0#
        oPointFromCS9.Evaluate
            
        Dim oEdgeFromCurveOrSurf9 As SP3DGeometricConstruction.GeometricConstruction
        Set oEdgeFromCurveOrSurf9 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "032-EdgeFromCurveOrSurf")
        oEdgeFromCurveOrSurf9.Inputs("Geometry").Add oSurfaceBodyExtractor7
        oEdgeFromCurveOrSurf9.Inputs("TrackPoint").Add oPointFromCS9
        oEdgeFromCurveOrSurf9.Parameter("TrackFlag") = 1
        oEdgeFromCurveOrSurf9.Evaluate
        
        Set oFacePortExtractor5 = oSurfaceBodyExtractor7
        Set oEdgePortExtractor8 = oEdgeFromCurveOrSurf8
        Set oEdgePortExtractor9 = oEdgeFromCurveOrSurf9
    End If

'    Dim oCircleByCenterRadius6 As SP3DGeometricConstruction.GeometricConstruction
'    Set oCircleByCenterRadius6 = oGCFactory.CreateEntity("CircleByCenterRadius", pPOM)
'    oCircleByCenterRadius6.Inputs("Center").Add oPointAtCurveMiddle3
'    oCircleByCenterRadius6.Inputs("Plane").Add oPlaneFromCS4
'    oCircleByCenterRadius6.Parameter("Radius") = oParamDistMeasureLength1
'    oCircleByCenterRadius6.Public = True
'    oCircleByCenterRadius6.Evaluate
'
'    Dim oCpxStringByCurves7 As SP3DGeometricConstruction.GeometricConstruction
'    Set oCpxStringByCurves7 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM)
'    oCpxStringByCurves7.Inputs("Curves").Add oCircleByCenterRadius6
'    oCpxStringByCurves7.Public = True
'    oCpxStringByCurves7.Evaluate

    Dim oCurveOrientate8a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate8a = oGCFactory.CreateEntity("CurveOrientate", pPOM, "041-CurveOrientate")
    oCurveOrientate8a.Inputs("Curve").Add oEdgePortExtractor8
    oCurveOrientate8a.Inputs("CoordinateSystem").Add oCSByLines2i
    oCurveOrientate8a.Public = True
    oCurveOrientate8a.Evaluate
    
    Dim oCurveOrientate9a As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate9a = oGCFactory.CreateEntity("CurveOrientate", pPOM, "042-CurveOrientate9")
    oCurveOrientate9a.Inputs("Curve").Add oEdgePortExtractor9
    oCurveOrientate9a.Inputs("CoordinateSystem").Add oCSByLines2i
    oCurveOrientate9a.Public = True
    oCurveOrientate9a.Evaluate
   
    Dim oCurveOrientate10 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate10 = oGCFactory.CreateEntity("CurveOrientate", pPOM, "043-CurveOrientate")
    oCurveOrientate10.Inputs("Curve").Add oLineAxisPortExtractor0a
    oCurveOrientate10.Inputs("CoordinateSystem").Add oCSByLines2i
    oCurveOrientate10.Public = True
    oCurveOrientate10.Evaluate
   
'''    Dim oCurveByProjection11 As SP3DGeometricConstruction.GeometricConstruction
'''    Set oCurveByProjection11 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "018-CurveByProjection")
'''    oCurveByProjection11.Inputs("Curve").Add oCurveOrientate10
'''    oCurveByProjection11.Inputs("Surface").Add oFacePortExtractor5
'''    oCurveByProjection11.Inputs("Line").Add oLineByPoints2h ' pGeometricConstruction.Input("ZAxis")
'''    oCurveByProjection11.Inputs("CoordinateSystem").Add oCSByLines2i
'''    oCurveByProjection11.Parameter("TrackFlag") = 1
'''    oCurveByProjection11.Public = True
'''    oCurveByProjection11.Evaluate

    ' create the MemberAxis as a line at equal distance from curve left and curve right
    Dim oPointAtCurveStart8a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart8a = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "044-PointAtCurveStart")
    oPointAtCurveStart8a.Inputs("Curve").Add oCurveOrientate8a
    oPointAtCurveStart8a.Public = True
    oPointAtCurveStart8a.Evaluate
    
    Dim oPointAtCurveEnd8a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd8a = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "045-PointAtCurveEnd")
    oPointAtCurveEnd8a.Inputs("Curve").Add oCurveOrientate8a
    oPointAtCurveEnd8a.Public = True
    oPointAtCurveEnd8a.Evaluate
    
    Dim oPointAtCurveStart9a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart9a = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "046PointAtCurveStart")
    oPointAtCurveStart9a.Inputs("Curve").Add oCurveOrientate9a
    oPointAtCurveStart9a.Public = True
    oPointAtCurveStart9a.Evaluate
    
    Dim oPointAtCurveEnd9a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd9a = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "047-PointAtCurveEnd")
    oPointAtCurveEnd9a.Inputs("Curve").Add oCurveOrientate9a
    oPointAtCurveEnd9a.Public = True
    oPointAtCurveEnd9a.Evaluate
    
    Dim oLineByPoints10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints10 = oGCFactory.CreateEntity("LineByPoints", pPOM, "048-LineByPoints")
    oLineByPoints10.Inputs("StartPoint").Add oPointAtCurveStart8a
    oLineByPoints10.Inputs("EndPoint").Add oPointAtCurveStart9a
    oLineByPoints10.Public = True
    oLineByPoints10.Evaluate
    
    Dim oLineByPoints11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints11 = oGCFactory.CreateEntity("LineByPoints", pPOM, "049-LineByPoints")
    oLineByPoints11.Inputs("StartPoint").Add oPointAtCurveEnd8a
    oLineByPoints11.Inputs("EndPoint").Add oPointAtCurveEnd9a
    oLineByPoints11.Public = True
    oLineByPoints11.Evaluate
    
    Dim oPointAtCurveMiddle12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "050-PointAtCurveMiddle")
    oPointAtCurveMiddle12.Inputs("Curve").Add oLineByPoints10
    oPointAtCurveMiddle12.Public = True
    oPointAtCurveMiddle12.Evaluate

    Dim oPointAtCurveMiddle13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle13 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "051-PointAtCurveMiddle")
    oPointAtCurveMiddle13.Inputs("Curve").Add oLineByPoints11
    oPointAtCurveMiddle13.Public = True
    oPointAtCurveMiddle13.Evaluate

    Dim oLineByPoints14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints14 = oGCFactory.CreateEntity("LineByPoints", pPOM, "052-LineByPoints")
    oLineByPoints14.Inputs("StartPoint").Add oPointAtCurveMiddle12
    oLineByPoints14.Inputs("EndPoint").Add oPointAtCurveMiddle13
    oLineByPoints14.Public = True
    oLineByPoints14.Evaluate
    
    ' reproject the member axis
    Dim oPointAtCurveStart10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart10 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "053-PointAtCurveStart")
    oPointAtCurveStart10.Inputs("Curve").Add oCurveOrientate10
    oPointAtCurveStart10.Public = True
    oPointAtCurveStart10.Evaluate
    
    Dim oPointAtCurveEnd10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd10 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "054-PointAtCurveEnd")
    oPointAtCurveEnd10.Inputs("Curve").Add oCurveOrientate10
    oPointAtCurveEnd10.Public = True
    oPointAtCurveEnd10.Evaluate
    
    Dim oPointAlongCurve10a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve10a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "055-PointAlongCurve")
    oPointAlongCurve10a.Inputs("Curve").Add oLineByPoints14
    oPointAlongCurve10a.Inputs("Point").Add oPointAtCurveStart10
    oPointAlongCurve10a.Parameter("Distance") = 0#
    oPointAlongCurve10a.Parameter("TrackFlag") = 1
    oPointAlongCurve10a.Evaluate
    
    Dim oPointAlongCurve10b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve10b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "056-PointAlongCurve")
    oPointAlongCurve10b.Inputs("Curve").Add oLineByPoints14
    oPointAlongCurve10b.Inputs("Point").Add oPointAtCurveEnd10
    oPointAlongCurve10b.Parameter("Distance") = 0#
    oPointAlongCurve10b.Parameter("TrackFlag") = 1
    oPointAlongCurve10b.Evaluate
      
    Dim oLineByPoints15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints15 = oGCFactory.CreateEntity("LineByPoints", pPOM, "057-LineByPoints")
    oLineByPoints15.Inputs("StartPoint").Add oPointAlongCurve10a
    oLineByPoints15.Inputs("EndPoint").Add oPointAlongCurve10b
    oLineByPoints15.Public = True
    oLineByPoints15.Evaluate
    
    Dim oPointAlongCurve11a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "058-PointAlongCurve")
    oPointAlongCurve11a.Inputs("Curve").Add oCurveOrientate8a
    oPointAlongCurve11a.Inputs("Point").Add oPointAtCurveStart10
    oPointAlongCurve11a.Parameter("Distance") = 0#
    oPointAlongCurve11a.Parameter("TrackFlag") = 1
    oPointAlongCurve11a.Evaluate
    
    Dim oPointAlongCurve11b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "059-PointAlongCurve")
    oPointAlongCurve11b.Inputs("Curve").Add oCurveOrientate8a
    oPointAlongCurve11b.Inputs("Point").Add oPointAtCurveEnd10
    oPointAlongCurve11b.Parameter("Distance") = 0#
    oPointAlongCurve11b.Parameter("TrackFlag") = 1
    oPointAlongCurve11b.Evaluate
      
    Dim oLineByPoints16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints16 = oGCFactory.CreateEntity("LineByPoints", pPOM, "060-LineByPoints")
    oLineByPoints16.Inputs("StartPoint").Add oPointAlongCurve11a
    oLineByPoints16.Inputs("EndPoint").Add oPointAlongCurve11b
    oLineByPoints16.Public = True
    oLineByPoints16.Evaluate

    Dim oPointAlongCurve12a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve12a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "061-PointAlongCurve")
    oPointAlongCurve12a.Inputs("Curve").Add oCurveOrientate9a
    oPointAlongCurve12a.Inputs("Point").Add oPointAtCurveStart10
    oPointAlongCurve12a.Parameter("Distance") = 0#
    oPointAlongCurve12a.Parameter("TrackFlag") = 1
    oPointAlongCurve12a.Evaluate
    
    Dim oPointAlongCurve12b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve12b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "062-PointAlongCurve")
    oPointAlongCurve12b.Inputs("Curve").Add oCurveOrientate9a
    oPointAlongCurve12b.Inputs("Point").Add oPointAtCurveEnd10
    oPointAlongCurve12b.Parameter("Distance") = 0#
    oPointAlongCurve11b.Parameter("TrackFlag") = 1
    oPointAlongCurve12b.Evaluate
      
    Dim oLineByPoints17 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints17 = oGCFactory.CreateEntity("LineByPoints", pPOM, "063-LineByPoints")
    oLineByPoints17.Inputs("StartPoint").Add oPointAlongCurve12a
    oLineByPoints17.Inputs("EndPoint").Add oPointAlongCurve12b
    oLineByPoints17.Public = True
    oLineByPoints17.Evaluate

  
'    Dim oPlaneByBoundaries10 As SP3DGeometricConstruction.GeometricConstruction
'    Set oPlaneByBoundaries10 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM)
'    oPlaneByBoundaries10.Inputs("OuterBoundary").Add oCpxStringByCurves7
'    oPlaneByBoundaries10.Public = True
'    oPlaneByBoundaries10.Evaluate
'
'    Dim oCurveByProjection11 As SP3DGeometricConstruction.GeometricConstruction
'    Set oCurveByProjection11 = oGCFactory.CreateEntity("CurveByProjection", pPOM)
'    oCurveByProjection11.Inputs("Curve").Add oCurveOrientate9a
'    oCurveByProjection11.Inputs("Surface").Add oPlaneByBoundaries10
'    oCurveByProjection11.Inputs("Line").Add pGeometricConstruction.Inputs("ZAxis").Item(1)
'    oCurveByProjection11.Parameter("TrackFlag") = 1
'    oCurveByProjection11.Public = True
'    oCurveByProjection11.Evaluate
'
'    Dim oCurveByProjection12 As SP3DGeometricConstruction.GeometricConstruction
'    Set oCurveByProjection12 = oGCFactory.CreateEntity("CurveByProjection", pPOM)
'    oCurveByProjection12.Inputs("Curve").Add oCurveOrientate8a
'    oCurveByProjection12.Inputs("Surface").Add oPlaneByBoundaries10
'    oCurveByProjection12.Inputs("Line").Add pGeometricConstruction.Inputs("ZAxis").Item(1)
'    oCurveByProjection12.Parameter("TrackFlag") = 1
'    oCurveByProjection12.Public = True
'    oCurveByProjection12.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    ' orientate the Support towards the inside of the built-up
    Call SurfaceBody_OrientateFollowingLine(oFacePortExtractor5.Output, pGeometricConstruction.Input("ZAxis"), CInt(pGeometricConstruction.Parameter("Support")) = 1)
    
    oGCMacro.Output("Support", 1) = oFacePortExtractor5.Output
    oGCMacro.Output("MemberAxis", 1) = oLineByPoints15.Output
    oGCMacro.Output("CurveRight", 1) = oLineByPoints17.Output ' oCurveOrientate9a.Output
    oGCMacro.Output("CurveLeft", 1) = oLineByPoints16.Output ' oCurveOrientate8a.Output
End Sub
Sub Elements_ReplaceElements(pElementsToBeReplaced As IJElements, pElementsReplacing As IJElements)
    If pElementsToBeReplaced.Count = 1 Then
        If Not pElementsToBeReplaced(1) Is pElementsReplacing(1) Then
            pElementsToBeReplaced.Clear
            pElementsToBeReplaced.AddElements pElementsReplacing
        End If
    Else
        pElementsToBeReplaced.Clear
        pElementsToBeReplaced.AddElements pElementsReplacing
    End If
End Sub
