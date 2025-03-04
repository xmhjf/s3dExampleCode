VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrimmingPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a MemberPart", "ISPSMemberPartPrismatic or ISPSDesignedMember  OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("Point", "Select a point", "{EA4AC251-14E7-11D2-B6F3-0060974CAEC9}", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    Call pGeometricConstructionDefinition.AddOutput(3, "Plane")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    If Not TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is IJStiffenerSystem Then
        Dim oLineFromCrossSection1 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection1 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "001-LineFromCrossSection")
        oLineFromCrossSection1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oLineFromCrossSection1.Inputs("Point").Add pGeometricConstruction.Inputs("Point").Item(1)
        oLineFromCrossSection1.Parameter("EdgeSelector") = 513
        oLineFromCrossSection1.Evaluate
        oLineFromCrossSection1.Public = True
        
        Dim oLineFromCrossSection2 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection2 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "002-LineFromCrossSection")
        oLineFromCrossSection2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oLineFromCrossSection2.Inputs("Point").Add pGeometricConstruction.Inputs("Point").Item(1)
        oLineFromCrossSection2.Parameter("EdgeSelector") = 258
        oLineFromCrossSection2.Evaluate
        oLineFromCrossSection2.Public = True
    
        Dim oLineFromCrossSection3 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection3 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "003-LineFromCrossSection")
        oLineFromCrossSection3.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
        oLineFromCrossSection3.Inputs("Point").Add pGeometricConstruction.Inputs("Point").Item(1)
        oLineFromCrossSection3.Parameter("EdgeSelector") = 257
        oLineFromCrossSection3.Evaluate
        oLineFromCrossSection3.Public = True
    
        Dim oCSByLines4 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines4 = oGCFactory.CreateEntity("CSByLines", pPOM, "004-CSByLines")
        oCSByLines4.Inputs("AxisLine1").Add oLineFromCrossSection1
        oCSByLines4.Inputs("AxisLine2").Add oLineFromCrossSection3
        oCSByLines4.Parameter("AxesRoles") = 1
        oCSByLines4.Parameter("CSOrientation") = 1
        oCSByLines4.Evaluate
        oCSByLines4.Public = True
    
        Dim oPointFromCS5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS5 = oGCFactory.CreateEntity("PointFromCS", pPOM, "005-PointFromCS")
        oPointFromCS5.Inputs("CoordinateSystem").Add oCSByLines4
        oPointFromCS5.Parameter("X") = -CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS5.Parameter("Y") = -CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS5.Parameter("Z") = 0#
        oPointFromCS5.Evaluate
        oPointFromCS5.Public = True
    
        Dim oPointFromCS6 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS6 = oGCFactory.CreateEntity("PointFromCS", pPOM, "006-PointFromCS")
        oPointFromCS6.Inputs("CoordinateSystem").Add oCSByLines4
        oPointFromCS6.Parameter("X") = 0#
        oPointFromCS6.Parameter("Y") = 0#
        oPointFromCS6.Parameter("Z") = 0#
        oPointFromCS6.Evaluate
        oPointFromCS6.Public = True
    
        Dim oPointAtCurveExtremity7 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity7 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "007-PointAtCurveExtremity")
        oPointAtCurveExtremity7.Inputs("Curve").Add oLineFromCrossSection2
        oPointAtCurveExtremity7.Inputs("TrackPoint").Add oPointFromCS6
        oPointAtCurveExtremity7.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity7.Evaluate
        oPointAtCurveExtremity7.Public = True
    
        Dim oPointAtCurveExtremity8 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity8 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "008-PointAtCurveExtremity")
        oPointAtCurveExtremity8.Inputs("Curve").Add oLineFromCrossSection2
        oPointAtCurveExtremity8.Inputs("TrackPoint").Add oPointFromCS6
        oPointAtCurveExtremity8.Parameter("TrackFlag") = 2
        oPointAtCurveExtremity8.Evaluate
        oPointAtCurveExtremity8.Public = True
    
        Dim oPointAtCurveExtremity9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity9 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "009-PointAtCurveExtremity")
        oPointAtCurveExtremity9.Inputs("Curve").Add oLineFromCrossSection3
        oPointAtCurveExtremity9.Inputs("TrackPoint").Add oPointFromCS6
        oPointAtCurveExtremity9.Parameter("TrackFlag") = 2
        oPointAtCurveExtremity9.Evaluate
        oPointAtCurveExtremity9.Public = True
    
        Dim oPointFromCS10 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS10 = oGCFactory.CreateEntity("PointFromCS", pPOM, "010-PointFromCS")
        oPointFromCS10.Inputs("CoordinateSystem").Add oCSByLines4
        oPointFromCS10.Inputs("Point").Add oPointAtCurveExtremity7
        oPointFromCS10.Parameter("X") = CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS10.Parameter("Y") = -CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS10.Parameter("Z") = 0#
        oPointFromCS10.Evaluate
        oPointFromCS10.Public = True
    
        Dim oPointFromCS11 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS11 = oGCFactory.CreateEntity("PointFromCS", pPOM, "011-PointFromCS")
        oPointFromCS11.Inputs("CoordinateSystem").Add oCSByLines4
        oPointFromCS11.Inputs("Point").Add oPointAtCurveExtremity8
        oPointFromCS11.Parameter("X") = CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS11.Parameter("Y") = CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS11.Parameter("Z") = 0#
        oPointFromCS11.Evaluate
        oPointFromCS11.Public = True
    
        Dim oPointFromCS12 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS12 = oGCFactory.CreateEntity("PointFromCS", pPOM, "012-PointFromCS")
        oPointFromCS12.Inputs("CoordinateSystem").Add oCSByLines4
        oPointFromCS12.Inputs("Point").Add oPointAtCurveExtremity9
        oPointFromCS12.Parameter("X") = -CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS12.Parameter("Y") = CDbl(pGeometricConstruction.Parameter("Extension"))
        oPointFromCS12.Parameter("Z") = 0#
        oPointFromCS12.Evaluate
        oPointFromCS12.Public = True
    
        Dim oLineByPoints13 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints13 = oGCFactory.CreateEntity("LineByPoints", pPOM, "013-LineByPoints")
        oLineByPoints13.Inputs("StartPoint").Add oPointFromCS12
        oLineByPoints13.Inputs("EndPoint").Add oPointFromCS5
        oLineByPoints13.Evaluate
        oLineByPoints13.Public = True
    
        Dim oLineByPoints14 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints14 = oGCFactory.CreateEntity("LineByPoints", pPOM, "014-LineByPoints")
        oLineByPoints14.Inputs("StartPoint").Add oPointFromCS11
        oLineByPoints14.Inputs("EndPoint").Add oPointFromCS12
        oLineByPoints14.Evaluate
        oLineByPoints14.Public = True
    
        Dim oLineByPoints15 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints15 = oGCFactory.CreateEntity("LineByPoints", pPOM, "015-LineByPoints")
        oLineByPoints15.Inputs("StartPoint").Add oPointFromCS10
        oLineByPoints15.Inputs("EndPoint").Add oPointFromCS11
        oLineByPoints15.Evaluate
        oLineByPoints15.Public = True
    
        Dim oLineByPoints16 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints16 = oGCFactory.CreateEntity("LineByPoints", pPOM, "016-LineByPoints")
        oLineByPoints16.Inputs("StartPoint").Add oPointFromCS5
        oLineByPoints16.Inputs("EndPoint").Add oPointFromCS10
        oLineByPoints16.Evaluate
        oLineByPoints16.Public = True
    
        Dim oCpxStringByCurves17 As SP3DGeometricConstruction.GeometricConstruction
        Set oCpxStringByCurves17 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "017-CpxStringByCurves")
        oCpxStringByCurves17.Inputs("Curves").Add oLineByPoints16
        oCpxStringByCurves17.Inputs("Curves").Add oLineByPoints15
        oCpxStringByCurves17.Inputs("Curves").Add oLineByPoints14
        oCpxStringByCurves17.Inputs("Curves").Add oLineByPoints13
        oCpxStringByCurves17.Evaluate
        oCpxStringByCurves17.Public = True
    
        Dim oPlaneByBoundaries18 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByBoundaries18 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM, "018-PlaneByBoundaries")
        oPlaneByBoundaries18.Inputs("OuterBoundary").Add oCpxStringByCurves17
        oPlaneByBoundaries18.Evaluate
        oPlaneByBoundaries18.Public = True
    Else
        Dim pProfileAttributes As IJProfileAttributes: Set pProfileAttributes = New ProfileUtils
        Dim oLandingCurve As Object: Call pProfileAttributes.GetRootLandingCurveFromProfile(pGeometricConstruction.Inputs("MemberPart").Item(1), oLandingCurve)

        Set oPlaneByBoundaries18 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "018-PlaneByPointNormal")
        oPlaneByBoundaries18.Inputs("Point").Add pGeometricConstruction.Inputs("Point").Item(1)
        oPlaneByBoundaries18.Inputs("Line").Add oLandingCurve
        oPlaneByBoundaries18.Parameter("Range") = 1#
        oPlaneByBoundaries18.Evaluate
    End If
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Plane", 1) = oPlaneByBoundaries18.Output
End Sub
