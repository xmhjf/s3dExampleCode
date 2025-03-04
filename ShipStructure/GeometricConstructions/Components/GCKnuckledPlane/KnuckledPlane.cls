VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckledPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum errors
    lKNUCKLE_HEIGHT_TOO_SMALL = 3
    lKNUCKLE_HEIGHT_TOO_BIG = 4
    
    lUNEXPECTED_INTERNAL_ERROR = 100
End Enum

Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("OriginPoint", "Select origin point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("XAxisPoint", "Select x-axis point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("YAxisPoint", "Select y-axis point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("OppositePoint", "Select opposite point", "IJPoint", 1, 1)
    Call pGeometricConstructionDefinition.AddOutput(6, "Surface")
    Call pGeometricConstructionDefinition.AddOutput(1, "Knuckle")
    Call pGeometricConstructionDefinition.AddParameter("Distance", "Distance", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Extension", "Extension", 8, 1, 59, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddErrorValue(lKNUCKLE_HEIGHT_TOO_SMALL, "ErrorValue3", "Knuckle height too small")
    Call pGeometricConstructionDefinition.AddErrorValue(lKNUCKLE_HEIGHT_TOO_BIG, "ErrorValue4", "Knuckle height too big")
    Call pGeometricConstructionDefinition.AddErrorValue(lUNEXPECTED_INTERNAL_ERROR, "ErrorValue100", "Unexpected error in KnuckledPlane")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
   
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    'Knuckled height should be less than dist from "XAxisPoint" to "OppositePoint"
    
    Dim oParamDistMeasureParallel1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel1 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "001-ParamDistMeasureParallel")
    oParamDistMeasureParallel1.Inputs("Graphics").Add pGeometricConstruction.Inputs("XAxisPoint").Item(1)
    oParamDistMeasureParallel1.Inputs("Graphics").Add pGeometricConstruction.Inputs("OppositePoint").Item(1)
    oParamDistMeasureParallel1.Evaluate
    
    If pGeometricConstruction.Parameter("Distance") > oParamDistMeasureParallel1.Parameter("Value") Then Err.Raise lKNUCKLE_HEIGHT_TOO_BIG
    
    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = pGeometricConstruction.Parameter("Distance")
    oParamDistConstant1.Evaluate

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Extension")
    oParamDistConstant2.Evaluate

    Dim oLineByPoints3 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints3 = oGCFactory.CreateEntity("LineByPoints", pPOM, "003-LineByPoints")
    oLineByPoints3.Inputs("StartPoint").Add pGeometricConstruction.Inputs("OriginPoint").Item(1)
    oLineByPoints3.Inputs("EndPoint").Add pGeometricConstruction.Inputs("YAxisPoint").Item(1)
    oLineByPoints3.Evaluate
    
    
    Dim oLineByPoints4 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints4 = oGCFactory.CreateEntity("LineByPoints", pPOM, "004-LineByPoints")
    oLineByPoints4.Inputs("StartPoint").Add pGeometricConstruction.Inputs("YAxisPoint").Item(1)
    oLineByPoints4.Inputs("EndPoint").Add pGeometricConstruction.Inputs("OppositePoint").Item(1)
    oLineByPoints4.Evaluate

    'Create line parallel to oLineByPoints3 at OppositePoint and infinite on the other side
    Dim oCSByLines7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines7 = oGCFactory.CreateEntity("CSByLines", pPOM, "007-CSByLines")
    oCSByLines7.Inputs("AxisLine1").Add oLineByPoints3
    oCSByLines7.Inputs("AxisLine2").Add oLineByPoints4
    oCSByLines7.Parameter("AxesRoles") = 1
    oCSByLines7.Parameter("CSOrientation") = 1
    oCSByLines7.Evaluate

    Dim oCSByCs7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCs7 = oGCFactory.CreateEntity("CSByCS", pPOM, "007-CSByCs")
    oCSByCs7.Inputs("CoordinateSystem").Add oCSByLines7
    oCSByCs7.Inputs("Origin").Add pGeometricConstruction.Inputs("OppositePoint").Item(1)
    oCSByCs7.Parameter("AxesRoles") = 1
    oCSByCs7.Parameter("CSOrientation") = 1
    oCSByCs7.Evaluate

    ' axis x of oCSByCs7
    Dim oLineFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7 = oGCFactory.CreateEntity("LineFromCS", pPOM, "007-LineFromCS")
    oLineFromCS7.Inputs("CoordinateSystem").Add oCSByCs7
    oLineFromCS7.Parameter("LookingAxis") = 1
    oLineFromCS7.Parameter("Length") = 100#
    oLineFromCS7.Parameter("CSOrientation") = 1
    oLineFromCS7.Parameter("LineJustification") = 1
    oLineFromCS7.Evaluate
    
    Dim oPointAtCurveEnd7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd7 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "007-PointAtCurveEnd")
    oPointAtCurveEnd7.Inputs("Curve").Add oLineFromCS7
    oPointAtCurveEnd7.Evaluate
    
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        Dim oSphereByCenterRadius5 As SP3DGeometricConstruction.GeometricConstruction
        Set oSphereByCenterRadius5 = oGCFactory.CreateEntity("SphereByCenterRadius", pPOM, "005-SphereByCenterRadius")
        oSphereByCenterRadius5.Inputs("Center").Add pGeometricConstruction.Inputs("XAxisPoint").Item(1)
        oSphereByCenterRadius5.Parameter("Radius") = oParamDistConstant1
        oSphereByCenterRadius5.Evaluate
    End If
    
    Dim oPointAlongCurve6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "006-PointAlongCurve")
    oPointAlongCurve6.Inputs("Curve").Add oLineByPoints3
    oPointAlongCurve6.Inputs("Point").Add pGeometricConstruction.Inputs("YAxisPoint").Item(1)
    oPointAlongCurve6.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("OriginPoint").Item(1)
    oPointAlongCurve6.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve6.Parameter("TrackFlag") = 2
    oPointAlongCurve6.Evaluate


    Dim oPointByCurveAndSurface8 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        Set oPointByCurveAndSurface8 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "008-PointByCurveAndSurface")
        oPointByCurveAndSurface8.Inputs("Curve1").Add oLineFromCS7
        oPointByCurveAndSurface8.Inputs("Surface2").Add oSphereByCenterRadius5
        oPointByCurveAndSurface8.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("OppositePoint").Item(1)
        oPointByCurveAndSurface8.Parameter("TrackFlag") = 1
        On Error Resume Next
        oPointByCurveAndSurface8.Evaluate
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise lKNUCKLE_HEIGHT_TOO_SMALL
        End If
    Else
        Set oPointByCurveAndSurface8 = pGeometricConstruction.Inputs("OppositePoint").Item(1)
    End If
    
    Dim oLineByPoints9 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints9 = oGCFactory.CreateEntity("LineByPoints", pPOM, "009-LineByPoints")
    oLineByPoints9.Inputs("StartPoint").Add pGeometricConstruction.Inputs("OriginPoint").Item(1)
    oLineByPoints9.Inputs("EndPoint").Add oPointByCurveAndSurface8
    oLineByPoints9.Evaluate

    Dim oLineByPoints10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints10 = oGCFactory.CreateEntity("LineByPoints", pPOM, "010-LineByPoints")
    oLineByPoints10.Inputs("StartPoint").Add pGeometricConstruction.Inputs("XAxisPoint").Item(1)
    oLineByPoints10.Inputs("EndPoint").Add oPointByCurveAndSurface8
    oLineByPoints10.Evaluate

    Dim oPointAlongCurve11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "011-PointAlongCurve")
    oPointAlongCurve11.Inputs("Curve").Add oLineByPoints9
    oPointAlongCurve11.Inputs("Point").Add pGeometricConstruction.Inputs("OriginPoint").Item(1)
    oPointAlongCurve11.Inputs("TrackPoint").Add oPointByCurveAndSurface8
    oPointAlongCurve11.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve11.Parameter("TrackFlag") = 2
    oPointAlongCurve11.Evaluate

    Dim oPointAlongCurve12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve12 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "012-PointAlongCurve")
    oPointAlongCurve12.Inputs("Curve").Add oLineFromCS7
    oPointAlongCurve12.Inputs("Point").Add pGeometricConstruction.Inputs("OppositePoint").Item(1)
    oPointAlongCurve12.Inputs("TrackPoint").Add oPointAtCurveEnd7
    oPointAlongCurve12.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve12.Parameter("TrackFlag") = 2
    oPointAlongCurve12.Evaluate

    Dim oPointAlongCurve13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve13 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "013-PointAlongCurve")
    oPointAlongCurve13.Inputs("Curve").Add oLineByPoints10
    oPointAlongCurve13.Inputs("Point").Add pGeometricConstruction.Inputs("XAxisPoint").Item(1)
    oPointAlongCurve13.Inputs("TrackPoint").Add oPointByCurveAndSurface8
    oPointAlongCurve13.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve13.Parameter("TrackFlag") = 2
    oPointAlongCurve13.Evaluate

    Dim oPointAlongCurve13bis As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointAlongCurve14 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        Set oPointAlongCurve14 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "014-PointAlongCurve")
        oPointAlongCurve14.Inputs("Curve").Add oLineByPoints9
        oPointAlongCurve14.Inputs("Point").Add oPointByCurveAndSurface8
        oPointAlongCurve14.Inputs("TrackPoint").Add oPointAlongCurve11
        oPointAlongCurve14.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve14.Parameter("TrackFlag") = 2
        oPointAlongCurve14.Evaluate
    Else
        'knuckled point for full knuckled (to create extended edge "XAxisPoint"-"OppositePoint" for output surface)
        Set oPointAlongCurve13bis = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "013-PointAlongCurveBis")
        oPointAlongCurve13bis.Inputs("Curve").Add oLineByPoints10
        oPointAlongCurve13bis.Inputs("Point").Add oPointByCurveAndSurface8
        oPointAlongCurve13bis.Inputs("TrackPoint").Add pGeometricConstruction.Inputs("XAxisPoint").Item(1)
        oPointAlongCurve13bis.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve13bis.Parameter("TrackFlag") = 2
        oPointAlongCurve13bis.Evaluate
        
        Set oPointAlongCurve14 = oPointAlongCurve13bis
    End If
    
    Dim oCSByPoints15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints15 = oGCFactory.CreateEntity("CSByPoints", pPOM, "015-CSByPoints")
    oCSByPoints15.Inputs("OriginPoint").Add oPointAlongCurve12
    oCSByPoints15.Inputs("AxisPoint1").Add oPointAlongCurve6
    oCSByPoints15.Inputs("AxisPoint2").Add oPointByCurveAndSurface8
    oCSByPoints15.Parameter("AxesRoles") = 1
    oCSByPoints15.Parameter("CSOrientation") = 1
    oCSByPoints15.Evaluate

    Dim oLineByPoints16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints16 = oGCFactory.CreateEntity("LineByPoints", pPOM, "016-LineByPoints")
    oLineByPoints16.Inputs("StartPoint").Add oPointAlongCurve6
    oLineByPoints16.Inputs("EndPoint").Add oPointAlongCurve12
    oLineByPoints16.Evaluate

    Dim oLineByPoints17 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints17 = oGCFactory.CreateEntity("LineByPoints", pPOM, "017-LineByPoints")
    oLineByPoints17.Inputs("StartPoint").Add oPointAlongCurve11
    oLineByPoints17.Inputs("EndPoint").Add oPointAlongCurve13
    oLineByPoints17.Evaluate

    Dim oPointAlongCurve18 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        Set oPointAlongCurve18 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "018-PointAlongCurve")
        oPointAlongCurve18.Inputs("Curve").Add oLineByPoints16
        oPointAlongCurve18.Inputs("Point").Add oPointAlongCurve12
        oPointAlongCurve18.Inputs("TrackPoint").Add oPointAlongCurve6
        oPointAlongCurve18.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve18.Parameter("TrackFlag") = 2
        oPointAlongCurve18.Evaluate
    Else
        Set oPointAlongCurve18 = oPointAlongCurve14
    End If
    
    Dim oPointAlongCurve19 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve19 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "019-PointAlongCurve")
    oPointAlongCurve19.Inputs("Curve").Add oLineByPoints16
    oPointAlongCurve19.Inputs("Point").Add oPointAlongCurve6
    oPointAlongCurve19.Inputs("TrackPoint").Add oPointAlongCurve12
    oPointAlongCurve19.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve19.Parameter("TrackFlag") = 2
    oPointAlongCurve19.Evaluate

    Dim oLineByPointAngleLength20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPointAngleLength20 = oGCFactory.CreateEntity("LineByPointAngleLength", pPOM, "020-LineByPointAngleLength")
    oLineByPointAngleLength20.Inputs("CoordinateSystem").Add oCSByPoints15
    oLineByPointAngleLength20.Inputs("Point").Add oPointAlongCurve19
    oLineByPointAngleLength20.Parameter("AxesRoles") = 2
    oLineByPointAngleLength20.Parameter("Angle") = 3.132866
    oLineByPointAngleLength20.Parameter("Length") = oParamDistConstant2
    oLineByPointAngleLength20.Parameter("TrackFlag") = 1
    oLineByPointAngleLength20.Evaluate

    Dim oPointAlongCurve21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve21 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "021-PointAlongCurve")
    oPointAlongCurve21.Inputs("Curve").Add oLineByPoints17
    oPointAlongCurve21.Inputs("Point").Add oPointAlongCurve13
    oPointAlongCurve21.Inputs("TrackPoint").Add oPointAlongCurve11
    oPointAlongCurve21.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve21.Parameter("TrackFlag") = 2
    oPointAlongCurve21.Evaluate

    Dim oLineByPointAngleLength22 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPointAngleLength22 = oGCFactory.CreateEntity("LineByPointAngleLength", pPOM, "022-LineByPointAngleLength")
    oLineByPointAngleLength22.Inputs("CoordinateSystem").Add oCSByPoints15
    oLineByPointAngleLength22.Inputs("Point").Add oPointAlongCurve18
    oLineByPointAngleLength22.Parameter("AxesRoles") = 2
    oLineByPointAngleLength22.Parameter("Angle") = 3.132866
    oLineByPointAngleLength22.Parameter("Length") = oParamDistConstant2
    oLineByPointAngleLength22.Parameter("TrackFlag") = 1
    oLineByPointAngleLength22.Evaluate

    Dim oLineStringByPoints23 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineStringByPoints23 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "023-LineStringByPoints")
    oLineStringByPoints23.Inputs("Points").Add oPointAlongCurve11
    oLineStringByPoints23.Inputs("Points").Add oPointAlongCurve14
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        oLineStringByPoints23.Inputs("Points").Add oPointAlongCurve18
    End If
    oLineStringByPoints23.Inputs("Points").Add oPointAlongCurve19
    oLineStringByPoints23.Parameter("ClosedFlag") = 1
    oLineStringByPoints23.Evaluate

    Dim oLineStringByPoints24 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineStringByPoints24 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "024-LineStringByPoints")
    oLineStringByPoints24.Inputs("Points").Add oPointAlongCurve11
    oLineStringByPoints24.Inputs("Points").Add oPointAlongCurve21
    oLineStringByPoints24.Inputs("Points").Add oPointAlongCurve14
    oLineStringByPoints24.Parameter("ClosedFlag") = 1
    oLineStringByPoints24.Evaluate

    Dim oPointAtCurveEnd25 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd25 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "025-PointAtCurveEnd")
    oPointAtCurveEnd25.Inputs("Curve").Add oLineByPointAngleLength20
    oPointAtCurveEnd25.Evaluate

    Dim oPointAtCurveEnd26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd26 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "026-PointAtCurveEnd")
    oPointAtCurveEnd26.Inputs("Curve").Add oLineByPointAngleLength22
    oPointAtCurveEnd26.Evaluate

    Dim oCpxStringByCurves27 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves27 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "027-CpxStringByCurves")
    oCpxStringByCurves27.Inputs("Curves").Add oLineStringByPoints23
    oCpxStringByCurves27.Evaluate

    Dim oCpxStringByCurves28 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves28 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "028-CpxStringByCurves")
    oCpxStringByCurves28.Inputs("Curves").Add oLineStringByPoints24
    oCpxStringByCurves28.Evaluate

    Dim oLineStringByPoints29 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineStringByPoints29 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "029-LineStringByPoints")
    oLineStringByPoints29.Inputs("Points").Add oPointAlongCurve19
    oLineStringByPoints29.Inputs("Points").Add oPointAlongCurve18
    oLineStringByPoints29.Inputs("Points").Add oPointAtCurveEnd26
    oLineStringByPoints29.Inputs("Points").Add oPointAtCurveEnd25
    oLineStringByPoints29.Parameter("ClosedFlag") = 1
    oLineStringByPoints29.Evaluate

    Dim oPlaneByBoundaries30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByBoundaries30 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM, "030-PlaneByBoundaries")
    oPlaneByBoundaries30.Inputs("OuterBoundary").Add oCpxStringByCurves27
    oPlaneByBoundaries30.Evaluate

    Dim oPlaneByBoundaries31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByBoundaries31 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM, "031-PlaneByBoundaries")
    oPlaneByBoundaries31.Inputs("OuterBoundary").Add oCpxStringByCurves28
    oPlaneByBoundaries31.Evaluate

    Dim oCpxStringByCurves32 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves32 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "032-CpxStringByCurves")
    oCpxStringByCurves32.Inputs("Curves").Add oLineStringByPoints29
    oCpxStringByCurves32.Evaluate

    Dim oPlaneByBoundaries33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByBoundaries33 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM, "033-PlaneByBoundaries")
    oPlaneByBoundaries33.Inputs("OuterBoundary").Add oCpxStringByCurves32
    oPlaneByBoundaries33.Evaluate

    ' monikerize planes
    Dim pSurfaceBody30 As IJSurfaceBody: Set pSurfaceBody30 = ModelBody_FromGeometry(oPlaneByBoundaries30): Call SurfaceBody_Monikerize(pSurfaceBody30, 999, 1, 8, 1)
    Dim pSurfaceBody31 As IJSurfaceBody: Set pSurfaceBody31 = ModelBody_FromGeometry(oPlaneByBoundaries31): Call SurfaceBody_Monikerize(pSurfaceBody31, 999, 1, 8, 2)
    Dim pSurfaceBody33 As IJSurfaceBody: Set pSurfaceBody33 = ModelBody_FromGeometry(oPlaneByBoundaries33): Call SurfaceBody_Monikerize(pSurfaceBody33, 999, 1, 8, 3)
    
    Dim oSurfBySurfaces34 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfBySurfaces34 = oGCFactory.CreateEntity("SurfBySurfaces", pPOM, "034-SurfBySurfaces")
    oSurfBySurfaces34.Inputs("Surfaces").Add pSurfaceBody30
    oSurfBySurfaces34.Inputs("Surfaces").Add pSurfaceBody31
    oSurfBySurfaces34.Inputs("Surfaces").Add pSurfaceBody33
    oSurfBySurfaces34.Evaluate

    Dim oSurfWithKnuckles35 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfWithKnuckles35 = oGCFactory.CreateEntity("SurfWithKnuckles", pPOM, "035-SurfWithKnuckles")
    oSurfWithKnuckles35.Inputs("Surface").Add oSurfBySurfaces34
    oSurfWithKnuckles35.Parameter("Angle") = 0.01745329
    oSurfWithKnuckles35.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Surface", 1) = oSurfWithKnuckles35.Output
    If pGeometricConstruction.Parameter("Distance") > 0 Then
        oGCMacro.Output("Knuckle", 1) = oPointByCurveAndSurface8.Output
    Else
        oGCMacro.Output("Knuckle", 1) = oPointAlongCurve13bis.Output
    End If
        
    
End Sub
