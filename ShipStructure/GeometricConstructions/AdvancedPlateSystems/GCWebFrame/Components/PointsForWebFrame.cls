VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PointsForWebFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal vInfo As Variant, bIsOnRibbonBar As Boolean)
    Select Case sName
'        Case "Location"
'            bIsOnRibbonBar = pGC.Parameter("BoundaryCount") = 2
        Case "Orientation"
            bIsOnRibbonBar = pGC.Parameter("BoundaryCount") = 1
        Case "UpperConnectable"
            bIsOnRibbonBar = pGC.Parameter("BoundaryCount") = 2
    End Select
End Sub
'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SketchingPlane", "Select sketching plane", "IJPlane", 1, 1, "IJSurface IJSurfaceBody")
    Call pGeometricConstructionDefinition.AddInput("MiddleConnectable", "Select a middle plate system", "IJPlateSystem AND [GCCMNSTRDefinitions.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("LowerConnectable", "Select a lower plate/profile system or plane", "IJPlane AND [GCWebFrame.Filters,IsNotPlateSystem] OR IJPlateSystem AND [GCCMNSTRDefinitions.Filters,IsRootPlateSystem OR IJStiffenerSystem AND [GCCMNSTRDefinitions.Filters,IsRootStiffenerSystem]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("UpperConnectable", "Select an upper plate/profile system or plane", "IJPlane AND [GCWebFrame.Filters,IsNotPlateSystem] OR IJPlateSystem AND [GCCMNSTRDefinitions.Filters,IsRootPlateSystem OR IJStiffenerSystem AND [GCCMNSTRDefinitions.Filters,IsRootStiffenerSystem]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("BoundaryCount", "Boundary Count", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("BoundaryCount", "One single boundary", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("BoundaryCount", "Two separate boundaries", 2)
    
    Call pGeometricConstructionDefinition.AddParameter("Side", "Side", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Positive", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Side", "Negative", 2)
    
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", GCCodeList, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Inside", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Outside", 2)
    
    If False Then
        Call pGeometricConstructionDefinition.AddParameter("Orientation", "Orientation", GCCodeList, 0, 0, 0, 0, 1)
        Call pGeometricConstructionDefinition.AddParameterValue("Orientation", "Positive", 1)
        Call pGeometricConstructionDefinition.AddParameterValue("Orientation", "Negative", 2)
    End If
    
    Call pGeometricConstructionDefinition.AddControlledInput("Port1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port2")
    Call pGeometricConstructionDefinition.AddControlledInput("Port3")
    Call pGeometricConstructionDefinition.AddControlledInput("Port4")
    Call pGeometricConstructionDefinition.AddControlledInput("Port5")
    
    Call pGeometricConstructionDefinition.AddOutput(1, "Points")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCLocalCoordinateSystem, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundaries")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCWireBody2, "Curves")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCGTypeLine3d, "OrientationLine")

    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

'   ' Dim lSide As Long: Let lSide = pGeometricConstruction.Parameter("Side")
'    Dim lLocation As Long: Let lLocation = pGeometricConstruction.Parameter("Location")
'    If TypeOf pGeometricConstruction.Input("MiddleConnectable") Is IJPlane Then lLocation = 2

    Dim oMiddleConnectable As Object: Set oMiddleConnectable = pGeometricConstruction.Input("MiddleConnectable")
    Dim oLowerConnectable As Object: Set oLowerConnectable = pGeometricConstruction.Input("LowerConnectable")
    Dim oUpperConnectable As Object: Set oUpperConnectable = pGeometricConstruction.Input("UpperConnectable")
'    If pGeometricConstruction.Parameter("BoundaryCount") = 2 Then
'        Set oUpperConnectable = pGeometricConstruction.Input("UpperConnectable")
'    Else
'        Set oUpperConnectable = oLowerConnectable
'    End If


    
    Dim pDesignChild As IJDesignChild
    Dim bIsLowerConnectableStiffeningMiddleConnectable As Boolean: Let bIsLowerConnectableStiffeningMiddleConnectable = False
    If TypeOf oLowerConnectable Is IJStiffenerSystem Then
        Set pDesignChild = oLowerConnectable
        If pDesignChild.GetParent Is oMiddleConnectable Then Let bIsLowerConnectableStiffeningMiddleConnectable = True
    End If
    
    Dim bIsUpperConnectableStiffeningMiddleConnectable As Boolean: Let bIsUpperConnectableStiffeningMiddleConnectable = False
    If TypeOf oUpperConnectable Is IJStiffenerSystem Then
        Set pDesignChild = oUpperConnectable
        If pDesignChild.GetParent Is oMiddleConnectable Then Let bIsUpperConnectableStiffeningMiddleConnectable = True
    End If
    
    Dim oCurve1 As Object
    If bIsLowerConnectableStiffeningMiddleConnectable Then
        Dim oAxisPortExtractor21 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor21 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0001-AxisPortExtractor2")
        oAxisPortExtractor21.Inputs("Connectable").Add oLowerConnectable, "1"
        oAxisPortExtractor21.Parameter("GeometrySelector") = 2
        oAxisPortExtractor21.Evaluate
        pGeometricConstruction.ControlledInputs("Port1").Clear
        pGeometricConstruction.ControlledInputs("Port1").Add oAxisPortExtractor21.ControlledInputs("Port").Item(1)
        
        Set oCurve1 = oAxisPortExtractor21
    Else
        Dim oCurveInPlaneFromBO1 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveInPlaneFromBO1 = oGCFactory.CreateEntity("CurveInPlaneFromBO", pPOM, "0001-CurveInPlaneFromBO")
        oCurveInPlaneFromBO1.Inputs("Connectable").Add oLowerConnectable, "1"
        oCurveInPlaneFromBO1.Inputs("Plane").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveInPlaneFromBO1.Evaluate
        pGeometricConstruction.ControlledInputs("Port4").Clear
        If oCurveInPlaneFromBO1.ControlledInputs("Port").Count > 0 Then
            pGeometricConstruction.ControlledInputs("Port4").Add oCurveInPlaneFromBO1.ControlledInputs("Port").Item(1)
        End If

        Set oCurve1 = oCurveInPlaneFromBO1.Output("Curve")
    End If
    
    Dim oCurve2 As Object
    If bIsUpperConnectableStiffeningMiddleConnectable Then
        Dim oAxisPortExtractor22 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor22 = oGCFactory.CreateEntity("AxisPortExtractor2", pPOM, "0002-AxisPortExtractor2")
        oAxisPortExtractor22.Inputs("Connectable").Add oUpperConnectable, "2"
        oAxisPortExtractor22.Parameter("GeometrySelector") = 2
        oAxisPortExtractor22.Evaluate
        pGeometricConstruction.ControlledInputs("Port2").Clear
        pGeometricConstruction.ControlledInputs("Port2").Add oAxisPortExtractor22.ControlledInputs("Port").Item(1)
        
        Set oCurve2 = oAxisPortExtractor22
    Else
        Dim oCurveInPlaneFromBO2 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveInPlaneFromBO2 = oGCFactory.CreateEntity("CurveInPlaneFromBO", pPOM, "0002-CurveInPlaneFromBO")
        oCurveInPlaneFromBO2.Inputs("Connectable").Add oUpperConnectable, "2"
        oCurveInPlaneFromBO2.Inputs("Plane").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveInPlaneFromBO2.Evaluate
        pGeometricConstruction.ControlledInputs("Port5").Clear
        If oCurveInPlaneFromBO2.ControlledInputs("Port").Count > 0 Then
            pGeometricConstruction.ControlledInputs("Port5").Add oCurveInPlaneFromBO2.ControlledInputs("Port").Item(1)
        End If
        
        Set oCurve2 = oCurveInPlaneFromBO2.Output("Curve")
    End If
    
    Dim oFacePortExtractor03 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor03 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0003-FacePortExtractor0")
    oFacePortExtractor03.Inputs("Connectable").Add pGeometricConstruction.Inputs("MiddleConnectable").Item(1), "1"
    oFacePortExtractor03.Parameter("Offset") = 0#
    oFacePortExtractor03.Parameter("GeometrySelector") = 4 ' 2
    oFacePortExtractor03.Evaluate
    pGeometricConstruction.ControlledInputs("Port3").Clear
    pGeometricConstruction.ControlledInputs("Port3").Add oFacePortExtractor03.ControlledInputs("Port").Item("1")
    
    Dim oCurve0 As Object
    If True Then
        Dim oCurveByIntersection4 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection4 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0004-CurveByIntersection")
        oCurveByIntersection4.Inputs("Surface1").Add oFacePortExtractor03, "1"
        oCurveByIntersection4.Inputs("Surface2").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
        oCurveByIntersection4.Parameter("TrackFlag") = 1
        oCurveByIntersection4.Parameter("ConstructionSurface") = 0
        oCurveByIntersection4.Evaluate

        Set oCurve0 = oCurveByIntersection4
    End If
    
'    Dim oPointsAsRangeBox4a As SP3DGeometricConstruction.GeometricConstruction
'    Set oPointsAsRangeBox4a = oGCFactory.CreateEntity("PointsAsRangeBox", pPOM, "0004a-PointsAsRangeBox")
'    oPointsAsRangeBox4a.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
'    oPointsAsRangeBox4a.Inputs("Geometries").Add oCurve0, "1"
'    oPointsAsRangeBox4a.Evaluate
'
'    Dim oParamDistMeasureParallel4b As SP3DGeometricConstruction.GeometricConstruction
'    Set oParamDistMeasureParallel4b = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0004b-ParamDistMeasureParallel")
'    oParamDistMeasureParallel4b.Inputs("Graphics").Add oPointsAsRangeBox4a.Output("Points", "Low")
'    oParamDistMeasureParallel4b.Inputs("Graphics").Add oPointsAsRangeBox4a.Output("Points", "High")
'    oParamDistMeasureParallel4b.Evaluate

    
'    Dim oLineByPoints4b As SP3DGeometricConstruction.GeometricConstruction
'    Set oLineByPoints4b = oGCFactory.CreateEntity("LineByPoints", pPOM, "0004b-LineByPoints")
'    oLineByPoints4b.Inputs("StartPoint").Add oPointsAsRangeBox4a.Output("Points", "Low"), "1"
'    oLineByPoints4b.Inputs("EndPoint").Add oPointsAsRangeBox4a.Output("Points", "High"), "1"
'    oLineByPoints4b.Evaluate
'
'    Dim oPointAtCurveMiddle4c As SP3DGeometricConstruction.GeometricConstruction
'    Set oPointAtCurveMiddle4c = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0004c-PointAtCurveMiddle")
'    oPointAtCurveMiddle4c.Inputs("Curve").Add oLineByPoints4b, "1"
'    oPointAtCurveMiddle4c.Evaluate
   
    Dim oPointByCurves5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves5 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0005-PointByCurves")
    oPointByCurves5.Inputs("Curve1").Add oCurve0, "1"
    oPointByCurves5.Inputs("Curve2").Add oCurve1, "1"
    oPointByCurves5.Inputs("Surface").Add oFacePortExtractor03, "1"
    oPointByCurves5.Parameter("TrackFlag") = 1
    oPointByCurves5.Evaluate

    Dim oVectorTangentToCurve6 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter("BoundaryCount") = 2 Then
'        Dim oVectorNormalToSurface6 As SP3DGeometricConstruction.GeometricConstruction
'        Set oVectorNormalToSurface6 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "0006-VectorNormalToSurface")
'        oVectorNormalToSurface6.Inputs("Surface").Add oFacePortExtractor03, "1"
'        oVectorNormalToSurface6.Inputs("Point").Add oPointByCurves5, "1"
'        oVectorNormalToSurface6.Parameter("Range") = 10#
'        oVectorNormalToSurface6.Parameter("Orientation") = 1
'        oVectorNormalToSurface6.Parameter("TrackFlag") = 1
'        oVectorNormalToSurface6.Evaluate
        Set oVectorTangentToCurve6 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0006-VectorTangentToCurve")
        oVectorTangentToCurve6.Inputs("Curve").Add oCurve1, "1"
        oVectorTangentToCurve6.Inputs("Point").Add oPointByCurves5, "1"
        oVectorTangentToCurve6.Parameter("Range") = 10#
        oVectorTangentToCurve6.Parameter("Orientation") = 1
        oVectorTangentToCurve6.Parameter("TrackFlag") = 1
        oVectorTangentToCurve6.Evaluate
    Else
        Set oVectorTangentToCurve6 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0006-VectorTangentToCurve")
        oVectorTangentToCurve6.Inputs("Curve").Add oCurve0, "1"
        oVectorTangentToCurve6.Inputs("Point").Add oPointByCurves5, "1"
        oVectorTangentToCurve6.Parameter("Range") = 10#
        oVectorTangentToCurve6.Parameter("Orientation") = 1
        oVectorTangentToCurve6.Parameter("TrackFlag") = 1
        oVectorTangentToCurve6.Evaluate
    End If

    Dim oCSByPlane7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane7 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0007-CSByPlane")
    oCSByPlane7.Inputs("Plane").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oCSByPlane7.Inputs("Point").Add oPointByCurves5, "1"
    oCSByPlane7.Evaluate

    Dim oLineFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS8 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0008-LineFromCS")
    oLineFromCS8.Inputs("CoordinateSystem").Add oCSByPlane7, "1"
    oLineFromCS8.Parameter("LookingAxis") = 3
    oLineFromCS8.Parameter("Length") = 1#
    oLineFromCS8.Parameter("CSOrientation") = 1
    oLineFromCS8.Parameter("LineJustification") = 1
    oLineFromCS8.Evaluate

    Dim oCurveByProjection9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection9 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0009-CurveByProjection")
    oCurveByProjection9.Inputs("Curve").Add oVectorTangentToCurve6, "1"
    oCurveByProjection9.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "2"
    oCurveByProjection9.Inputs("Line").Add oLineFromCS8, "1"
    oCurveByProjection9.Parameter("TrackFlag") = 1
    oCurveByProjection9.Parameter("ConstructionSurface") = 0
    oCurveByProjection9.Evaluate

    Dim oLineOrientate10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineOrientate10 = oGCFactory.CreateEntity("LineOrientate", pPOM, "0010-LineOrientate")
    oLineOrientate10.Inputs("Line").Add oCurveByProjection9, "1"
    oLineOrientate10.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    If pGeometricConstruction.Parameter("BoundaryCount") = 2 Then
        oLineOrientate10.Parameter("Direction") = CLng(pGeometricConstruction.Parameter("Side"))
    Else
        oLineOrientate10.Parameter("Direction") = CLng(pGeometricConstruction.Parameter("Side"))
    End If
    oLineOrientate10.Evaluate

    Dim oPointAlongCurve11a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0011a-PointAlongCurve")
    oPointAlongCurve11a.Inputs("Curve").Add oLineOrientate10, "1"
    'oPointAlongCurve11a.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oPointAlongCurve11a.Parameter("Distance") = 1000 'oParamDistMeasureParallel4b.Parameter("Value")
    oPointAlongCurve11a.Parameter("TrackFlag") = 1
    oPointAlongCurve11a.Evaluate

    Dim oPointAlongCurve11b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0011b-PointAlongCurve")
    oPointAlongCurve11b.Inputs("Curve").Add oLineOrientate10, "1"
    'oPointAlongCurve11b.Inputs("Surface").Add pGeometricConstruction.Inputs("SketchingPlane").Item(1), "1"
    oPointAlongCurve11b.Parameter("Distance") = 1000 'oParamDistMeasureParallel4b.Parameter("Value")
    WebFrameLocation = CLng(pGeometricConstruction.Parameter("Location"))
    If pGeometricConstruction.Parameter("BoundaryCount") = 2 Then
        oPointAlongCurve11b.Parameter("TrackFlag") = 3 - WebFrameLocation
    Else
        oPointAlongCurve11b.Parameter("TrackFlag") = 3 - WebFrameLocation
    End If
    oPointAlongCurve11b.Evaluate

    Dim oPointByCurves12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves12 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0012-PointByCurves")
    oPointByCurves12.Inputs("Curve1").Add oCurve0, "1"
    oPointByCurves12.Inputs("Curve2").Add oCurve1, "1"
    oPointByCurves12.Inputs("TrackPoint").Add oPointAlongCurve11a, "1"
    oPointByCurves12.Parameter("TrackFlag") = 1
    oPointByCurves12.Evaluate

    Dim iTrackFlag As Integer
    If oUpperConnectable Is oLowerConnectable Then iTrackFlag = 2 Else iTrackFlag = 1
    Dim oPointByCurves13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves13 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0013-PointByCurves")
    oPointByCurves13.Inputs("Curve1").Add oCurve0, "1"
    oPointByCurves13.Inputs("Curve2").Add oCurve2, "1"
    oPointByCurves13.Inputs("TrackPoint").Add oPointAlongCurve11a, "1"
    oPointByCurves13.Parameter("TrackFlag") = iTrackFlag
    oPointByCurves13.Evaluate
    
    Dim oLineByPoints14a As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints14a = oGCFactory.CreateEntity("LineByPoints", pPOM, "0014a-LineByPoints")
    oLineByPoints14a.Inputs("StartPoint").Add oPointByCurves12, "1"
    oLineByPoints14a.Inputs("EndPoint").Add oPointByCurves13, "1"
    oLineByPoints14a.Evaluate
    
    Dim oLineOrientate14b As SP3DGeometricConstruction.GeometricConstruction
    Set oLineOrientate14b = oGCFactory.CreateEntity("LineOrientate", pPOM, "0014b-LineOrientate")
    oLineOrientate14b.Inputs("Line").Add oLineByPoints14a, "1"
    oLineOrientate14b.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oLineOrientate14b.Parameter("Direction") = 1
    oLineOrientate14b.Evaluate
    
    Dim oPointAtCurveEnd14c As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd14c = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0014c-PointAtCurveEnd")
    oPointAtCurveEnd14c.Inputs("Curve").Add oLineOrientate14b, "1"
    oPointAtCurveEnd14c.Evaluate
    
    Dim oPointAtCurveMiddle15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle15 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0015-PointAtCurveMiddle")
    oPointAtCurveMiddle15.Inputs("Curve").Add oLineByPoints14a, "1"
    oPointAtCurveMiddle15.Evaluate
    
    Dim oCSByPoints16a As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints16a = oGCFactory.CreateEntity("CSByPoints", pPOM, "0016a-CSByPoints")
    oCSByPoints16a.Inputs("OriginPoint").Add oPointAtCurveMiddle15, "1"
    oCSByPoints16a.Inputs("AxisPoint1").Add oPointAtCurveEnd14c, "1"
    If Not oUpperConnectable Is oLowerConnectable Then
        oCSByPoints16a.Inputs("AxisPoint2").Add oPointAlongCurve11b, "1"
    Else
        oCSByPoints16a.Inputs("AxisPoint2").Add oPointAlongCurve11a, "1"
    End If
    oCSByPoints16a.Parameter("AxesRoles") = 5
    oCSByPoints16a.Parameter("CSOrientation") = 1
    oCSByPoints16a.Evaluate
    
    Dim oLineFromCS16b As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS16b = oGCFactory.CreateEntity("LineFromCS", pPOM, "0016b-LineFromCS")
    oLineFromCS16b.Inputs("CoordinateSystem").Add oCSByPoints16a, "1"
    oLineFromCS16b.Parameter("LookingAxis") = 2
    oLineFromCS16b.Parameter("Length") = -100#
    oLineFromCS16b.Parameter("CSOrientation") = 1
    oLineFromCS16b.Parameter("LineJustification") = 1
    oLineFromCS16b.Evaluate
   
    Dim oPointByCurves16c As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves16c = oGCFactory.CreateEntity("PointByCurves", pPOM, "0016c-PointByCurves")
    oPointByCurves16c.Inputs("Curve1").Add oCurve0, "1"
    oPointByCurves16c.Inputs("Curve2").Add oLineFromCS16b, "1"
    oPointByCurves16c.Inputs("TrackPoint").Add oPointAtCurveMiddle15, "1"
    oPointByCurves16c.Parameter("TrackFlag") = 1
    oPointByCurves16c.Evaluate
    
    Dim oPointAlongCurve16d As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve16d = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0016d-PointAlongCurve")
    oPointAlongCurve16d.Inputs("Curve").Add oLineFromCS16b, "1"
    oPointAlongCurve16d.Inputs("Point").Add oPointByCurves16c, "1"
    oPointAlongCurve16d.Parameter("Distance") = 1#
    oPointAlongCurve16d.Evaluate

    Dim oCSByCS16e  As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS16e = oGCFactory.CreateEntity("CSByCS", pPOM, "0016e-CSByCS")
    oCSByCS16e.Inputs("CoordinateSystem").Add oCSByPoints16a, "1"
    oCSByCS16e.Inputs("Origin").Add oPointAlongCurve16d, "1"
    oCSByCS16e.Evaluate

    Dim oLowerBoundary As Object
    If bIsLowerConnectableStiffeningMiddleConnectable Then
        Dim oSphereByCenterRadius17 As SP3DGeometricConstruction.GeometricConstruction
        Set oSphereByCenterRadius17 = oGCFactory.CreateEntity("SphereByCenterRadius", pPOM, "0017-SphereByCenterRadius")
        oSphereByCenterRadius17.Inputs("Center").Add oPointByCurves12
        oSphereByCenterRadius17.Parameter("Radius") = 1
        oSphereByCenterRadius17.Evaluate
        
        Dim oSurfFromGType17 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType17 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0017-SurfFromGType")
        oSurfFromGType17.Inputs("Surface").Add oSphereByCenterRadius17
        oSurfFromGType17.Evaluate

        Set oLowerBoundary = oSurfFromGType17
    Else
        Dim oSurfByNonLinearExtrusion17 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByNonLinearExtrusion17 = oGCFactory.CreateEntity("SurfByNonLinearExtrusion", pPOM, "0017-SurfByNonLinearExtrusion")
        oSurfByNonLinearExtrusion17.Inputs("PlanarCrossSection").Add oCurve1, "1"
        oSurfByNonLinearExtrusion17.Inputs("PlanarExtrusionPath").Add oLineFromCS8, "1"
        oSurfByNonLinearExtrusion17.Evaluate
        Set oLowerBoundary = oSurfByNonLinearExtrusion17
    End If

    Dim oUpperBoundary As Object
    If bIsUpperConnectableStiffeningMiddleConnectable Then
        Dim oSphereByCenterRadius18 As SP3DGeometricConstruction.GeometricConstruction
        Set oSphereByCenterRadius18 = oGCFactory.CreateEntity("SphereByCenterRadius", pPOM, "0018-SphereByCenterRadius")
        oSphereByCenterRadius18.Inputs("Center").Add oPointByCurves13
        oSphereByCenterRadius18.Parameter("Radius") = 1
        oSphereByCenterRadius18.Evaluate
        
        Dim oSurfFromGType18 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType18 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0018-SurfFromGType")
        oSurfFromGType18.Inputs("Surface").Add oSphereByCenterRadius18
        oSurfFromGType18.Evaluate

        Set oUpperBoundary = oSurfFromGType18
    Else
        Dim oSurfByNonLinearExtrusion18 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByNonLinearExtrusion18 = oGCFactory.CreateEntity("SurfByNonLinearExtrusion", pPOM, "0018-SurfByNonLinearExtrusion")
        oSurfByNonLinearExtrusion18.Inputs("PlanarCrossSection").Add oCurve2, "1"
        oSurfByNonLinearExtrusion18.Inputs("PlanarExtrusionPath").Add oLineFromCS8, "1"
        oSurfByNonLinearExtrusion18.Evaluate
        Set oUpperBoundary = oSurfByNonLinearExtrusion18
    End If
    
    Dim oParamDistMeasureParallel19 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel19 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0019-ParamDistMeasureParallel")
    oParamDistMeasureParallel19.Inputs("Graphics").Add oPointByCurves12
    oParamDistMeasureParallel19.Inputs("Graphics").Add oPointByCurves13
    oParamDistMeasureParallel19.Evaluate
        
    Dim oPlaneFromCS20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS20 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "0020-PlaneFromCS")
    oPlaneFromCS20.Inputs("CoordinateSystem").Add oCSByCS16e
    oPlaneFromCS20.Parameter("LookingAxis") = 1
    oPlaneFromCS20.Parameter("Range") = oParamDistMeasureParallel19.Parameter("Value")
    oPlaneFromCS20.Evaluate
        
    Dim oSurfFromGType21 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType21 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0021-SurfFromGType")
    oSurfFromGType21.Inputs("Surface").Add oPlaneFromCS20
    oSurfFromGType21.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Points", "Middle") = oPointAlongCurve16d.Output
    oGCMacro.Output("Points", "Lower") = oPointByCurves12.Output
    oGCMacro.Output("Points", "Upper") = oPointByCurves13.Output
    oGCMacro.Output("CoordinateSystem") = oCSByCS16e.Output
    oGCMacro.Output("Support", 1) = oSurfFromGType21.Output
    oGCMacro.Output("Curves", "Lower") = oCurve1
    oGCMacro.Output("Curves", "Upper") = oCurve2
    oGCMacro.Output("Boundaries", "MiddleConnectable") = oFacePortExtractor03.Output
    oGCMacro.Output("Boundaries", "LowerConnectable") = oLowerBoundary 'oSurfByNonLinearExtrusion17.Output
    oGCMacro.Output("Boundaries", "UpperConnectable") = oUpperBoundary 'oSurfByNonLinearExtrusion18.Output
    oGCMacro.Output("OrientationLine", 1) = oLineFromCS16b.Output
End Sub
