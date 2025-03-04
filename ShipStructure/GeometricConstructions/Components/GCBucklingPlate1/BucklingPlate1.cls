VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlate1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate1", "Select the first TransitionPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate2", "Select the second TransitionPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("ContinuityPlate", "Select the incoming ContinuityPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate1", "Select first SidePlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate2", "Select second SidePlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfContinuityPlate")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate2")
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oFacePortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor1 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "001-FacePortExtractor")
    oFacePortExtractor1.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
    oFacePortExtractor1.Parameter("GeometrySelector") = 2
    oFacePortExtractor1.Parameter("FacesContext") = 1
    oFacePortExtractor1.Parameter("LookingAxis") = 1
    oFacePortExtractor1.Parameter("TrackFlag") = 1
    oFacePortExtractor1.Parameter("Offset") = 0
    oFacePortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").Clear
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").AddElements oFacePortExtractor1.ControlledInputs("Port")

    Dim oFacePortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor2 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "002-FacePortExtractor")
    oFacePortExtractor2.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
    oFacePortExtractor2.Parameter("GeometrySelector") = 2
    oFacePortExtractor2.Parameter("FacesContext") = 1
    oFacePortExtractor2.Parameter("LookingAxis") = 1
    oFacePortExtractor2.Parameter("TrackFlag") = 1
    oFacePortExtractor2.Parameter("Offset") = 0
    oFacePortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").Clear
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").AddElements oFacePortExtractor2.ControlledInputs("Port")

    Dim oFacePortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor3 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "003-FacePortExtractor")
    oFacePortExtractor3.Inputs("Connectable").Add pGeometricConstruction.Inputs("ContinuityPlate").Item(1)
    oFacePortExtractor3.Parameter("GeometrySelector") = 2
    oFacePortExtractor3.Parameter("FacesContext") = 1
    oFacePortExtractor3.Parameter("LookingAxis") = 1
    oFacePortExtractor3.Parameter("TrackFlag") = 1
    oFacePortExtractor3.Parameter("Offset") = 0
    oFacePortExtractor3.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate").Clear
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate").AddElements oFacePortExtractor3.ControlledInputs("Port")

    Dim oFacePortExtractor5 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        Set oFacePortExtractor5 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "004-FacePortExtractor")
        oFacePortExtractor5.Inputs("Connectable").Add pGeometricConstruction.Inputs("SidePlate1").Item(1)
        oFacePortExtractor5.Parameter("GeometrySelector") = 2
        oFacePortExtractor5.Parameter("FacesContext") = 1
        oFacePortExtractor5.Parameter("LookingAxis") = 1
        oFacePortExtractor5.Parameter("TrackFlag") = 1
        oFacePortExtractor5.Parameter("Offset") = 0
        oFacePortExtractor5.Evaluate
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").AddElements oFacePortExtractor5.ControlledInputs("Port")
    End If
    
    Dim oFacePortExtractor6 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        Set oFacePortExtractor6 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "005-FacePortExtractor")
        oFacePortExtractor6.Inputs("Connectable").Add pGeometricConstruction.Inputs("SidePlate2").Item(1)
        oFacePortExtractor6.Parameter("GeometrySelector") = 2
        oFacePortExtractor6.Parameter("FacesContext") = 1
        oFacePortExtractor6.Parameter("LookingAxis") = 1
        oFacePortExtractor6.Parameter("TrackFlag") = 1
        oFacePortExtractor6.Parameter("Offset") = 0
        oFacePortExtractor6.Evaluate
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").AddElements oFacePortExtractor6.ControlledInputs("Port")
    End If
    Dim oCurveByIntersection4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection4 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "006-CurveByIntersection")
    oCurveByIntersection4.Inputs("Surface1").Add oFacePortExtractor3
    oCurveByIntersection4.Inputs("Surface2").Add oFacePortExtractor2
    oCurveByIntersection4.Parameter("TrackFlag") = 1
    oCurveByIntersection4.Evaluate

    Dim oParamDistMeasureLength5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength5 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "007-ParamDistMeasureLength")
    oParamDistMeasureLength5.Inputs("Curve").Add oCurveByIntersection4
    oParamDistMeasureLength5.Parameter("Value") = 0
    oParamDistMeasureLength5.Evaluate

    Dim oPointAtCurveMiddle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle6 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "008-PointAtCurveMiddle")
    oPointAtCurveMiddle6.Inputs("Curve").Add oCurveByIntersection4
    oPointAtCurveMiddle6.Evaluate

    Dim oPointAtCurveEnd7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd7 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "009-PointAtCurveEnd")
    oPointAtCurveEnd7.Inputs("Curve").Add oCurveByIntersection4
    oPointAtCurveEnd7.Evaluate

    Dim oPointAtCurveStart8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart8 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "010-PointAtCurveStart")
    oPointAtCurveStart8.Inputs("Curve").Add oCurveByIntersection4
    oPointAtCurveStart8.Evaluate

    Dim oPointAlongCurve9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve9 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "011-PointAlongCurve")
    oPointAlongCurve9.Inputs("Curve").Add oCurveByIntersection4
    oPointAlongCurve9.Inputs("Point").Add oPointAtCurveStart8
    oPointAlongCurve9.Inputs("TrackPoint").Add oPointAtCurveMiddle6
    oPointAlongCurve9.Parameter("Distance") = oParamDistMeasureLength5
    oPointAlongCurve9.Parameter("TrackFlag") = 2
    oPointAlongCurve9.Evaluate

    Dim oPointAlongCurve10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve10 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "012-PointAlongCurve")
    oPointAlongCurve10.Inputs("Curve").Add oCurveByIntersection4
    oPointAlongCurve10.Inputs("Point").Add oPointAtCurveEnd7
    oPointAlongCurve10.Inputs("TrackPoint").Add oPointAtCurveMiddle6
    oPointAlongCurve10.Parameter("Distance") = oParamDistMeasureLength5
    oPointAlongCurve10.Parameter("TrackFlag") = 2
    oPointAlongCurve10.Evaluate

    Dim oCSByPlane11 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane11 = oGCFactory.CreateEntity("CSByPlane", pPOM, "013-CSByPlane")
    oCSByPlane11.Inputs("Plane").Add oFacePortExtractor2
    oCSByPlane11.Inputs("Point").Add oPointAtCurveStart8
    oCSByPlane11.Evaluate

    Dim oLineByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints12 = oGCFactory.CreateEntity("LineByPoints", pPOM, "014-LineByPoints")
'''    oLineByPoints12.Inputs("StartPoint").Add oPointAlongCurve9
'''    oLineByPoints12.Inputs("EndPoint").Add oPointAlongCurve10
    oLineByPoints12.Inputs("StartPoint").Add oPointAtCurveStart8
    oLineByPoints12.Inputs("EndPoint").Add oPointAtCurveEnd7
    oLineByPoints12.Evaluate

    Dim oLineFromCS13 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS13 = oGCFactory.CreateEntity("LineFromCS", pPOM, "015-LineFromCS")
    oLineFromCS13.Inputs("CoordinateSystem").Add oCSByPlane11
    oLineFromCS13.Parameter("LookingAxis") = 3
    oLineFromCS13.Parameter("Length") = 1#
    oLineFromCS13.Evaluate

    Dim oCSByLines14 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines14 = oGCFactory.CreateEntity("CSByLines", pPOM, "016-CSByLines")
    oCSByLines14.Inputs("AxisLine1").Add oLineFromCS13
    oCSByLines14.Inputs("AxisLine2").Add oCurveByIntersection4
    oCSByLines14.Parameter("AxesRoles") = 3
    oCSByLines14.Parameter("CSOrientation") = 1
    oCSByLines14.Evaluate

    Dim oCurveByProjection15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection15 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "017-CurveByProjection")
    oCurveByProjection15.Inputs("Curve").Add oLineByPoints12
    oCurveByProjection15.Inputs("Surface").Add oFacePortExtractor1
    oCurveByProjection15.Inputs("Line").Add oLineFromCS13
    oCurveByProjection15.Inputs("CoordinateSystem").Add oCSByLines14
    oCurveByProjection15.Parameter("TrackFlag") = 1
    oCurveByProjection15.Evaluate

    Dim oExtRuledSurface16 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface16 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "018-ExtRuledSurface")
    oExtRuledSurface16.Inputs("Curve1").Add oCurveByIntersection4
    oExtRuledSurface16.Inputs("Curve2").Add oCurveByProjection15
    oExtRuledSurface16.Parameter("Extension") = 0.1
    oExtRuledSurface16.Evaluate

    Dim oCSByPlane17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane17 = oGCFactory.CreateEntity("CSByPlane", pPOM, "019-CSByPlane")
    oCSByPlane17.Inputs("Plane").Add oExtRuledSurface16.Output("Support", 1)
    oCSByPlane17.Inputs("Point").Add oPointAtCurveMiddle6
    oCSByPlane17.Evaluate

    Dim oLateralBoundaries18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries18 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "020-LateralBoundaries")
    oLateralBoundaries18.Inputs("CoordinateSystem1").Add oCSByPlane17
    oLateralBoundaries18.Inputs("PointS1").Add oPointAtCurveStart8
    oLateralBoundaries18.Inputs("PointE1").Add oExtRuledSurface16.Output("PointS2", 1)
    oLateralBoundaries18.Inputs("CoordinateSystem2").Add oCSByPlane17
    oLateralBoundaries18.Inputs("PointS2").Add oPointAtCurveEnd7
    oLateralBoundaries18.Inputs("PointE2").Add oExtRuledSurface16.Output("PointE2", 1)
    oLateralBoundaries18.Parameter("Extension") = 0.1
    oLateralBoundaries18.Evaluate

    Dim oLateralBoundaries19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries19 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "021-LateralBoundaries")
    oLateralBoundaries19.Inputs("CoordinateSystem1").Add oCSByPlane17
    oLateralBoundaries19.Inputs("PointS1").Add oPointAtCurveStart8
    oLateralBoundaries19.Inputs("PointE1").Add oPointAtCurveEnd7
    oLateralBoundaries19.Inputs("CoordinateSystem2").Add oCSByPlane17
    oLateralBoundaries19.Inputs("PointS2").Add oExtRuledSurface16.Output("PointS2", 1)
    oLateralBoundaries19.Inputs("PointE2").Add oExtRuledSurface16.Output("PointE2", 1)
    oLateralBoundaries19.Parameter("Extension") = 0.1
    oLateralBoundaries19.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    ' orientate the support as the support of the ContinuityPlate
    Call SurfaceBody_OrientateFollowingVector(oExtRuledSurface16.Output("Support", 1), Vector_FromPlane(oFacePortExtractor3), True)
    
    oGCMacro.Output("Support", 1) = oExtRuledSurface16.Output("Support", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate1") = oFacePortExtractor2 'oLateralBoundaries18.Output("Boundary", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate2") = oFacePortExtractor1 'oLateralBoundaries18.Output("Boundary", 2)
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oFacePortExtractor5 'oLateralBoundaries19.Output("Boundary", 1)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oLateralBoundaries19.Output("Boundary", 2)
    End If
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oFacePortExtractor6 'ooLateralBoundaries19.Output("Boundary", 2)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oLateralBoundaries19.Output("Boundary", 1)
    End If

    ' populate output "CoordinateSystem"
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstruction, Nothing, GetGCGeomFactory(), GetGCGeomFactory2(), Position_FromPoint(oPointAtCurveMiddle6))
End Sub
