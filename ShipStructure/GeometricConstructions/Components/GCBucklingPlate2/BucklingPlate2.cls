VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate1", "Select first TransitionPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate2", "Select second TransitionPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("ContinuityPlate1", "Select first ContinuityPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("ContinuityPlate2", "Select second ContinuityPlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate1", "Select first SidePlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate2", "Select second SidePlate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfContinuityPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfContinuityPlate2")
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
    oFacePortExtractor3.Inputs("Connectable").Add pGeometricConstruction.Inputs("ContinuityPlate2").Item(1)
    oFacePortExtractor3.Parameter("GeometrySelector") = 2
    oFacePortExtractor3.Parameter("FacesContext") = 1
    oFacePortExtractor3.Parameter("LookingAxis") = 1
    oFacePortExtractor3.Parameter("TrackFlag") = 1
    oFacePortExtractor3.Parameter("Offset") = 0
    oFacePortExtractor3.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate2").Clear
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate2").AddElements oFacePortExtractor3.ControlledInputs("Port")

    Dim oFacePortExtractor4 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor4 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "004-FacePortExtractor")
    oFacePortExtractor4.Inputs("Connectable").Add pGeometricConstruction.Inputs("ContinuityPlate1").Item(1)
    oFacePortExtractor4.Parameter("GeometrySelector") = 2
    oFacePortExtractor4.Parameter("FacesContext") = 1
    oFacePortExtractor4.Parameter("LookingAxis") = 1
    oFacePortExtractor4.Parameter("TrackFlag") = 1
    oFacePortExtractor4.Parameter("Offset") = 0
    oFacePortExtractor4.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").Clear
    pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").AddElements oFacePortExtractor4.ControlledInputs("Port")

    Dim oFacePortExtractor5 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        Set oFacePortExtractor5 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "005-FacePortExtractor")
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
        Set oFacePortExtractor6 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "006-FacePortExtractor")
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

    Dim oCurveByIntersection5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection5 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "007-CurveByIntersection")
    oCurveByIntersection5.Inputs("Surface1").Add oFacePortExtractor4
    oCurveByIntersection5.Inputs("Surface2").Add oFacePortExtractor2
    oCurveByIntersection5.Parameter("TrackFlag") = 1
    oCurveByIntersection5.Evaluate

    Dim oPointAtCurveStart6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart6 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "008-PointAtCurveStart")
    oPointAtCurveStart6.Inputs("Curve").Add oCurveByIntersection5
    oPointAtCurveStart6.Evaluate

    Dim oPointAtCurveMiddle7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle7 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "009-PointAtCurveMiddle")
    oPointAtCurveMiddle7.Inputs("Curve").Add oCurveByIntersection5
    oPointAtCurveMiddle7.Evaluate

    Dim oCSByPlane8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane8 = oGCFactory.CreateEntity("CSByPlane", pPOM, "010-CSByPlane")
    oCSByPlane8.Inputs("Plane").Add oFacePortExtractor2
    oCSByPlane8.Inputs("Point").Add oPointAtCurveStart6
    oCSByPlane8.Evaluate

    Dim oLineFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS9 = oGCFactory.CreateEntity("LineFromCS", pPOM, "011-LineFromCS")
    oLineFromCS9.Inputs("CoordinateSystem").Add oCSByPlane8
    oLineFromCS9.Parameter("LookingAxis") = 3
    oLineFromCS9.Parameter("Length") = 1#
    oLineFromCS9.Evaluate

    Dim oCSByLines10 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines10 = oGCFactory.CreateEntity("CSByLines", pPOM, "012-CSByLines")
    oCSByLines10.Inputs("AxisLine1").Add oCurveByIntersection5
    oCSByLines10.Inputs("AxisLine2").Add oLineFromCS9
    oCSByLines10.Parameter("AxesRoles") = 1
    oCSByLines10.Parameter("CSOrientation") = 1
    oCSByLines10.Evaluate

    Dim oCSByLines11 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines11 = oGCFactory.CreateEntity("CSByLines", pPOM, "013-CSByLines")
    oCSByLines11.Inputs("AxisLine1").Add oLineFromCS9
    oCSByLines11.Inputs("AxisLine2").Add oCurveByIntersection5
    oCSByLines11.Parameter("AxesRoles") = 3
    oCSByLines11.Parameter("CSOrientation") = 1
    oCSByLines11.Evaluate

    Dim oCurveByIntersection12 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection12 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "014-CurveByIntersection")
    oCurveByIntersection12.Inputs("Surface1").Add oFacePortExtractor1
    oCurveByIntersection12.Inputs("Surface2").Add oFacePortExtractor3
    oCurveByIntersection12.Inputs("CoordinateSystem").Add oCSByLines11
    oCurveByIntersection12.Parameter("TrackFlag") = 1
    oCurveByIntersection12.Evaluate

    Dim oExtRuledSurface13 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface13 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "015-ExtRuledSurface")
    oExtRuledSurface13.Inputs("Curve1").Add oCurveByIntersection5
    oExtRuledSurface13.Inputs("Curve2").Add oCurveByIntersection12
    oExtRuledSurface13.Parameter("Extension") = 0.1
    oExtRuledSurface13.Evaluate

    Dim oLateralBoundaries14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries14 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "016-LateralBoundaries")
    oLateralBoundaries14.Inputs("CoordinateSystem1").Add oCSByLines10
    oLateralBoundaries14.Inputs("PointS1").Add oExtRuledSurface13.Output("PointS1", 1)
    oLateralBoundaries14.Inputs("PointE1").Add oExtRuledSurface13.Output("PointS2", 1)
    oLateralBoundaries14.Inputs("CoordinateSystem2").Add oCSByLines10
    oLateralBoundaries14.Inputs("PointS2").Add oExtRuledSurface13.Output("PointE1", 1)
    oLateralBoundaries14.Inputs("PointE2").Add oExtRuledSurface13.Output("PointE2", 1)
    oLateralBoundaries14.Parameter("Extension") = 0.1
    oLateralBoundaries14.Evaluate

    Dim oCSByPlane15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane15 = oGCFactory.CreateEntity("CSByPlane", pPOM, "017-CSByPlane")
    oCSByPlane15.Inputs("Plane").Add oExtRuledSurface13.Output("Support", 1)
    oCSByPlane15.Inputs("Point").Add oPointAtCurveMiddle7
    oCSByPlane15.Evaluate

    Dim oLateralBoundaries16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries16 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "018-LateralBoundaries")
    oLateralBoundaries16.Inputs("CoordinateSystem1").Add oCSByPlane15
    oLateralBoundaries16.Inputs("PointS1").Add oExtRuledSurface13.Output("PointS1", 1)
    oLateralBoundaries16.Inputs("PointE1").Add oExtRuledSurface13.Output("PointE1", 1)
    oLateralBoundaries16.Inputs("CoordinateSystem2").Add oCSByPlane15
    oLateralBoundaries16.Inputs("PointS2").Add oExtRuledSurface13.Output("PointS2", 1)
    oLateralBoundaries16.Inputs("PointE2").Add oExtRuledSurface13.Output("PointE2", 1)
    oLateralBoundaries16.Parameter("Extension") = 0.1
    oLateralBoundaries16.Evaluate

    ' orientate the support as the support of the first ContinuityPlate
    Call SurfaceBody_OrientateFollowingVector(oExtRuledSurface13.Output("Support", 1), Vector_FromPlane(oFacePortExtractor4), True)

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oExtRuledSurface13.Output("Support", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate1") = oFacePortExtractor2 'oLateralBoundaries14.Output("Boundary", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate2") = oFacePortExtractor1 'oLateralBoundaries14.Output("Boundary", 2)
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oFacePortExtractor5 'oLateralBoundaries16.Output("Boundary", 1)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oLateralBoundaries16.Output("Boundary", 2)
    End If
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oFacePortExtractor6 'oLateralBoundaries16.Output("Boundary", 2)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oLateralBoundaries16.Output("Boundary", 1)
    End If

    ' populate output "CoordinateSystem"
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstruction, Nothing, GetGCGeomFactory(), GetGCGeomFactory2(), Position_FromPoint(oPointAtCurveMiddle7))
End Sub
