VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlate0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Errors
    MISSING_MEMBER_PART = 1
    MISSING_TRANSITION_PLATE1 = 2
    MISSING_TRANSITION_PLATE2 = 3
    DUPLICATE_INPUTS = 4
End Enum

Implements IJGeometricConstructionDefinitionService
Implements IJGCMigrate
Implements IJGCToDoDelegate

Private Property Get Source() As String
    Let Source = "GCBucklingPlate0.BucklingPlate0"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
Private Sub Class_Initialize()
    Call DebugStart(Source)

End Sub
Private Sub Class_Terminate()

    Call DebugStop(Source)
End Sub
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    ' located input
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select incoming Member", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate1", "Select the first limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate2", "Select the second limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate1", "Select a first lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate2", "Select a second lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfMemberPart", "IJGeometry")
    
    ' parameters
    Call pGeometricConstructionDefinition.AddParameter("Distance", "Distance from Node", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)

    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER_PART, "MissingMemberPart", "Missing the member part")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE1, "MissingTransitionPlate1", "Missing the transition plate #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE2, "MissingTransitionPlate2", "Missing the transition plate #2")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_INPUTS, "DuplicateInputs", "2 inputs are the same")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call DebugOut
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Call DebugIn(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    On Error GoTo ErrorHandler
    
    Call Evaluate(pGeometricConstruction, pPOM, False)
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM, ByVal bIsAPSDeleting As Boolean)
    Call DebugIn(Method("Evaluate"))
    
    ' check inputs
    If pGeometricConstruction.Inputs("MemberPart").Count = 0 Then Err.Raise MISSING_MEMBER_PART
    If pGeometricConstruction.Inputs("TransitionPlate1").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE1
    If pGeometricConstruction.Inputs("TransitionPlate2").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE2
    If True Then
        Dim pElements As IJElements: Set pElements = New JObjectCollection
        Call pElements.Add(pGeometricConstruction.Input("MemberPart"))
        If pElements.Contains(pGeometricConstruction.Input("TransitionPlate1")) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input("TransitionPlate1"))
        If pElements.Contains(pGeometricConstruction.Input("TransitionPlate2")) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input("TransitionPlate2"))
        If pGeometricConstruction.Inputs("SidePlate1").Count > 0 Then
            If pElements.Contains(pGeometricConstruction.Input("SidePlate1")) Then Err.Raise DUPLICATE_INPUTS
            Call pElements.Add(pGeometricConstruction.Input("SidePlate1"))
        End If
        If pGeometricConstruction.Inputs("SidePlate2").Count > 0 Then
            If pElements.Contains(pGeometricConstruction.Input("SidePlate2")) Then Err.Raise DUPLICATE_INPUTS
        End If
    End If
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("Distance"))
    oParamDistConstant1.Evaluate

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfMemberPart").Clear
    pGeometricConstruction.ControlledInputs("PortOfMemberPart").AddElements oLineAxisPortExtractor2.ControlledInputs("Port")

    Dim oFacePortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor1 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "003-FacePortExtractor")
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
    Set oFacePortExtractor2 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "004-FacePortExtractor")
    oFacePortExtractor2.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
    oFacePortExtractor2.Parameter("GeometrySelector") = 2
    oFacePortExtractor2.Parameter("FacesContext") = 1
    oFacePortExtractor2.Parameter("LookingAxis") = 1
    oFacePortExtractor2.Parameter("TrackFlag") = 1
    oFacePortExtractor2.Parameter("Offset") = 0
    oFacePortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").Clear
    pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").AddElements oFacePortExtractor2.ControlledInputs("Port")

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
    
    Dim oPointAtCurveMiddle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle6 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "007-PointAtCurveMiddle")
    oPointAtCurveMiddle6.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveMiddle6.Evaluate

    Dim oCSFromMember8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember8 = oGCFactory.CreateEntity("CSFromMember", pPOM, "008-CSFromMember")
    oCSFromMember8.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember8.Inputs("Point").Add oPointAtCurveMiddle6
    oCSFromMember8.Evaluate

    Dim oEdgePortExtractor5a As SP3DGeometricConstruction.GeometricConstruction
    Set oEdgePortExtractor5a = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "009-EdgePortExtractor")
    oEdgePortExtractor5a.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
    oEdgePortExtractor5a.Inputs("FacePort").Add oFacePortExtractor2
    oEdgePortExtractor5a.Inputs("CoordinateSystem").Add oCSFromMember8
    oEdgePortExtractor5a.Parameter("GeometrySelector") = 2
'    oEdgePortExtractor5a.Parameter("FacesContext") = 1
    oEdgePortExtractor5a.Parameter("LookingAxis") = 1
    oEdgePortExtractor5a.Parameter("TrackFlag") = 1
    oEdgePortExtractor5a.Evaluate
        
    Dim oEdgePortExtractor5b As SP3DGeometricConstruction.GeometricConstruction
    Set oEdgePortExtractor5b = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "010-EdgePortExtractor")
    oEdgePortExtractor5b.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
    oEdgePortExtractor5b.Inputs("FacePort").Add oFacePortExtractor2
    oEdgePortExtractor5b.Inputs("CoordinateSystem").Add oCSFromMember8
    oEdgePortExtractor5b.Parameter("GeometrySelector") = 2
'    oEdgePortExtractor5b.Parameter("FacesContext") = 1
    oEdgePortExtractor5b.Parameter("LookingAxis") = 1
    oEdgePortExtractor5b.Parameter("TrackFlag") = 2
    oEdgePortExtractor5b.Evaluate
        
    Dim oPointAtCurveMiddle7a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle7a = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "011-PointAtCurveMiddle")
    oPointAtCurveMiddle7a.Inputs("Curve").Add oEdgePortExtractor5a
    oPointAtCurveMiddle7a.Evaluate
    
    Dim oPointAtCurveMiddle7b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle7b = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "012-PointAtCurveMiddle")
    oPointAtCurveMiddle7b.Inputs("Curve").Add oEdgePortExtractor5b
    oPointAtCurveMiddle7b.Evaluate
    
    Dim oLineByPoints7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints7 = oGCFactory.CreateEntity("LineByPoints", pPOM, "013-LineByPoints")
    oLineByPoints7.Input("StartPoint") = oPointAtCurveMiddle7a
    oLineByPoints7.Input("EndPoint") = oPointAtCurveMiddle7b
    oLineByPoints7.Evaluate
    
    Dim oPointAtCurveMiddle7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle7 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "014-PointAtCurveMiddle")
    oPointAtCurveMiddle7.Input("Curve") = oLineByPoints7
    oPointAtCurveMiddle7.Evaluate
    
    Dim oPointAtCurveExtremity9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity9 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "015-PointAtCurveExtremity")
    oPointAtCurveExtremity9.Inputs("Curve").Add oLineAxisPortExtractor2
    'oPointAtCurveExtremity9.Inputs("TrackPoint").Add oPointAtCurveMiddle7
    oPointAtCurveExtremity9.Inputs("TrackPoint").Add oPointAtCurveMiddle7a
    oPointAtCurveExtremity9.Parameter("TrackFlag") = 1
    oPointAtCurveExtremity9.Evaluate

    Dim oPointAtCurveExtremity10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveExtremity10 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "016-PointAtCurveExtremity")
    oPointAtCurveExtremity10.Inputs("Curve").Add oLineAxisPortExtractor2
    'oPointAtCurveExtremity10.Inputs("TrackPoint").Add oPointAtCurveMiddle7
    oPointAtCurveExtremity10.Inputs("TrackPoint").Add oPointAtCurveMiddle7a
    oPointAtCurveExtremity10.Parameter("TrackFlag") = 2
    oPointAtCurveExtremity10.Evaluate

    Dim oLineByPoints10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints10 = oGCFactory.CreateEntity("LineByPoints", pPOM, "017-LineByPoints")
    oLineByPoints10.Input("StartPoint") = oPointAtCurveExtremity9
    oLineByPoints10.Input("EndPoint") = oPointAtCurveExtremity10
    oLineByPoints10.Evaluate
    
    Dim oEdgePortExtractor10a As SP3DGeometricConstruction.GeometricConstruction
    Set oEdgePortExtractor10a = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "018-EdgePortExtractor")
    oEdgePortExtractor10a.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
    oEdgePortExtractor10a.Inputs("FacePort").Add oFacePortExtractor1
    oEdgePortExtractor10a.Inputs("CoordinateSystem").Add oCSFromMember8
    oEdgePortExtractor10a.Parameter("GeometrySelector") = 2
'    oEdgePortExtractor10a.Parameter("FacesContext") = 1
    oEdgePortExtractor10a.Parameter("LookingAxis") = 1
    oEdgePortExtractor10a.Parameter("TrackFlag") = 1
    oEdgePortExtractor10a.Evaluate

    Dim oEdgePortExtractor10b As SP3DGeometricConstruction.GeometricConstruction
    Set oEdgePortExtractor10b = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "019-EdgePortExtractor")
    oEdgePortExtractor10b.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
    oEdgePortExtractor10b.Inputs("FacePort").Add oFacePortExtractor1
    oEdgePortExtractor10b.Inputs("CoordinateSystem").Add oCSFromMember8
    oEdgePortExtractor10b.Parameter("GeometrySelector") = 2
'    oEdgePortExtractor10b.Parameter("FacesContext") = 1
    oEdgePortExtractor10b.Parameter("LookingAxis") = 1
    oEdgePortExtractor10b.Parameter("TrackFlag") = 2
    oEdgePortExtractor10b.Evaluate

    Dim oPointAtCurveMiddle12a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12a = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "020-PointAtCurveMiddle")
    oPointAtCurveMiddle12a.Inputs("Curve").Add oEdgePortExtractor10a
    oPointAtCurveMiddle12a.Evaluate
    
    Dim oPointAtCurveMiddle12b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12b = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "021-PointAtCurveMiddle")
    oPointAtCurveMiddle12b.Inputs("Curve").Add oEdgePortExtractor10b
    oPointAtCurveMiddle12b.Evaluate
    
    Dim oLineByPoints12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints12 = oGCFactory.CreateEntity("LineByPoints", pPOM, "022-LineByPoints")
    oLineByPoints12.Input("StartPoint") = oPointAtCurveMiddle12a
    oLineByPoints12.Input("EndPoint") = oPointAtCurveMiddle12b
    oLineByPoints12.Evaluate
    
    Dim oPointAtCurveMiddle12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "023-PointAtCurveMiddle")
    oPointAtCurveMiddle12.Inputs("Curve").Add oLineByPoints12
    oPointAtCurveMiddle12.Evaluate
    
    Dim oPointAlongCurve11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "024-PointAlongCurve")
    oPointAlongCurve11.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAlongCurve11.Inputs("Point").Add oPointAtCurveExtremity9
    oPointAlongCurve11.Inputs("TrackPoint").Add oPointAtCurveMiddle6
    oPointAlongCurve11.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve11.Parameter("TrackFlag") = 1
    oPointAlongCurve11.Evaluate

    Dim oParamDistMeasureParallel14 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel14 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "025-ParamDistMeasureParallel")
    oParamDistMeasureParallel14.Inputs("Graphics").Add oPointAtCurveMiddle12
    oParamDistMeasureParallel14.Inputs("Graphics").Add oPointAtCurveMiddle7
    oParamDistMeasureParallel14.Parameter("Value") = 0
    oParamDistMeasureParallel14.Evaluate
    
    Dim oPlaneByPointNormal13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal13 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "026-PlaneByPointNormal")
    oPlaneByPointNormal13.Inputs("Point").Add oPointAlongCurve11
    oPlaneByPointNormal13.Inputs("Line").Add oLineAxisPortExtractor2
    oPlaneByPointNormal13.Parameter("Range") = CDbl(oParamDistMeasureParallel14.Parameter("Value"))
    oPlaneByPointNormal13.Evaluate

    Dim oCurveByIntersection15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection15 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "027-CurveByIntersection")
    oCurveByIntersection15.Inputs("Surface1").Add oPlaneByPointNormal13
    oCurveByIntersection15.Inputs("Surface2").Add oFacePortExtractor2
    oCurveByIntersection15.Inputs("CoordinateSystem").Add oCSFromMember8
    oCurveByIntersection15.Parameter("TrackFlag") = 1
    oCurveByIntersection15.Evaluate

    Dim oCurveByIntersection16 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection16 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "028-CurveByIntersection")
    oCurveByIntersection16.Inputs("Surface1").Add oPlaneByPointNormal13
    oCurveByIntersection16.Inputs("Surface2").Add oFacePortExtractor1
    oCurveByIntersection16.Inputs("CoordinateSystem").Add oCSFromMember8
    oCurveByIntersection16.Parameter("TrackFlag") = 1
    oCurveByIntersection16.Evaluate

    Dim oExtRuledSurface17 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface17 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "029-ExtRuledSurface")
    oExtRuledSurface17.Inputs("Curve1").Add oCurveByIntersection15
    oExtRuledSurface17.Inputs("Curve2").Add oCurveByIntersection16
    oExtRuledSurface17.Parameter("Extension") = 0.1
    oExtRuledSurface17.Evaluate

    Dim oCSByPlane18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane18 = oGCFactory.CreateEntity("CSByPlane", pPOM, "030-CSByPlane")
    oCSByPlane18.Inputs("Plane").Add oExtRuledSurface17.Output("Support", 1)
    oCSByPlane18.Inputs("Point").Add oPointAlongCurve11
    oCSByPlane18.Evaluate

    Dim oLateralBoundaries19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries19 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "031-LateralBoundaries")
    oLateralBoundaries19.Inputs("CoordinateSystem1").Add oCSByPlane18
    oLateralBoundaries19.Inputs("PointS1").Add oExtRuledSurface17.Output("PointS1", 1)
    oLateralBoundaries19.Inputs("PointE1").Add oExtRuledSurface17.Output("PointS2", 1)
    oLateralBoundaries19.Inputs("CoordinateSystem2").Add oCSByPlane18
    oLateralBoundaries19.Inputs("PointS2").Add oExtRuledSurface17.Output("PointE1", 1)
    oLateralBoundaries19.Inputs("PointE2").Add oExtRuledSurface17.Output("PointE2", 1)
    oLateralBoundaries19.Parameter("Extension") = 0.1
    oLateralBoundaries19.Evaluate

    Dim oLateralBoundaries20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries20 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "032-LateralBoundaries")
    oLateralBoundaries20.Inputs("CoordinateSystem1").Add oCSByPlane18
    oLateralBoundaries20.Inputs("PointS1").Add oExtRuledSurface17.Output("PointS1", 1)
    oLateralBoundaries20.Inputs("PointE1").Add oExtRuledSurface17.Output("PointE1", 1)
    oLateralBoundaries20.Inputs("CoordinateSystem2").Add oCSByPlane18
    oLateralBoundaries20.Inputs("PointS2").Add oExtRuledSurface17.Output("PointS2", 1)
    oLateralBoundaries20.Inputs("PointE2").Add oExtRuledSurface17.Output("PointE2", 1)
    oLateralBoundaries20.Parameter("Extension") = 0.1
    oLateralBoundaries20.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    Call SurfaceBody_OrientateFollowingLine(oExtRuledSurface17.Output("Support", 1), oLineByPoints10, False)
    
    oGCMacro.Output("Support", 1) = oExtRuledSurface17.Output("Support", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate1") = oFacePortExtractor2 'oLateralBoundaries19.Output("Boundary", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate2") = oFacePortExtractor1 'oLateralBoundaries19.Output("Boundary", 2)
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oFacePortExtractor5 'oLateralBoundaries20.Output("Boundary", 2)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate1") = oLateralBoundaries20.Output("Boundary", 2)
    End If
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oFacePortExtractor6 'oLateralBoundaries20.Output("Boundary", 1)
    Else
        oGCMacro.Output("Boundary", "PortOfSidePlate2") = oLateralBoundaries20.Output("Boundary", 1)
    End If

    ' populate output "CoordinateSystem"
    Call GeometricConstructionMacro_CreateCoordinateSystemOfSupport(pGeometricConstruction, Nothing, GetGCGeomFactory(), GetGCGeomFactory2(), Position_FromPoint(oPointAlongCurve11))

    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' add new boundaries
        Call PlateSystem_AddNewBoundaries(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), oGCMacro)
    End If
    
    Call DebugOut
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call DebugIn(Method("IJGCMigrate_Migrate"))
    
    ' retrieve the position of the node
    Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = GetPositionOfNode(MyGC)
    
    ' loop on all the inputs, whose name is prefixed by "MemberPart"
    Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = MyGC.definition
    Dim i  As Integer
    For i = 1 To pGeometricConstructionDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
                
        ' migrate these inputs, if one of them has been replaced
        If Mid(sName, 1, Len("MemberPart")) = "MemberPart" Then Call GeometricConstruction_MigrateInputs(MyGC, pMigrateHelper, sName, pPositionOfNode)
    Next
    
    Call DebugOut
End Sub
Function GetPositionOfNode(ByVal MyGC As IJGeometricConstruction) As IJDPosition
    Call DebugIn(Method("GetPositionOfNode"))
    
    Dim MyGCMacro As IJGeometricConstructionMacro: Set MyGCMacro = MyGC
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim pPointAtExtremity As IJGeometricConstruction
    Set pPointAtExtremity = oGCFactory.CreateEntity("PointAtCurveExtremity", Nothing)
    pPointAtExtremity.Input("Curve") = MemberPart_GetLine(MyGC.Input("MemberPart"))
    pPointAtExtremity.Input("TrackPoint") = Point_FromPosition(Position_FromPlane(MyGCMacro.Outputs("Support")(1)))
    pPointAtExtremity.Parameter("TrackFlag") = GCNear
    pPointAtExtremity.Evaluate

    Set GetPositionOfNode = Position_FromPoint(pPointAtExtremity)
    
    Call DebugOut
End Function
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    ' Delegate to the APS when exists
    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Set IJGCToDoDelegate_ToDoDelegate = pGC.ControlledInputs("AdvancedPlateSystem")(1)
    End If
End Property
