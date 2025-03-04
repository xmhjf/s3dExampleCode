VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlate1r"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oRegularAPSHelper As Object
'
' definition of the error codes
'
Private Enum Errors
    MISSING_TRANSITION_PLATE1 = 2
    MISSING_TRANSITION_PLATE2 = 3
    MISSING_MEMBERPART = 4
    DUPLICATE_INPUTS = 5
    MEMBER_AND_PLATE_DO_NOT_INTERSECT = 6
End Enum

'
' implemented interfaces
'
Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGCToDoDelegate

'
' some definitions for debug purpose
'
Dim m_oAdvancedDebug As New AdvancedDebug
Private Property Get Source() As String
    Let Source = "GCCustomPlate.CustomPlate"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
'
' constructor/destructor of the class
'
Private Sub Class_Initialize()
    Call m_oAdvancedDebug.StartSource(Source)
    
    Set m_oRegularAPSHelper = CreateObject("GCRegularAPS.RegularAPSHelper")
End Sub
Private Sub Class_Terminate()
    Set m_oRegularAPSHelper = Nothing
    
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub
'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oRegularAPSHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oRegularAPSHelper.PostInitialize(pGCDefinition)
    
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call m_oRegularAPSHelper.PreEvaluate(pGC, pPOM)
    Call Evaluate(pGC, pPOM)
    Call m_oRegularAPSHelper.PostEvaluate(pGC, pPOM)

    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oRegularAPSHelper.ProcessError(pGC, Err.Number)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oRegularAPSHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oRegularAPSHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Dim oGenericPlateHelper As Object
    Set oGenericPlateHelper = CreateObject("GCGenericPlateHelper.GenericPlateHelper")

    Call oGenericPlateHelper.Migrate(MyGC, pMigrateHelper)
    Set oGenericPlateHelper = Nothing

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal MyGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))
    
    Set IJGCToDoDelegate_ToDoDelegate = m_oRegularAPSHelper.ToDoDelegate(MyGC)
    
    Call m_oAdvancedDebug.ExitMethod
End Property
'
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a rolled Member", "ISPSMemberPartPrismatic", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate1", "Select the first limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate2", "Select the second limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate1", "Select a first lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate2", "Select a second lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePortOfmemberPart")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePortOfTransitionPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePortOfTransitionPlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePortOfMemberPart")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePortOfSidePlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePortOfSidePlate2")
    
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", GCCodeList, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Inside", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Outside", 2)
    
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE1, "MissingTransitionPlate1", "Missing the transition plate #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE2, "MissingTransitionPlate2", "Missing the transition plate #2")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBERPART, "MissingMemberPart", "Missing the member part")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_INPUTS, "DuplicateInputs", "2 inputs are the same")
    Call pGeometricConstructionDefinition.AddErrorValue(MEMBER_AND_PLATE_DO_NOT_INTERSECT, "MemberAndPlateDoNotIntersect", "The face of the member does not intersect the transition plate #1")
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    ' check inputs
    If pGeometricConstruction.Inputs("TransitionPlate1").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE1
    If pGeometricConstruction.Inputs("TransitionPlate2").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE2
    If pGeometricConstruction.Inputs("MemberPart").Count = 0 Then Err.Raise MISSING_MEMBERPART
    If True Then
        Dim pElements As IJElements: Set pElements = New JObjectCollection
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

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("EdgePortOfmemberPart").Clear
    pGeometricConstruction.ControlledInputs("EdgePortOfmemberPart").AddElements oLineAxisPortExtractor1.ControlledInputs("Port")

    Dim oFacePortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor2 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "002-FacePortExtractor")
    oFacePortExtractor2.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
    oFacePortExtractor2.Parameter("GeometrySelector") = 4
    oFacePortExtractor2.Parameter("FacesContext") = 1
    oFacePortExtractor2.Parameter("LookingAxis") = 1
    oFacePortExtractor2.Parameter("SurfaceType") = 1
    oFacePortExtractor2.Parameter("TrackFlag") = 2
    oFacePortExtractor2.Parameter("Offset") = 0#
    oFacePortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("FacePortOfTransitionPlate1").Clear
    pGeometricConstruction.ControlledInputs("FacePortOfTransitionPlate1").AddElements oFacePortExtractor2.ControlledInputs("Port")

    Dim oFacePortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor3 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "003-FacePortExtractor")
    oFacePortExtractor3.Inputs("Connectable").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
    oFacePortExtractor3.Parameter("GeometrySelector") = 4
    oFacePortExtractor3.Parameter("FacesContext") = 1
    oFacePortExtractor3.Parameter("LookingAxis") = 1
    oFacePortExtractor3.Parameter("SurfaceType") = 1
    oFacePortExtractor3.Parameter("TrackFlag") = 2
    oFacePortExtractor3.Parameter("Offset") = 0#
    oFacePortExtractor3.Evaluate
    pGeometricConstruction.ControlledInputs("FacePortOfTransitionPlate2").Clear
    pGeometricConstruction.ControlledInputs("FacePortOfTransitionPlate2").AddElements oFacePortExtractor3.ControlledInputs("Port")

    Dim oFacePortExtractor2a As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        Set oFacePortExtractor2a = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "004-FacePortExtractor")
        oFacePortExtractor2a.Inputs("Connectable").Add pGeometricConstruction.Inputs("SidePlate1").Item(1)
        oFacePortExtractor2a.Parameter("GeometrySelector") = 2
        oFacePortExtractor2a.Parameter("FacesContext") = 1
        oFacePortExtractor2a.Parameter("LookingAxis") = 1
        oFacePortExtractor2a.Parameter("TrackFlag") = 1
        oFacePortExtractor2a.Parameter("Offset") = 0
        oFacePortExtractor2a.Evaluate
        pGeometricConstruction.ControlledInputs("FacePortOfSidePlate1").Clear
        pGeometricConstruction.ControlledInputs("FacePortOfSidePlate1").AddElements oFacePortExtractor2a.ControlledInputs("Port")
    End If
    
    Dim oFacePortExtractor3a As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        Set oFacePortExtractor3a = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "005-FacePortExtractor")
        oFacePortExtractor3a.Inputs("Connectable").Add pGeometricConstruction.Inputs("SidePlate2").Item(1)
        oFacePortExtractor3a.Parameter("GeometrySelector") = 2
        oFacePortExtractor3a.Parameter("FacesContext") = 1
        oFacePortExtractor3a.Parameter("LookingAxis") = 1
        oFacePortExtractor3a.Parameter("TrackFlag") = 1
        oFacePortExtractor3a.Parameter("Offset") = 0
        oFacePortExtractor3a.Evaluate
        pGeometricConstruction.ControlledInputs("FacePortOfSidePlate2").Clear
        pGeometricConstruction.ControlledInputs("FacePortOfSidePlate2").AddElements oFacePortExtractor3a.ControlledInputs("Port")
    End If
   
    Dim oCSByPlane4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane4 = oGCFactory.CreateEntity("CSByPlane", pPOM, "004-CSByPlane")
    oCSByPlane4.Inputs("Plane").Add oFacePortExtractor2
    oCSByPlane4.Evaluate

    Dim oLineFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS5 = oGCFactory.CreateEntity("LineFromCS", pPOM, "005-LineFromCS")
    oLineFromCS5.Inputs("CoordinateSystem").Add oCSByPlane4
    oLineFromCS5.Parameter("LookingAxis") = 3
    oLineFromCS5.Parameter("Length") = 1#
    oLineFromCS5.Evaluate

    Dim oCurveByProjection6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection6 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "006-CurveByProjection")
    oCurveByProjection6.Inputs("Curve").Add oLineAxisPortExtractor1
    oCurveByProjection6.Inputs("Surface").Add oFacePortExtractor2
    oCurveByProjection6.Inputs("Line").Add oLineFromCS5
    oCurveByProjection6.Parameter("TrackFlag") = 1
    oCurveByProjection6.Evaluate

    Dim oCSByLines7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines7 = oGCFactory.CreateEntity("CSByLines", pPOM, "007-CSByLines")
    oCSByLines7.Inputs("AxisLine1").Add oLineAxisPortExtractor1
    oCSByLines7.Inputs("AxisLine2").Add oCurveByProjection6
    oCSByLines7.Parameter("AxesRoles") = 1
    oCSByLines7.Parameter("CSOrientation") = 1
    oCSByLines7.Evaluate

    Dim oFacePortExtractor8 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor8 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "008-FacePortExtractor")
    oFacePortExtractor8.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor8.Inputs("CoordinateSystem").Add oCSByLines7
    oFacePortExtractor8.Parameter("GeometrySelector") = 4
    oFacePortExtractor8.Parameter("FacesContext") = 2
    oFacePortExtractor8.Parameter("LookingAxis") = 2
    oFacePortExtractor8.Parameter("SurfaceType") = 1
    If pGeometricConstruction.Parameter("Location") = 1 Then
        oFacePortExtractor8.Parameter("TrackFlag") = 2
    Else
        oFacePortExtractor8.Parameter("TrackFlag") = 1
    End If
    oFacePortExtractor8.Parameter("Offset") = 0#
    oFacePortExtractor8.Evaluate
    pGeometricConstruction.ControlledInputs("FacePortOfMemberPart").Clear
    pGeometricConstruction.ControlledInputs("FacePortOfMemberPart").AddElements oFacePortExtractor8.ControlledInputs("Port")

    Dim oCurveByIntersection9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection9 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "009-CurveByIntersection")
    oCurveByIntersection9.Inputs("Surface1").Add oFacePortExtractor8
    oCurveByIntersection9.Inputs("Surface2").Add oFacePortExtractor2
    oCurveByIntersection9.Parameter("TrackFlag") = 2
    On Error Resume Next
    oCurveByIntersection9.Evaluate
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise MEMBER_AND_PLATE_DO_NOT_INTERSECT
    End If
    On Error GoTo 0
    
    Dim oCurveByProjection10 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection10 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "010-CurveByProjection")
    oCurveByProjection10.Inputs("Curve").Add oCurveByIntersection9
    oCurveByProjection10.Inputs("Surface").Add oFacePortExtractor3
    oCurveByProjection10.Inputs("Line").Add oLineFromCS5
    oCurveByProjection10.Parameter("TrackFlag") = 1
    oCurveByProjection10.Evaluate

    Dim oExtRuledSurface11 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface11 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "011-ExtRuledSurface")
    oExtRuledSurface11.Inputs("Curve1").Add oCurveByIntersection9
    oExtRuledSurface11.Inputs("Curve2").Add oCurveByProjection10
    oExtRuledSurface11.Parameter("Extension") = 0.1
    oExtRuledSurface11.Evaluate

    Dim oCSByPlane12 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane12 = oGCFactory.CreateEntity("CSByPlane", pPOM, "012-CSByPlane")
    oCSByPlane12.Inputs("Plane").Add oExtRuledSurface11.Output("Support", 1)
    oCSByPlane12.Evaluate

    Dim oLateralBoundaries13 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries13 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "013-LateralBoundaries")
    oLateralBoundaries13.Inputs("CoordinateSystem1").Add oCSByPlane12
    oLateralBoundaries13.Inputs("PointS1").Add oExtRuledSurface11.Output("PointE1", 1)
    oLateralBoundaries13.Inputs("PointE1").Add oExtRuledSurface11.Output("PointS1", 1)
    oLateralBoundaries13.Inputs("CoordinateSystem2").Add oCSByPlane12
    oLateralBoundaries13.Inputs("PointS2").Add oExtRuledSurface11.Output("PointE2", 1)
    oLateralBoundaries13.Inputs("PointE2").Add oExtRuledSurface11.Output("PointS2", 1)
    oLateralBoundaries13.Parameter("Extension") = 0.1
    oLateralBoundaries13.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oExtRuledSurface11.Output("Support", 1)
    oGCMacro.Output("Boundary", "FacePortOfTransitionPlate1") = oFacePortExtractor2.Output
    oGCMacro.Output("Boundary", "FacePortOfTransitionPlate2") = oFacePortExtractor3.Output
    If pGeometricConstruction.Inputs("SidePlate1").Count = 1 Then
        oGCMacro.Output("Boundary", "FacePortOfSidePlate1") = oFacePortExtractor2a.Output
    Else
        oGCMacro.Output("Boundary", "FacePortOfSidePlate1") = oLateralBoundaries13.Output("Boundary", 1)
    End If
    If pGeometricConstruction.Inputs("SidePlate2").Count = 1 Then
        oGCMacro.Output("Boundary", "FacePortOfSidePlate2") = oFacePortExtractor3a.Output
    Else
        oGCMacro.Output("Boundary", "FacePortOfSidePlate2") = oLateralBoundaries13.Output("Boundary", 2)
    End If
End Sub
