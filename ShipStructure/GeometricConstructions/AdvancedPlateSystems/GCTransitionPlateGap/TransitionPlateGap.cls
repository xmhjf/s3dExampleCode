VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransitionPlateGap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oGenericPlateHelper As Object
'
' definition of the error codes
'
Private Enum Errors
    MISSING_MANDATORY_INPUT = 1
    DUPLICATE_PROFILES = 2
    DUPLICATE_PLATE_SYSTEMS = 3
    MISSING_GAP = 4
    BOUNDING_PLATE_SYSTEM_NOT_INTERSECTING = 5
End Enum
'
' implemented interfaces
'
Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGCMirror
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

'
' some definitions for debug purpose
'
Dim m_oAdvancedDebug As New AdvancedDebug
Private Property Get Source() As String
    Let Source = "GCTransitionPlateGap.TransitionPlateGap"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
'
' constructor/destructor of the class
'
Private Sub Class_Initialize()
    Call m_oAdvancedDebug.StartSource(Source)
    
    Set m_oGenericPlateHelper = CreateObject("GCGenericPlateHelper.GenericPlateHelper")
End Sub
Private Sub Class_Terminate()
    Set m_oGenericPlateHelper = Nothing
    
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub

Private Sub IJGCConnectedElementSC_PostDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PostDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PostDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
   Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PreDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PreDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod


End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oGenericPlateHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oGenericPlateHelper.PostInitialize(pGCDefinition)
    
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Dim bEvaluateGeometry As Boolean
    Call m_oGenericPlateHelper.PreEvaluate(pGC, pPOM, bEvaluateGeometry)
    Call m_oAdvancedDebug.EnterMethod(Method("EvaluateGeometry"))
    If bEvaluateGeometry Then
        Call Evaluate(pGC, pPOM)
    Else
        Call m_oAdvancedDebug.ShowMsg("Skip")
    End If
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.PostEvaluate(pGC, pPOM)

'''    Dim oAPS As Object:
'''    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
'''        Set oAPS = pGC.ControlledInputs("AdvancedPlateSystem")(1)
'''        Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = oAPS
'''        pPlateMoldedConventions.plateThicknessDirection = WithNormalDir
'''    End If
    
    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oAdvancedDebug.ExitMethod
    Call pGC.RaiseError(Err.Number)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oGenericPlateHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oGenericPlateHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Call m_oGenericPlateHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Public Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As GeometricConstruction, ByVal vInfo As Variant, ByRef bIsOnRibbonBar As Boolean)
    Call m_oGenericPlateHelper.IsOnRibbonBar(sName, pGC, vInfo, bIsOnRibbonBar)
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))

    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    On Error Resume Next
    Set IJGCToDoDelegate_ToDoDelegate = m_oGenericPlateHelper.ToDoDelegate(pGC)
    On Error GoTo 0

    Call m_oAdvancedDebug.ExitMethod
End Property
'
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Profile1", "Select a first ProfileSystem, MemberPart or Built-Up", "IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] OR ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 1, "")
    Call pGeometricConstructionDefinition.AddInput("Profile2", "Select a second ProfileSystem, MemberPart or Built-Up", "IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] OR ISPSMemberPartPrismatic OR ISPSDesignedMember", 1, 1, "")
    Call pGeometricConstructionDefinition.AddInput("PlateSystem1", "Select a first bounding PlateSystem", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "")
    Call pGeometricConstructionDefinition.AddInput("PlateSystem2", "Select a second bounding PlateSystem", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort4")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem1", "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem1", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem2", "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem2", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort101")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort102")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort103")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort104")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort201")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort202")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort106")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort107")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort108")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort204")
    
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", GCCodeList, 0, 0, 0, 0, 3)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Bottom", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Centered-horiz", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Top", 3)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Left", 4)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Centered-vert", 5)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Right", 6)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCLocalCoordinateSystem, "CoordinateSystem")
    
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_PROFILES, "ErrorValue2", "Selected Profiles are identical")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_PLATE_SYSTEMS, "ErrorValue3", "Selected PlateSystems are identical")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_GAP, "ErrorValue4", "Missing a gap between the 2 Members")
    Call pGeometricConstructionDefinition.AddErrorValue(BOUNDING_PLATE_SYSTEM_NOT_INTERSECTING, "ErrorValue5", "A bounding PlateSystems doesn't intersect the support")
    
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    If pGeometricConstruction.Inputs("Profile1").Count <> 1 _
    And pGeometricConstruction.Inputs("Profile1").Count <> 1 Then
        Err.Raise MISSING_MANDATORY_INPUT
    End If
    
    If pGeometricConstruction.Inputs("Profile1")(1) Is pGeometricConstruction.Inputs("Profile2")(1) Then
        Err.Raise DUPLICATE_PROFILES
    End If
    If pGeometricConstruction.Inputs("PlateSystem1").Count = 1 _
    And pGeometricConstruction.Inputs("PlateSystem2").Count = 1 Then
        If pGeometricConstruction.Inputs("PlateSystem1")(1) Is pGeometricConstruction.Inputs("PlateSystem2")(1) Then
            Err.Raise DUPLICATE_PLATE_SYSTEMS
        End If
    End If
    
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCTop3, GCLeft4, GCRight6:
            If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
                Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)
                If pPlateMoldedConventions.plateThicknessDirection <> WithNormalDir Then pPlateMoldedConventions.plateThicknessDirection = WithNormalDir
            End If
        Case GCCenteredHor2, GCCenteredVer5:
            If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
                Set pPlateMoldedConventions = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)
                If pPlateMoldedConventions.plateThicknessDirection <> Centered Then pPlateMoldedConventions.plateThicknessDirection = Centered
            End If
    End Select
    
    Dim oAxis1 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oAxis1 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "001-AxisPortExtractor")
        oAxis1.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oAxis1.Parameter("GeometrySelector") = GCCurrent
        oAxis1.Evaluate
        pGeometricConstruction.ControlledInputs("AxisPort1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort1").AddElements oAxis1.ControlledInputs("Port")
    Else
        Set oAxis1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
        oAxis1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oAxis1.Evaluate
        pGeometricConstruction.ControlledInputs("AxisPort1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort1").AddElements oAxis1.ControlledInputs("Port")
    End If

    Dim oAxis2 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        Set oAxis2 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "002-AxisPortExtractor")
        oAxis2.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oAxis2.Parameter("GeometrySelector") = GCCurrent
        oAxis2.Evaluate
        pGeometricConstruction.ControlledInputs("AxisPort2").Clear
        pGeometricConstruction.ControlledInputs("AxisPort2").AddElements oAxis2.ControlledInputs("Port")
    Else
        Set oAxis2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
        oAxis2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oAxis2.Evaluate
        pGeometricConstruction.ControlledInputs("AxisPort1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort1").AddElements oAxis2.ControlledInputs("Port")
    End If
    
    Dim oClosestEndPoints3 As SP3DGeometricConstruction.GeometricConstruction
    Set oClosestEndPoints3 = oGCFactory.CreateEntity("PointsAtClosestEndsOfCurves", pPOM, "003-PointsAtClosestEndsOfCurves")
    oClosestEndPoints3.Inputs("Curve1").Add oAxis1
    oClosestEndPoints3.Inputs("Curve2").Add oAxis2
    oClosestEndPoints3.Evaluate
    
    Dim oParamDistMeasureParallel4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel4 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "004-ParamDistMeasureParallel")
    oParamDistMeasureParallel4.Inputs("Graphics").Add oClosestEndPoints3.Output("Point1", 1)
    oParamDistMeasureParallel4.Inputs("Graphics").Add oClosestEndPoints3.Output("Point2", 1)
    oParamDistMeasureParallel4.Evaluate
    
    If oParamDistMeasureParallel4.Parameter("Value") < 0.001 Then Err.Raise MISSING_GAP
    
    Dim oCS5 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointAlongCurve5 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPlaneByPointNormal5 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oPointAlongCurve5 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "005a-PointAlongCurve")
        oPointAlongCurve5.Inputs("Curve").Add oAxis1
        oPointAlongCurve5.Inputs("Point").Add oClosestEndPoints3.Output("Point1", 1)
        oPointAlongCurve5.Inputs("TrackPoint").Add oClosestEndPoints3.Output("Point2", 1)
        oPointAlongCurve5.Parameter("Distance") = 0.001
        oPointAlongCurve5.Parameter("TrackFlag") = GCFar
        oPointAlongCurve5.Evaluate
        
        Dim oVectorTangentToCurve5 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve5 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "005-VectorTangentToCurve")
        oVectorTangentToCurve5.Inputs("Curve").Add oAxis1
        oVectorTangentToCurve5.Inputs("Point").Add oPointAlongCurve5
        oVectorTangentToCurve5.Parameter("Range") = 1#
        oVectorTangentToCurve5.Evaluate
        
        'CSByStiffener does not work if use plane at oPointAlongCurve5 ?
        Set oPlaneByPointNormal5 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "005-PlaneByPointNormal")
        oPlaneByPointNormal5.Inputs("Point").Add oPointAlongCurve5
        oPlaneByPointNormal5.Inputs("Line").Add oVectorTangentToCurve5
        oPlaneByPointNormal5.Parameter("Range") = 5#
        oPlaneByPointNormal5.Evaluate
        
        Set oCS5 = oGCFactory.CreateEntity("CSByStiffener", pPOM, "005-CSByStiffener")
        oCS5.Inputs("Stiffener").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oCS5.Inputs("Plane").Add oPlaneByPointNormal5 'oClosestEndPoints3.Output("Point1", 1)
        oCS5.Evaluate
    Else
        Set oCS5 = oGCFactory.CreateEntity("CSFromMember", pPOM, "005-CSFromMember")
        oCS5.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oCS5.Inputs("Point").Add oClosestEndPoints3.Output("Point1", 1)
        oCS5.Evaluate
    End If
    
    Dim oLineFromCS5a As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS5a = oGCFactory.CreateEntity("LineFromCS", pPOM, "005a-LineFromCS")
    oLineFromCS5a.Inputs("CoordinateSystem").Add oCS5
    oLineFromCS5a.Parameter("LookingAxis") = GCYAxis
    oLineFromCS5a.Evaluate
    
    Dim oLineFromCS5b As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS5b = oGCFactory.CreateEntity("LineFromCS", pPOM, "005b-LineFromCS")
    oLineFromCS5b.Inputs("CoordinateSystem").Add oCS5
    oLineFromCS5b.Parameter("LookingAxis") = GCZAxis
    oLineFromCS5b.Evaluate
    
    Dim oCSByLines5c As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines5c = oGCFactory.CreateEntity("CSByLines", pPOM, "005c-CSByLines")
    oCSByLines5c.Inputs("AxisLine1").Add oLineFromCS5a
    oCSByLines5c.Inputs("AxisLine2").Add oLineFromCS5b
    oCSByLines5c.Inputs("TrackPoint").Add oClosestEndPoints3.Output("Point2", 1)
    oCSByLines5c.Parameter("AxesRoles") = GCYZ
    oCSByLines5c.Parameter("TrackFlag") = GCNear
    oCSByLines5c.Evaluate
    
    Dim oCS6 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointAlongCurve6 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPlaneByPointNormal6 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        Set oPointAlongCurve6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "006a-PointAlongCurve")
        oPointAlongCurve6.Inputs("Curve").Add oAxis2
        oPointAlongCurve6.Inputs("Point").Add oClosestEndPoints3.Output("Point2", 1)
        oPointAlongCurve6.Inputs("TrackPoint").Add oClosestEndPoints3.Output("Point1", 1)
        oPointAlongCurve6.Parameter("Distance") = 0.001
        oPointAlongCurve6.Parameter("TrackFlag") = GCFar
        oPointAlongCurve6.Evaluate
        
        Dim oVectorTangentToCurve6 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve6 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "006-VectorTangentToCurve")
        oVectorTangentToCurve6.Inputs("Curve").Add oAxis2
        oVectorTangentToCurve6.Inputs("Point").Add oPointAlongCurve6
        oVectorTangentToCurve6.Parameter("Range") = 1#
        oVectorTangentToCurve6.Evaluate
        
        Set oPlaneByPointNormal6 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "006-PlaneByPointNormal")
        oPlaneByPointNormal6.Inputs("Point").Add oPointAlongCurve6
        oPlaneByPointNormal6.Inputs("Line").Add oVectorTangentToCurve6
        oPlaneByPointNormal6.Parameter("Range") = 5#
        oPlaneByPointNormal6.Evaluate
        
        Set oCS6 = oGCFactory.CreateEntity("CSByStiffener", pPOM, "006-CSByStiffener")
        oCS6.Inputs("Stiffener").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oCS6.Inputs("Plane").Add oPlaneByPointNormal6 'oClosestEndPoints3.Output("Point2", 1)
        oCS6.Evaluate
    Else
        Set oCS6 = oGCFactory.CreateEntity("CSFromMember", pPOM, "006-CSFromMember")
        oCS6.Inputs("MemberPart").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oCS6.Inputs("Point").Add oClosestEndPoints3.Output("Point2", 1)
        oCS6.Evaluate
    End If
    
    Dim oLineFromCS6a As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS6a = oGCFactory.CreateEntity("LineFromCS", pPOM, "006a-LineFromCS")
    oLineFromCS6a.Inputs("CoordinateSystem").Add oCS6
    oLineFromCS6a.Parameter("LookingAxis") = GCYAxis
    oLineFromCS6a.Evaluate
    
    Dim oLineFromCS6b As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS6b = oGCFactory.CreateEntity("LineFromCS", pPOM, "006b-LineFromCS")
    oLineFromCS6b.Inputs("CoordinateSystem").Add oCS6
    oLineFromCS6b.Parameter("LookingAxis") = GCZAxis
    oLineFromCS6b.Evaluate
    
    Dim oCSByLines6c As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines6c = oGCFactory.CreateEntity("CSByLines", pPOM, "006c-CSByLines")
    oCSByLines6c.Inputs("AxisLine1").Add oLineFromCS6a
    oCSByLines6c.Inputs("AxisLine2").Add oLineFromCS6b
    oCSByLines6c.Inputs("TrackPoint").Add oClosestEndPoints3.Output("Point1", 1)
    oCSByLines6c.Parameter("AxesRoles") = GCYZ
    oCSByLines6c.Parameter("TrackFlag") = GCNear
    oCSByLines6c.Evaluate
    
    Dim oFacePortExtractor7 As SP3DGeometricConstruction.GeometricConstruction
    'Get capping faces of Bounding plates near Y axis
    If pGeometricConstruction.Inputs("PlateSystem1").Count = 1 Then
        Set oFacePortExtractor7 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "007-FacePortExtractor")
        oFacePortExtractor7.Inputs("Connectable").Add pGeometricConstruction.Inputs("PlateSystem1").Item(1)
        oFacePortExtractor7.Parameter("GeometrySelector") = GCCurrent
        oFacePortExtractor7.Parameter("FacesContext") = GCCapping
        oFacePortExtractor7.Parameter("LookingAxis") = GCYAxis
        oFacePortExtractor7.Parameter("TrackFlag") = GCNear
        oFacePortExtractor7.Parameter("Offset") = 0
        oFacePortExtractor7.Public = True
        oFacePortExtractor7.Evaluate
        pGeometricConstruction.ControlledInputs("FacePort3").Clear
        pGeometricConstruction.ControlledInputs("FacePort3").AddElements oFacePortExtractor7.ControlledInputs("Port")
    End If
    
    Dim oFacePortExtractor8 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Inputs("PlateSystem2").Count = 1 Then
        Set oFacePortExtractor8 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "008-FacePortExtractor")
        oFacePortExtractor8.Inputs("Connectable").Add pGeometricConstruction.Inputs("PlateSystem2").Item(1)
        oFacePortExtractor8.Parameter("GeometrySelector") = GCCurrent
        oFacePortExtractor8.Parameter("FacesContext") = GCCapping
        oFacePortExtractor8.Parameter("LookingAxis") = GCYAxis
        oFacePortExtractor8.Parameter("TrackFlag") = GCNear
        oFacePortExtractor8.Parameter("Offset") = 0
        oFacePortExtractor8.Public = True
        oFacePortExtractor8.Evaluate
        pGeometricConstruction.ControlledInputs("FacePort4").Clear
        pGeometricConstruction.ControlledInputs("FacePort4").AddElements oFacePortExtractor8.ControlledInputs("Port")
    End If
    
    
    Dim oFacePortExtractor101 As SP3DGeometricConstruction.GeometricConstruction
    Dim oFacePortExtractor102 As SP3DGeometricConstruction.GeometricConstruction
    
    If TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
    
        'Create small TubeFromDetailedProfile for first profile
        Dim oTubeFromDetailedProfile1 As SP3DGeometricConstruction.GeometricConstruction
        Set oTubeFromDetailedProfile1 = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "01-TubeFromDetailedProfile")
        oTubeFromDetailedProfile1.Inputs("RootProfileSystem").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oTubeFromDetailedProfile1.Inputs("ReferencePoint").Add oClosestEndPoints3.Output("Point1", 1)
        oTubeFromDetailedProfile1.Parameter("Length") = 1#
        oTubeFromDetailedProfile1.Evaluate
        pGeometricConstruction.ControlledInputs("RootPlateSystem1").Clear
        pGeometricConstruction.ControlledInputs("RootPlateSystem1").AddElements oTubeFromDetailedProfile1.ControlledInputs("RootPlateSystem")
        pGeometricConstruction.ControlledInputs("LeafPlateSystem1").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem1").AddElements oTubeFromDetailedProfile1.ControlledInputs("LeafPlateSystem")
        
        'Create SurfaceBodyExtractors for first stiffener
        Set oFacePortExtractor101 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "101-SurfaceBodyExtractor")
        oFacePortExtractor101.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile1
        oFacePortExtractor101.Inputs("CoordinateSystem").Add oCSByLines5c
        
        Set oFacePortExtractor102 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "102-SurfaceBodyExtractor")
        oFacePortExtractor102.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile1
        oFacePortExtractor102.Inputs("CoordinateSystem").Add oCSByLines5c
        
    Else
        Set oFacePortExtractor101 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "101-FacePortExtractor")
        oFacePortExtractor101.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oFacePortExtractor101.Inputs("CoordinateSystem").Add oCSByLines5c
        oFacePortExtractor101.Parameter("GeometrySelector") = GCCurrent
        
        Set oFacePortExtractor102 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "102-FacePortExtractor")
        oFacePortExtractor102.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oFacePortExtractor102.Inputs("CoordinateSystem").Add oCSByLines5c
        oFacePortExtractor102.Parameter("GeometrySelector") = GCCurrent
        
    End If
    
    oFacePortExtractor101.Parameter("FacesContext") = GCLateral
    oFacePortExtractor101.Parameter("IntersectingPlane") = 0
    oFacePortExtractor101.Parameter("SurfaceType") = 1
    oFacePortExtractor101.Parameter("Offset") = 0#
    
    oFacePortExtractor102.Parameter("FacesContext") = GCLateral
    oFacePortExtractor102.Parameter("IntersectingPlane") = 0
    oFacePortExtractor102.Parameter("SurfaceType") = 1
    oFacePortExtractor102.Parameter("Offset") = 0#
    
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1:
            oFacePortExtractor101.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCNear
            oFacePortExtractor102.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCFar
        Case GCCenteredHor2:
            oFacePortExtractor101.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCNear
            oFacePortExtractor102.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCFar
        Case GCTop3:
            oFacePortExtractor101.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCFar
            oFacePortExtractor102.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCNear
        Case GCLeft4:
            oFacePortExtractor101.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCFar
            oFacePortExtractor102.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCNear
        Case GCCenteredVer5:
            oFacePortExtractor101.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCFar
            oFacePortExtractor102.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCNear
        Case GCRight6:
            oFacePortExtractor101.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor101.Parameter("TrackFlag") = GCNear
            oFacePortExtractor102.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor102.Parameter("TrackFlag") = GCFar
    End Select
    
    oFacePortExtractor101.Evaluate
    
    If Not TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        pGeometricConstruction.ControlledInputs("FacePort101").Clear
        pGeometricConstruction.ControlledInputs("FacePort101").AddElements oFacePortExtractor101.ControlledInputs("Port")
    End If
    
    oFacePortExtractor102.Evaluate
    If Not TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        pGeometricConstruction.ControlledInputs("FacePort102").Clear
        pGeometricConstruction.ControlledInputs("FacePort102").AddElements oFacePortExtractor102.ControlledInputs("Port")
    End If
   
   'Create track point from CsByStiffener1 and stiffener2
   Dim oTrackPoint105 As SP3DGeometricConstruction.GeometricConstruction
    Set oTrackPoint105 = oGCFactory.CreateEntity("PointFromCS", pPOM, "105-TrackPoint")
    oTrackPoint105.Inputs("CoordinateSystem").Add oCS5
    oTrackPoint105.Parameter("X") = 0
    oTrackPoint105.Parameter("Y") = 0
    oTrackPoint105.Parameter("Z") = 0
    oTrackPoint105.Evaluate
   
    Dim oEdgePortExtractor105 As SP3DGeometricConstruction.GeometricConstruction
    If Not TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oEdgePortExtractor105 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "105-EdgePortExtractor")
        oEdgePortExtractor105.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oEdgePortExtractor105.Inputs("FacePort").Add oFacePortExtractor101
        oEdgePortExtractor105.Inputs("CoordinateSystem").Add oCSByLines5c
        oEdgePortExtractor105.Parameter("GeometrySelector") = GCCurrent
        oEdgePortExtractor105.Parameter("LookingAxis") = GCXAxis
        oEdgePortExtractor105.Parameter("TrackFlag") = GCFar
        oEdgePortExtractor105.Evaluate
        pGeometricConstruction.ControlledInputs("EdgePort1").Clear
        pGeometricConstruction.ControlledInputs("EdgePort1").AddElements oEdgePortExtractor105.ControlledInputs("Port")
    Else
        Set oEdgePortExtractor105 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "105-EdgeFromCurveOrSurf")
        oEdgePortExtractor105.Inputs("Geometry").Add oFacePortExtractor101
        oEdgePortExtractor105.Inputs("TrackPoint").Add oTrackPoint105
        oEdgePortExtractor105.Parameter("TrackFlag") = GCNear
        oEdgePortExtractor105.Evaluate
    End If
    
    Dim oEdgePortExtractor106 As SP3DGeometricConstruction.GeometricConstruction
    If Not TypeOf pGeometricConstruction.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oEdgePortExtractor106 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "106-EdgePortExtractor")
        oEdgePortExtractor106.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile1").Item(1)
        oEdgePortExtractor106.Inputs("FacePort").Add oFacePortExtractor102
        oEdgePortExtractor106.Inputs("CoordinateSystem").Add oCSByLines5c
        oEdgePortExtractor106.Parameter("GeometrySelector") = GCCurrent
        oEdgePortExtractor106.Parameter("LookingAxis") = GCXAxis
        oEdgePortExtractor106.Parameter("TrackFlag") = GCFar
        oEdgePortExtractor106.Evaluate
        pGeometricConstruction.ControlledInputs("EdgePort106").Clear
        pGeometricConstruction.ControlledInputs("EdgePort106").AddElements oEdgePortExtractor106.ControlledInputs("Port")
    Else
        Set oEdgePortExtractor106 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "106-EdgeFromCurveOrSurf")
        oEdgePortExtractor106.Inputs("Geometry").Add oFacePortExtractor102
        oEdgePortExtractor106.Inputs("TrackPoint").Add oTrackPoint105
        oEdgePortExtractor106.Parameter("TrackFlag") = GCNear
        oEdgePortExtractor106.Evaluate
    End If
    
    Dim oCurveOrientate114 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate114 = oGCFactory.CreateEntity("CurveOrientate", pPOM, "114-CurveOrientate")
    oCurveOrientate114.Inputs("Curve").Add oEdgePortExtractor105
    oCurveOrientate114.Inputs("CoordinateSystem").Add oCSByLines5c
    oCurveOrientate114.Evaluate

    Dim oCurveOrientate115 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate115 = oGCFactory.CreateEntity("CurveOrientate", pPOM, "115-CurveOrientate")
    oCurveOrientate115.Inputs("Curve").Add oEdgePortExtractor106
    oCurveOrientate115.Inputs("CoordinateSystem").Add oCSByLines5c
    oCurveOrientate115.Evaluate

    Dim oPointAtCurveStart116 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart116 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "116-PointAtCurveStart")
    oPointAtCurveStart116.Inputs("Curve").Add oCurveOrientate114
    oPointAtCurveStart116.Evaluate
    
    Dim oPointAtCurveStart117 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart117 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "117-PointAtCurveStart")
    oPointAtCurveStart117.Inputs("Curve").Add oCurveOrientate115
    oPointAtCurveStart117.Evaluate
    
    Dim oPointAtCurveEnd118 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd118 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "118-PointAtCurveEnd")
    oPointAtCurveEnd118.Inputs("Curve").Add oCurveOrientate114
    oPointAtCurveEnd118.Evaluate
        
    Dim oPointAtCurveEnd119 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd119 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "119-PointAtCurveEnd")
    oPointAtCurveEnd119.Inputs("Curve").Add oCurveOrientate115
    oPointAtCurveEnd119.Evaluate
    
    Dim oFacePortExtractor201 As SP3DGeometricConstruction.GeometricConstruction
    Dim oFacePortExtractor202 As SP3DGeometricConstruction.GeometricConstruction
    
    If TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
    
        'Create small TubeFromDetailedProfile for second profile
        Dim oTubeFromDetailedProfile2 As SP3DGeometricConstruction.GeometricConstruction
        Set oTubeFromDetailedProfile2 = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "02-TubeFromDetailedProfile")
        oTubeFromDetailedProfile2.Inputs("RootProfileSystem").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oTubeFromDetailedProfile2.Inputs("ReferencePoint").Add oClosestEndPoints3.Output("Point2", 1)
        oTubeFromDetailedProfile2.Parameter("Length") = 1#
        oTubeFromDetailedProfile2.Evaluate
        pGeometricConstruction.ControlledInputs("RootPlateSystem1").Clear
        pGeometricConstruction.ControlledInputs("RootPlateSystem1").AddElements oTubeFromDetailedProfile2.ControlledInputs("RootPlateSystem")
        pGeometricConstruction.ControlledInputs("LeafPlateSystem1").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem1").AddElements oTubeFromDetailedProfile2.ControlledInputs("LeafPlateSystem")
        
        'Create SurfaceBodyExtractors for first stiffener
        Set oFacePortExtractor201 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "201-SurfaceBodyExtractor")
        oFacePortExtractor201.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile2
        oFacePortExtractor201.Inputs("CoordinateSystem").Add oCSByLines6c
    
        Set oFacePortExtractor202 = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "202-SurfaceBodyExtractor")
        oFacePortExtractor202.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile2
        oFacePortExtractor202.Inputs("CoordinateSystem").Add oCSByLines6c
    Else
        Set oFacePortExtractor201 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "201-FacePortExtractor")
        oFacePortExtractor201.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oFacePortExtractor201.Inputs("CoordinateSystem").Add oCSByLines6c
        oFacePortExtractor201.Parameter("GeometrySelector") = GCCurrent
    
        Set oFacePortExtractor202 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "202-FacePortExtractor")
        oFacePortExtractor202.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oFacePortExtractor202.Inputs("CoordinateSystem").Add oCSByLines6c
        oFacePortExtractor202.Parameter("GeometrySelector") = GCCurrent
    End If
    
    oFacePortExtractor201.Parameter("FacesContext") = GCLateral
    oFacePortExtractor201.Parameter("IntersectingPlane") = 0
    oFacePortExtractor201.Parameter("SurfaceType") = 1
    oFacePortExtractor201.Parameter("Offset") = 0#
    
    oFacePortExtractor202.Parameter("FacesContext") = GCLateral
    oFacePortExtractor202.Parameter("IntersectingPlane") = 0
    oFacePortExtractor202.Parameter("SurfaceType") = 1
    oFacePortExtractor202.Parameter("Offset") = 0#
    
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1:
            oFacePortExtractor201.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCNear
            oFacePortExtractor202.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCFar
        Case GCCenteredHor2:
            oFacePortExtractor201.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCNear
            oFacePortExtractor202.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCFar
        Case GCTop3:
            oFacePortExtractor201.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCFar
            oFacePortExtractor202.Parameter("LookingAxis") = GCZAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCNear
        Case GCLeft4:
            oFacePortExtractor201.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCFar
            oFacePortExtractor202.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCNear
        Case GCCenteredVer5:
            oFacePortExtractor201.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCFar
            oFacePortExtractor202.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCNear
        Case GCRight6:
            oFacePortExtractor201.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor201.Parameter("TrackFlag") = GCNear
            oFacePortExtractor202.Parameter("LookingAxis") = GCYAxis
            oFacePortExtractor202.Parameter("TrackFlag") = GCFar
    End Select
    
    oFacePortExtractor201.Evaluate
    
    If Not TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        pGeometricConstruction.ControlledInputs("FacePort201").Clear
        pGeometricConstruction.ControlledInputs("FacePort201").AddElements oFacePortExtractor201.ControlledInputs("Port")
    End If
    
    oFacePortExtractor202.Evaluate
    If Not TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        pGeometricConstruction.ControlledInputs("FacePort202").Clear
        pGeometricConstruction.ControlledInputs("FacePort202").AddElements oFacePortExtractor202.ControlledInputs("Port")
    End If
    
    'Create track point from CsByStiffener2
    Dim oTrackPoint203 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
         Set oTrackPoint203 = oGCFactory.CreateEntity("PointFromCS", pPOM, "203-TrackPoint")
         oTrackPoint203.Inputs("CoordinateSystem").Add oCS6
         oTrackPoint203.Parameter("X") = 0
         oTrackPoint203.Parameter("Y") = 0
         oTrackPoint203.Parameter("Z") = 0
         oTrackPoint203.Evaluate
    End If
    
    Dim oEdgePortExtractor203 As SP3DGeometricConstruction.GeometricConstruction
    If Not TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        Set oEdgePortExtractor203 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "203-EdgePortExtractor")
        oEdgePortExtractor203.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oEdgePortExtractor203.Inputs("FacePort").Add oFacePortExtractor201
        oEdgePortExtractor203.Inputs("CoordinateSystem").Add oCSByLines6c
        oEdgePortExtractor203.Parameter("GeometrySelector") = GCCurrent
        oEdgePortExtractor203.Parameter("LookingAxis") = GCXAxis
        oEdgePortExtractor203.Parameter("TrackFlag") = GCFar '1
        oEdgePortExtractor203.Evaluate
        pGeometricConstruction.ControlledInputs("EdgePort2").Clear
        pGeometricConstruction.ControlledInputs("EdgePort2").AddElements oEdgePortExtractor203.ControlledInputs("Port")
    Else
        Set oEdgePortExtractor203 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "203-EdgeFromCurveOrSurf")
        oEdgePortExtractor203.Inputs("Geometry").Add oFacePortExtractor201
        oEdgePortExtractor203.Inputs("TrackPoint").Add oTrackPoint203
        oEdgePortExtractor203.Parameter("TrackFlag") = GCNear
        oEdgePortExtractor203.Evaluate
    End If

    Dim oEdgePortExtractor204 As SP3DGeometricConstruction.GeometricConstruction
    If Not TypeOf pGeometricConstruction.Inputs("Profile2").Item(1) Is IJStiffenerSystem Then
        Set oEdgePortExtractor204 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "204-EdgePortExtractor")
        oEdgePortExtractor204.Inputs("Connectable").Add pGeometricConstruction.Inputs("Profile2").Item(1)
        oEdgePortExtractor204.Inputs("FacePort").Add oFacePortExtractor202
        oEdgePortExtractor204.Inputs("CoordinateSystem").Add oCSByLines6c
        oEdgePortExtractor204.Parameter("GeometrySelector") = GCCurrent
        oEdgePortExtractor204.Parameter("LookingAxis") = GCXAxis
        oEdgePortExtractor204.Parameter("TrackFlag") = GCFar '1
        oEdgePortExtractor204.Evaluate
        pGeometricConstruction.ControlledInputs("EdgePort204").Clear
        pGeometricConstruction.ControlledInputs("EdgePort204").AddElements oEdgePortExtractor204.ControlledInputs("Port")
    Else
        Set oEdgePortExtractor204 = oGCFactory.CreateEntity("EdgeFromCurveOrSurf", pPOM, "204-EdgeFromCurveOrSurf")
        oEdgePortExtractor204.Inputs("Geometry").Add oFacePortExtractor202
        oEdgePortExtractor204.Inputs("TrackPoint").Add oTrackPoint203
        oEdgePortExtractor204.Parameter("TrackFlag") = GCNear
        oEdgePortExtractor204.Evaluate
    End If
    
    Dim oCurveOrientate205 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate205 = oGCFactory.CreateEntity("CurveOrientate", pPOM, "205-CurveOrientate")
    oCurveOrientate205.Inputs("Curve").Add oEdgePortExtractor203
    oCurveOrientate205.Inputs("CoordinateSystem").Add oCSByLines6c
    oCurveOrientate205.Evaluate
    
    Dim oCurveOrientate206 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveOrientate206 = oGCFactory.CreateEntity("CurveOrientate", pPOM, "206-CurveOrientate")
    oCurveOrientate206.Inputs("Curve").Add oEdgePortExtractor204
    oCurveOrientate206.Inputs("CoordinateSystem").Add oCSByLines6c
    oCurveOrientate206.Evaluate
    
    Dim oPointAtCurveStart207 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart207 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "207-PointAtCurveStart")
    oPointAtCurveStart207.Inputs("Curve").Add oCurveOrientate205
    oPointAtCurveStart207.Evaluate
    
    Dim oPointAtCurveStart208 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart208 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "208-PointAtCurveStart")
    oPointAtCurveStart208.Inputs("Curve").Add oCurveOrientate206
    oPointAtCurveStart208.Evaluate
    
    Dim oPointAtCurveEnd209 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd209 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "209-PointAtCurveEnd")
    oPointAtCurveEnd209.Inputs("Curve").Add oCurveOrientate205
    oPointAtCurveEnd209.Evaluate
    
    Dim oPointAtCurveEnd210 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd210 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "210-PointAtCurveEnd")
    oPointAtCurveEnd210.Inputs("Curve").Add oCurveOrientate206
    oPointAtCurveEnd210.Evaluate

    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCTop3, GCLeft4, GCRight6:
        Case GCCenteredHor2, GCCenteredVer5:
            Dim oParamDistMeasureParallel301 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel301 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "301-ParamDistMeasureParallel")
            oParamDistMeasureParallel301.Inputs("Graphics").Add oPointAtCurveStart207
            oParamDistMeasureParallel301.Inputs("Graphics").Add oPointAtCurveStart208
            oParamDistMeasureParallel301.Evaluate

            Dim oLineByPoints306 As SP3DGeometricConstruction.GeometricConstruction
            If oParamDistMeasureParallel301.Parameter("Value") > 0.001 Then
                Dim oLineByPoints302 As SP3DGeometricConstruction.GeometricConstruction
                Set oLineByPoints302 = oGCFactory.CreateEntity("LineByPoints", pPOM, "302-LineByPoints")
                Select Case pGeometricConstruction.Parameter("Location")
                    Case GCCenteredHor2:
                        oLineByPoints302.Inputs("StartPoint").Add oPointAtCurveStart207
                        oLineByPoints302.Inputs("EndPoint").Add oPointAtCurveStart208
                    Case GCCenteredVer5:
                        oLineByPoints302.Inputs("StartPoint").Add oPointAtCurveStart207
                        oLineByPoints302.Inputs("EndPoint").Add oPointAtCurveStart208
                End Select
                oLineByPoints302.Evaluate
        
                Dim oLineByPoints303 As SP3DGeometricConstruction.GeometricConstruction
                Set oLineByPoints303 = oGCFactory.CreateEntity("LineByPoints", pPOM, "303-LineByPoints")
                Select Case pGeometricConstruction.Parameter("Location")
                    Case GCCenteredHor2:
                        oLineByPoints303.Inputs("StartPoint").Add oPointAtCurveEnd209
                        oLineByPoints303.Inputs("EndPoint").Add oPointAtCurveEnd210
                    Case GCCenteredVer5:
                        oLineByPoints303.Inputs("StartPoint").Add oPointAtCurveEnd209
                        oLineByPoints303.Inputs("EndPoint").Add oPointAtCurveEnd210
                End Select
                oLineByPoints303.Evaluate
        
                Dim oPointAtCurveMiddle304 As SP3DGeometricConstruction.GeometricConstruction
                Set oPointAtCurveMiddle304 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "304-PointAtCurveMiddle")
                oPointAtCurveMiddle304.Inputs("Curve").Add oLineByPoints302
                oPointAtCurveMiddle304.Evaluate
        
                Dim oPointAtCurveMiddle305 As SP3DGeometricConstruction.GeometricConstruction
                Set oPointAtCurveMiddle305 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "305-PointAtCurveMiddle")
                oPointAtCurveMiddle305.Inputs("Curve").Add oLineByPoints303
                oPointAtCurveMiddle305.Evaluate
        
                Set oLineByPoints306 = oGCFactory.CreateEntity("LineByPoints", pPOM, "306-LineByPoints")
                oLineByPoints306.Inputs("StartPoint").Add oPointAtCurveMiddle304
                oLineByPoints306.Inputs("EndPoint").Add oPointAtCurveMiddle305
                oLineByPoints306.Evaluate
            Else
                Set oLineByPoints306 = oCurveOrientate205
            End If
    End Select

    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCTop3, GCLeft4, GCRight6:
        Case GCCenteredHor2, GCCenteredVer5:
            Dim oParamDistMeasureParallel307 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel307 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "307-ParamDistMeasureParallel")
            oParamDistMeasureParallel307.Inputs("Graphics").Add oPointAtCurveStart207
            oParamDistMeasureParallel307.Inputs("Graphics").Add oPointAtCurveStart208
            oParamDistMeasureParallel307.Evaluate

            Dim oLineByPoints312 As SP3DGeometricConstruction.GeometricConstruction
            If oParamDistMeasureParallel307.Parameter("Value") > 0.001 Then
                Dim oLineByPoints308 As SP3DGeometricConstruction.GeometricConstruction
                Set oLineByPoints308 = oGCFactory.CreateEntity("LineByPoints", pPOM, "306-LineByPoints")
                Select Case pGeometricConstruction.Parameter("Location")
                    Case GCCenteredHor2:
                        oLineByPoints308.Inputs("StartPoint").Add oPointAtCurveStart116
                        oLineByPoints308.Inputs("EndPoint").Add oPointAtCurveStart117
                    Case GCCenteredVer5:
                        oLineByPoints308.Inputs("StartPoint").Add oPointAtCurveStart116
                        oLineByPoints308.Inputs("EndPoint").Add oPointAtCurveStart117
                End Select
                oLineByPoints308.Evaluate
   
                Dim oLineByPoints309 As SP3DGeometricConstruction.GeometricConstruction
                Set oLineByPoints309 = oGCFactory.CreateEntity("LineByPoints", pPOM, "307-LineByPoints")
                Select Case pGeometricConstruction.Parameter("Location")
                    Case GCCenteredHor2:
                        oLineByPoints309.Inputs("StartPoint").Add oPointAtCurveEnd118
                        oLineByPoints309.Inputs("EndPoint").Add oPointAtCurveEnd119
                     Case GCCenteredVer5:
                        oLineByPoints309.Inputs("StartPoint").Add oPointAtCurveEnd118
                        oLineByPoints309.Inputs("EndPoint").Add oPointAtCurveEnd119
                End Select
                oLineByPoints309.Evaluate
   
                Dim oPointAtCurveMiddle310 As SP3DGeometricConstruction.GeometricConstruction
                Set oPointAtCurveMiddle310 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "308-PointAtCurveMiddle")
                oPointAtCurveMiddle310.Inputs("Curve").Add oLineByPoints308
                oPointAtCurveMiddle310.Evaluate
   
                Dim oPointAtCurveMiddle311 As SP3DGeometricConstruction.GeometricConstruction
                Set oPointAtCurveMiddle311 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "309-PointAtCurveMiddle")
                oPointAtCurveMiddle311.Inputs("Curve").Add oLineByPoints309
                oPointAtCurveMiddle311.Evaluate
   
                Set oLineByPoints312 = oGCFactory.CreateEntity("LineByPoints", pPOM, "310-LineByPoints")
                oLineByPoints312.Inputs("StartPoint").Add oPointAtCurveMiddle310
                oLineByPoints312.Inputs("EndPoint").Add oPointAtCurveMiddle311
                oLineByPoints312.Evaluate
            Else
                Set oLineByPoints312 = oCurveOrientate114
            End If
    End Select
    
    ' capping planes
    Dim oPlaneFromCS401 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS401 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "401-PlaneFromCS")
    oPlaneFromCS401.Inputs("CoordinateSystem").Add oCSByLines5c
    oPlaneFromCS401.Parameter("LookingAxis") = GCXAxis
    oPlaneFromCS401.Parameter("Offset") = 0#
    oPlaneFromCS401.Parameter("Range") = 2#
    oPlaneFromCS401.Evaluate
            
    Dim oPlaneFromCS402 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS402 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "402-PlaneFromCS")
    oPlaneFromCS402.Inputs("CoordinateSystem").Add oCSByLines6c
    oPlaneFromCS402.Parameter("LookingAxis") = GCXAxis
    oPlaneFromCS402.Parameter("Offset") = 0#
    oPlaneFromCS402.Parameter("Range") = 2#
    oPlaneFromCS402.Evaluate
    
    Dim oSurfFromGType403 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType403 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "403-SurfFromGType")
    oSurfFromGType403.Inputs("Surface").Add oPlaneFromCS401
    oSurfFromGType403.Evaluate

    Dim oSurfFromGType404 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType404 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "404-SurfFromGType")
    oSurfFromGType404.Inputs("Surface").Add oPlaneFromCS402
    oSurfFromGType404.Evaluate

    Dim oExtRuledSurface501 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtRuledSurface501 = oGCFactory.CreateEntity("ExtRuledSurface", pPOM, "501-ExtRuledSurface")
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCTop3, GCLeft4, GCRight6:
            oExtRuledSurface501.Inputs("Curve1").Add oCurveOrientate114
            oExtRuledSurface501.Inputs("Curve2").Add oCurveOrientate205
        Case GCCenteredHor2, GCCenteredVer5:
            oExtRuledSurface501.Inputs("Curve1").Add oLineByPoints312
            oExtRuledSurface501.Inputs("Curve2").Add oLineByPoints306
    End Select
    oExtRuledSurface501.Parameter("Extension") = 0.1
    oExtRuledSurface501.Evaluate
    
    Dim oSupport As IJSurfaceBody: Set oSupport = oExtRuledSurface501.Output("Support", 1)
    
'    ' compute a point between the 2 surfaces
    Dim oPointAtSurfaceCenter502 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtSurfaceCenter502 = oGCFactory.CreateEntity("PointAtSurfaceCenter", pPOM, "502-PointAtSurfaceCenter")
    oPointAtSurfaceCenter502.Inputs("Surface").Add oSupport
    oPointAtSurfaceCenter502.Evaluate

    Dim oPointByProjectOnSurf503 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf503 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "503-PointByProjectOnSurf")
    oPointByProjectOnSurf503.Inputs("Point").Add oTrackPoint105
    oPointByProjectOnSurf503.Inputs("Surface").Add oSupport
    oPointByProjectOnSurf503.Evaluate
  
    Dim oParamDistMeasureParallel504 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel504 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "504-ParamDistMeasureParallel")
    oParamDistMeasureParallel504.Inputs("Graphics").Add oPointByProjectOnSurf503
    oParamDistMeasureParallel504.Inputs("Graphics").Add oTrackPoint105
    oParamDistMeasureParallel504.Evaluate

    If oParamDistMeasureParallel504.Parameter("Value") > 0.001 Then
       Dim oLineByPoints504 As SP3DGeometricConstruction.GeometricConstruction
       Set oLineByPoints504 = oGCFactory.CreateEntity("LineByPoints", pPOM, "504-LineByPoints")
       oLineByPoints504.Inputs("StartPoint").Add oPointByProjectOnSurf503
       oLineByPoints504.Inputs("EndPoint").Add oTrackPoint105
       oLineByPoints504.Evaluate
    
       Dim oVectorNormalToSurface505 As SP3DGeometricConstruction.GeometricConstruction
       Set oVectorNormalToSurface505 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "505-VectorNormalToSurface")
       oVectorNormalToSurface505.Inputs("Point").Add oPointByProjectOnSurf503
       oVectorNormalToSurface505.Inputs("Surface").Add oSupport
       oVectorNormalToSurface505.Evaluate
       
       Dim pVector1 As IJDVector: Set pVector1 = Vector_FromLine(oLineByPoints504)
       Dim pVector2 As IJDVector: Set pVector2 = Vector_FromLine(oVectorNormalToSurface505)
       If pVector1.Dot(pVector2) < -0.001 Then
            Dim oSurfReorientate As SP3DGeometricConstruction.GeometricConstruction
            Set oSurfReorientate = oGCFactory.CreateEntity("SurfReorientate", pPOM, "505-oSurfReorientate")
            oSurfReorientate.Inputs("Surface").Add oSupport
            oSurfReorientate.Evaluate
            Set oSupport = oSurfReorientate
       End If
    End If
    
    Dim oCSByCS506 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS506 = oGCFactory.CreateEntity("CSByCS", pPOM, "506-CSByCS")
    oCSByCS506.Inputs("CoordinateSystem").Add oCSByLines5c
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCCenteredHor2, GCTop3:
            oCSByCS506.Parameter("AxesRoles") = GCXY
        Case GCLeft4, GCCenteredVer5, GCRight6:
            oCSByCS506.Parameter("AxesRoles") = GCYZ
    End Select
    oCSByCS506.Evaluate
    
    Dim oCSByCS507 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS507 = oGCFactory.CreateEntity("CSByCS", pPOM, "507-CSByCS")
    oCSByCS507.Inputs("CoordinateSystem").Add oCSByLines6c
    Select Case pGeometricConstruction.Parameter("Location")
        Case GCBottom1, GCCenteredHor2, GCTop3:
            oCSByCS507.Parameter("AxesRoles") = GCXY
        Case GCLeft4, GCCenteredVer5, GCRight6:
            oCSByCS507.Parameter("AxesRoles") = GCYZ
    End Select
    oCSByCS507.Evaluate

    Dim oLateralBoundaries508 As SP3DGeometricConstruction.GeometricConstruction
    Set oLateralBoundaries508 = oGCFactory.CreateEntity("LateralBoundaries", pPOM, "508-LateralBoundaries")
    oLateralBoundaries508.Inputs("CoordinateSystem1").Add oCSByCS506 ' oCSByLines5c
    oLateralBoundaries508.Inputs("PointS1").Add oExtRuledSurface501.Output("PointS1", 1)
    oLateralBoundaries508.Inputs("PointE1").Add oExtRuledSurface501.Output("PointE1", 1)
    oLateralBoundaries508.Inputs("CoordinateSystem2").Add oCSByCS507 'oCSByLines6c
    oLateralBoundaries508.Inputs("PointS2").Add oExtRuledSurface501.Output("PointS2", 1)
    oLateralBoundaries508.Inputs("PointE2").Add oExtRuledSurface501.Output("PointE2", 1)
    oLateralBoundaries508.Parameter("Extension") = 0.1
    oLateralBoundaries508.Evaluate

    Dim oVectorNormalToSurface509 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface509 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "509-VectorNormalToSurface")
    oVectorNormalToSurface509.Inputs("Surface").Add oSupport
    oVectorNormalToSurface509.Inputs("Point").Add oPointAtSurfaceCenter502
    oVectorNormalToSurface509.Parameter("Range") = 10
    oVectorNormalToSurface509.Parameter("Orientation") = 1
    oVectorNormalToSurface509.Evaluate

    Dim oPointAtCurveEnd510 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd510 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "510-PointAtCurveEnd")
    oPointAtCurveEnd510.Inputs("Curve").Add oVectorNormalToSurface509
    oPointAtCurveEnd510.Evaluate

    Dim oCSByCS511 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS511 = oGCFactory.CreateEntity("CSByCS", pPOM, "511-CSByCS")
    oCSByCS511.Inputs("CoordinateSystem").Add oCSByCS506
    oCSByCS511.Inputs("Origin").Add oPointAtSurfaceCenter502
    oCSByCS511.Inputs("TrackPoint").Add oPointAtCurveEnd510
    oCSByCS511.Parameter("AxesRoles") = GCXY
    oCSByCS511.Parameter("TrackFlag") = GCNear
    oCSByCS511.Evaluate
   
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oSupport
    
    Dim bBoundByEdge As Boolean: Let bBoundByEdge = False
    If TypeOf pGeometricConstruction.Input("Profile1") Is ISPSDesignedMember _
    And TypeOf oSupport Is IJPlane _
    And TypeOf oFacePortExtractor101.Output Is IJPlane Then
        If ArePlanesCoplanar(oSupport, oFacePortExtractor101) Then Let bBoundByEdge = True
    End If
    
    If bBoundByEdge Then
        oGCMacro.Output("Boundary", "EdgePort1") = oSurfFromGType403.Output
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "Plane1")
    Else
        oGCMacro.Output("Boundary", "Plane1") = oSurfFromGType403.Output
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "EdgePort1")
    End If
    
    If TypeOf pGeometricConstruction.Input("Profile2") Is ISPSDesignedMember _
    And TypeOf oSupport Is IJPlane _
    And TypeOf oFacePortExtractor201.Output Is IJPlane Then
        If ArePlanesCoplanar(oSupport, oFacePortExtractor201) Then Let bBoundByEdge = True
    End If
        
    If bBoundByEdge Then
         oGCMacro.Output("Boundary", "EdgePort2") = oSurfFromGType404.Output
         Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "Plane2")
    Else
         oGCMacro.Output("Boundary", "Plane2") = oSurfFromGType404.Output
         Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "EdgePort2")
     End If
    
    Dim oBoundary3 As Object: Set oBoundary3 = oLateralBoundaries508.Output("Boundary", 1)
    Dim oBoundary4 As Object: Set oBoundary4 = oLateralBoundaries508.Output("Boundary", 2)
    If Not oFacePortExtractor7 Is Nothing _
    Or Not oFacePortExtractor8 Is Nothing Then
        Dim oPointByProjectOnSurf601 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByProjectOnSurf601 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "601-PointByProjectOnSurf")
        oPointByProjectOnSurf601.Inputs("Point").Add oPointAtSurfaceCenter502
        oPointByProjectOnSurf601.Inputs("Surface").Add oLateralBoundaries508.Output("Boundary", 1)
        oPointByProjectOnSurf601.Evaluate
    
        Dim oPointByProjectOnSurf602 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByProjectOnSurf602 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "602-PointByProjectOnSurf")
        oPointByProjectOnSurf602.Inputs("Point").Add oPointAtSurfaceCenter502
        oPointByProjectOnSurf602.Inputs("Surface").Add oLateralBoundaries508.Output("Boundary", 2)
        oPointByProjectOnSurf602.Evaluate
    
        If Not oFacePortExtractor7 Is Nothing Then
            Dim oCurveByIntersection603 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection603 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "603-CurveByIntersection")
            oCurveByIntersection603.Inputs("Surface1").Add oFacePortExtractor7
            oCurveByIntersection603.Inputs("Surface2").Add oSupport
            On Error Resume Next: Dim lError As Long: lError = 0: oCurveByIntersection603.Evaluate: lError = Err.Number: On Error GoTo 0
            If lError <> 0 Then Err.Raise BOUNDING_PLATE_SYSTEM_NOT_INTERSECTING
            
            Dim oPointByProjectOnSurf604 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointByProjectOnSurf604 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "604-PointByProjectOnSurf")
            oPointByProjectOnSurf604.Inputs("Point").Add oPointAtSurfaceCenter502
            oPointByProjectOnSurf604.Inputs("Surface").Add oFacePortExtractor7
            oPointByProjectOnSurf604.Evaluate
             
            Dim oParamDistMeasureParallel605 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel605 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "605-ParamDistMeasureParallel")
            oParamDistMeasureParallel605.Inputs("Graphics").Add oPointByProjectOnSurf604
            oParamDistMeasureParallel605.Inputs("Graphics").Add oPointByProjectOnSurf601
            oParamDistMeasureParallel605.Evaluate
         
            Dim oParamDistMeasureParallel606 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel606 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "606-ParamDistMeasureParallel")
            oParamDistMeasureParallel606.Inputs("Graphics").Add oPointByProjectOnSurf604
            oParamDistMeasureParallel606.Inputs("Graphics").Add oPointByProjectOnSurf602
            oParamDistMeasureParallel606.Evaluate
            
            Set oBoundary3 = oFacePortExtractor7
            If CDbl(oParamDistMeasureParallel605.Parameter("Value")) < CDbl(oParamDistMeasureParallel606.Parameter("Value")) Then
                Set oBoundary4 = oLateralBoundaries508.Output("Boundary", 2)
            Else
                Set oBoundary4 = oLateralBoundaries508.Output("Boundary", 1)
            End If
        End If
        
        If Not oFacePortExtractor8 Is Nothing Then
            Dim oCurveByIntersection607 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection607 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "607-CurveByIntersection")
            oCurveByIntersection607.Inputs("Surface1").Add oFacePortExtractor8
            oCurveByIntersection607.Inputs("Surface2").Add oSupport
            On Error Resume Next: lError = 0: oCurveByIntersection607.Evaluate: lError = Err.Number: On Error GoTo 0
            If lError <> 0 Then Err.Raise BOUNDING_PLATE_SYSTEM_NOT_INTERSECTING
            
            Dim oPointByProjectOnSurf608 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointByProjectOnSurf608 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "608-PointByProjectOnSurf")
            oPointByProjectOnSurf608.Inputs("Point").Add oPointAtSurfaceCenter502
            oPointByProjectOnSurf608.Inputs("Surface").Add oFacePortExtractor8
            oPointByProjectOnSurf608.Evaluate
        
            Dim oParamDistMeasureParallel609 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel609 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "609-ParamDistMeasureParallel")
            oParamDistMeasureParallel609.Inputs("Graphics").Add oPointByProjectOnSurf608
            oParamDistMeasureParallel609.Inputs("Graphics").Add oPointByProjectOnSurf601
            oParamDistMeasureParallel609.Evaluate
         
            Dim oParamDistMeasureParallel610 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel610 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "610-ParamDistMeasureParallel")
            oParamDistMeasureParallel610.Inputs("Graphics").Add oPointByProjectOnSurf608
            oParamDistMeasureParallel610.Inputs("Graphics").Add oPointByProjectOnSurf602
            oParamDistMeasureParallel610.Evaluate
        
            Set oBoundary4 = oFacePortExtractor8
            If oFacePortExtractor7 Is Nothing Then
                If CDbl(oParamDistMeasureParallel609.Parameter("Value")) < CDbl(oParamDistMeasureParallel610.Parameter("Value")) Then
                    Set oBoundary3 = oLateralBoundaries508.Output("Boundary", 2)
                Else
                    Set oBoundary3 = oLateralBoundaries508.Output("Boundary", 1)
                End If
            End If
        End If
    End If
    
    If TypeOf oBoundary3 Is IJGeometricConstruction Then
        oGCMacro.Output("Boundary", "FacePort3") = oBoundary3
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "Surface3")
    Else
        oGCMacro.Output("Boundary", "Surface3") = oBoundary3
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "FacePort3")
    End If
    
    If TypeOf oBoundary4 Is IJGeometricConstruction Then
        oGCMacro.Output("Boundary", "FacePort4") = oBoundary4
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "Surface4")
    Else
        oGCMacro.Output("Boundary", "Surface4") = oBoundary4
        Call GCMacro_DeleteOutputByKey(oGCMacro, "Boundary", "FacePort4")
    End If
    
    oGCMacro.Output("Node", 1) = oPointAtSurfaceCenter502.Output
    oGCMacro.Output("CoordinateSystem", 1) = oCSByCS511.Output
End Sub
Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function
Function ArePlanesCoplanar(pPlane1 As IJPlane, pPlane2 As IJPlane) As Boolean
    Let ArePlanesCoplanar = False
    Dim pVector1 As IJDVector: Set pVector1 = Vector_FromPlane(pPlane1)
    Dim pVector2 As IJDVector: Set pVector2 = Vector_FromPlane(pPlane2)
    If Abs(pVector1.Dot(pVector2)) > 0.999 Then
'        Call ShowMsg("Vectors are parallel")
        Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromPlane(pPlane1)
        Dim pPosition2 As IJDPosition: Set pPosition2 = Position_FromPlane(pPlane2)
        Dim pVector As IJDVector: Set pVector = pPosition2.Subtract(pPosition1)
        'Let pVector.Length = 1
        Dim dNonCoplanearity As Double
        Let dNonCoplanearity = Abs(pVector.Dot(pVector1))
        If Abs(dNonCoplanearity) < 0.001 Then
            Let ArePlanesCoplanar = True
        End If
    End If
End Function
Public Function Position_FromPlane(pPlane As IJPlane) As IJDPosition
    Dim dX As Double, dY As Double, dZ As Double
    Call pPlane.GetRootPoint(dX, dY, dZ)

    Dim pPosition As New DPosition
    Call pPosition.Set(dX, dY, dZ)

    Set Position_FromPlane = pPosition
End Function
Public Function Vector_FromPlane(pPlane As IJPlane) As IJDVector
    Dim dU As Double, dV As Double, dW As Double
    Call pPlane.GetNormal(dU, dV, dW)

    Dim pVector As New DVector
    Call pVector.Set(dU, dV, dW)

    Set Vector_FromPlane = pVector
End Function
Function GCMacro_GetOutputFromKey(pGCMacro As IJGeometricConstructionMacro, sNameOfOutput As String, sKey As String) As Object
    ' prepare result
    Dim oOutput As Object: Set oOutput = Nothing
    
    Dim i As Integer
    For i = 1 To pGCMacro.Outputs(sNameOfOutput).Count
        Dim oItem As Object: Set oItem = pGCMacro.Outputs(sNameOfOutput)(i)
        If pGCMacro.Outputs(sNameOfOutput).GetKey(oItem) = sKey Then
            Set oOutput = oItem
            Exit For
        End If
    Next
      
    ' return result
    Set GCMacro_GetOutputFromKey = oOutput
End Function
Sub GCMacro_DeleteOutputByKey(pGCMacro As IJGeometricConstructionMacro, sNameOfOutput As String, sKey As String)
    Dim oOutput As Object:
    Set oOutput = GCMacro_GetOutputFromKey(pGCMacro, sNameOfOutput, sKey)
    If Not oOutput Is Nothing Then
        pGCMacro.Outputs(sNameOfOutput).Remove (sKey)
        Dim pObject As IJDObject: Set pObject = oOutput
        pObject.Remove
    End If
End Sub
Function GCMacro(pGC As IJGeometricConstruction) As IJGeometricConstructionMacro
    Set GCMacro = pGC
End Function

'*******************************************************************************************************************
 'Adaptation of the GC for mirror copy or copy symmetry operations.
 '   This sub is called by the GC evaluate method (CGeometricConstruction::Evaluate) when a mirror is in progress
 '   (CTL_FLAG_MIRROR_IN_PROGRESS flag is set on the BO during mirror operations (CGeometricConstruction::Adapt))
 '   At this stage, the GC output (aggregated geometry) has been processed by the mirror (mirrored /adapted or transformed)
 '   and holds the mirrored geometry that we assume being the exact mirror. The inputs of the mirrored GC are also well defined:
 '   mirrored if they were in the copy set, identified, kept as original, or deleted (delete optional) by the user in the adapt form.
 '
 '   The purpose of this adapt method is to adapt the GC parameters so that the next recompute will match this exact mirrored geometry
 '   or be as close as possible.
 '
 '   FAILURE:
 '   We assume here that any failure or exception is caught by the caller which is the GC BO.
 '
 '   Adaptation of GCTransitionPlateGap GC:
 '      Mirror the normal of the macro 'support' output and look for the closest axes in the coordinate system of the new 'Profile1'.
 '      The 'Location' is adapted accordingly.
 '
Private Sub IJGCMirror_Adapt(ByVal pOriginalGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirroredGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirrorPlane As IngrGeom3D.IJPlane, ByVal pMatrix As AutoMath.IJDT4x4, ByVal pGCInputsFromInitialCopySet As IMSCoreCollections.IJElements)
    
    'Get the original GCmacro 'CoordinateSystem'
    Dim pOriginalMacroCoordinateSystem As IJLocalCoordinateSystem: Set pOriginalMacroCoordinateSystem = pOriginalGC.Output("CoordinateSystem", 1)
    
    Dim pPOMForDebug As IJDPOM
    If GetGlobalShowDetails() Then Set pPOMForDebug = Object_GetPOM(pOriginalMacroCoordinateSystem)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    'Compute normal to support output on original GC
    Dim oVectorNormalToSurface505 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface505 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOMForDebug, "505-VectorNormalToSurface")
    oVectorNormalToSurface505.Inputs("Point").Add pOriginalGC.Output("Node", 1)
    oVectorNormalToSurface505.Inputs("Surface").Add pOriginalGC.Output("Support", 1)
    oVectorNormalToSurface505.Evaluate
    
    Dim oNormal As IJDVector: Set oNormal = Vector_FromLine(oVectorNormalToSurface505)
    
    'mirror this normal
    Dim pMirroredSupportNormal As IJDVector: Set pMirroredSupportNormal = pMatrix.TransformVector(oNormal)
    
    'Get 'Profile1' coordinate system of mirrored GC
    Dim oAxis1 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pMirroredGC.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oAxis1 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "001-AxisPortExtractor")
        oAxis1.Inputs("Connectable").Add pMirroredGC.Inputs("Profile1").Item(1)
        oAxis1.Parameter("GeometrySelector") = GCCurrent
        oAxis1.Evaluate
    Else
        Set oAxis1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOMForDebug, "001-LineAxisPortExtractor")
        oAxis1.Inputs("MemberPart").Add pMirroredGC.Inputs("Profile1").Item(1)
        oAxis1.Evaluate
    End If

    Dim oAxis2 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pMirroredGC.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Set oAxis2 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "002-AxisPortExtractor")
        oAxis2.Inputs("Connectable").Add pMirroredGC.Inputs("Profile2").Item(1)
        oAxis2.Parameter("GeometrySelector") = GCCurrent
        oAxis2.Evaluate
    Else
        Set oAxis2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOMForDebug, "002-LineAxisPortExtractor")
        oAxis2.Inputs("MemberPart").Add pMirroredGC.Inputs("Profile2").Item(1)
        oAxis2.Evaluate
    End If
    
    Dim oClosestEndPoints3 As SP3DGeometricConstruction.GeometricConstruction
    Set oClosestEndPoints3 = oGCFactory.CreateEntity("PointsAtClosestEndsOfCurves", pPOMForDebug, "003-PointsAtClosestEndsOfCurves")
    oClosestEndPoints3.Inputs("Curve1").Add oAxis1
    oClosestEndPoints3.Inputs("Curve2").Add oAxis2
    oClosestEndPoints3.Evaluate
    
    Dim oCS5 As SP3DGeometricConstruction.GeometricConstruction
    If TypeOf pMirroredGC.Inputs("Profile1").Item(1) Is IJStiffenerSystem Then
        Dim oPointAlongCurve5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve5 = oGCFactory.CreateEntity("PointAlongCurve", pPOMForDebug, "005a-PointAlongCurve")
        oPointAlongCurve5.Inputs("Curve").Add oAxis1
        oPointAlongCurve5.Inputs("Point").Add oClosestEndPoints3.Output("Point1", 1)
        oPointAlongCurve5.Inputs("TrackPoint").Add oClosestEndPoints3.Output("Point2", 1)
        oPointAlongCurve5.Parameter("Distance") = 0.001
        oPointAlongCurve5.Parameter("TrackFlag") = GCFar
        oPointAlongCurve5.Evaluate
        
        Dim oVectorTangentToCurve5 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve5 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOMForDebug, "005-VectorTangentToCurve")
        oVectorTangentToCurve5.Inputs("Curve").Add oAxis1
        oVectorTangentToCurve5.Inputs("Point").Add oPointAlongCurve5
        oVectorTangentToCurve5.Parameter("Range") = 1#
        oVectorTangentToCurve5.Evaluate
        
        Dim oPlaneByPointNormal5 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal5 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOMForDebug, "005-PlaneByPointNormal")
        oPlaneByPointNormal5.Inputs("Point").Add oPointAlongCurve5
        oPlaneByPointNormal5.Inputs("Line").Add oVectorTangentToCurve5
        oPlaneByPointNormal5.Parameter("Range") = 5#
        oPlaneByPointNormal5.Evaluate
        
        Set oCS5 = oGCFactory.CreateEntity("CSByStiffener", pPOMForDebug, "005-CSByStiffener")
        oCS5.Inputs("Stiffener").Add pMirroredGC.Inputs("Profile1").Item(1)
        oCS5.Inputs("Plane").Add oPlaneByPointNormal5 'oClosestEndPoints3.Output("Point1", 1)
        oCS5.Evaluate
    Else
        Set oCS5 = oGCFactory.CreateEntity("CSFromMember", pPOMForDebug, "005-CSFromMember")
        oCS5.Inputs("MemberPart").Add pMirroredGC.Inputs("Profile1").Item(1)
        oCS5.Inputs("Point").Add oClosestEndPoints3.Output("Point1", 1)
        oCS5.Evaluate
    End If
    
    'Find pMirroredSupportNormal direction in 'Profile1' coordinate system
    Dim eOrientation As eGlobalTrackFlag
    Dim bRes As Boolean
    Dim pNewCs1 As IJLocalCoordinateSystem: Set pNewCs1 = oCS5.Output
    
    bRes = Orientation_FromCSVectors(pMirroredSupportNormal, pNewCs1.XAxis, pNewCs1.YAxis, pNewCs1.ZAxis, eOrientation)
    
    'if mirrored support normal along x, it is an invalid support regarding to new members, keep original location
    Select Case eOrientation
        Case GCMinusY:
'            MsgBox "mirrored normal along -Y =>Left or centervert"
            If pOriginalGC.Parameter("Location") = GCCenteredHor2 Or pOriginalGC.Parameter("Location") = GCCenteredVer5 Then
                pMirroredGC.Parameter("Location") = GCCenteredVer5
            Else
                pMirroredGC.Parameter("Location") = GCLeft4
            End If
        Case GCPlusY:
            ' if original centered, keep centered
'            MsgBox "mirrored normal along +Y =>Rigth or centervert"
            If pOriginalGC.Parameter("Location") = GCCenteredHor2 Or pOriginalGC.Parameter("Location") = GCCenteredVer5 Then
                pMirroredGC.Parameter("Location") = GCCenteredVer5
            Else
                pMirroredGC.Parameter("Location") = GCRight6
            End If
         Case GCMinusZ:
'            MsgBox "mirrored normal along -Z =>Top or centerHori"
            If pOriginalGC.Parameter("Location") = GCCenteredHor2 Or pOriginalGC.Parameter("Location") = GCCenteredVer5 Then
                pMirroredGC.Parameter("Location") = GCCenteredHor2
            Else
                pMirroredGC.Parameter("Location") = GCTop3
            End If
        Case GCPlusZ:
            ' if original centered, keep centered
'            MsgBox "mirrored normal along +Z =>Bottom or centerHori"
            If pOriginalGC.Parameter("Location") = GCCenteredHor2 Or pOriginalGC.Parameter("Location") = GCCenteredVer5 Then
                pMirroredGC.Parameter("Location") = GCCenteredHor2
            Else
                pMirroredGC.Parameter("Location") = GCBottom1
            End If
    End Select

End Sub

'Find which axis vector (ortho vectors) is colinear to the given vector, if none is found return the nearest axis
'each axis can be Null to only compare 2 axes
Private Function Orientation_FromCSVectors(pVectorToOrient As IJDVector, pGCVectorOfAxis1 As IJDVector, pGCVectorOfAxis2 As IJDVector, pGCVectorOfAxis3 As IJDVector, ByRef eOrientation As eGlobalTrackFlag) As Boolean
    Dim dDot1 As Double
    Dim dDot2 As Double
    Dim dDot3 As Double
    Dim dDotMax As Double
    Dim bIsMaxSet As Boolean
    
    bIsMaxSet = False
    Orientation_FromCSVectors = False
    
    Set pVectorToOrient = Vector_Normalize(pVectorToOrient)
    
    If Not pGCVectorOfAxis1 Is Nothing Then
        Set pGCVectorOfAxis1 = Vector_Normalize(pGCVectorOfAxis1)
        dDot1 = pGCVectorOfAxis1.Dot(pVectorToOrient)
        If (1 - Abs(dDot1)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusX
            If dDot1 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusX
            Orientation_FromCSVectors = True
            Exit Function
        Else
            dDotMax = dDot1
            bIsMaxSet = True
            eOrientation = eGlobalTrackFlag.GCPlusX
        End If
    End If
    If Not pGCVectorOfAxis2 Is Nothing Then
        Set pGCVectorOfAxis2 = Vector_Normalize(pGCVectorOfAxis2)
        dDot2 = pGCVectorOfAxis2.Dot(pVectorToOrient)
        If (1 - Abs(dDot2)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusY
            If dDot2 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusY
            Orientation_FromCSVectors = True
            Exit Function
        Else
            If bIsMaxSet = False Or Abs(dDot2) - Abs(dDotMax) > EPSILON Then
                dDotMax = dDot2
                bIsMaxSet = True
                eOrientation = eGlobalTrackFlag.GCPlusY
            End If
        End If
    End If
    If Not pGCVectorOfAxis3 Is Nothing Then
        Set pGCVectorOfAxis3 = Vector_Normalize(pGCVectorOfAxis3)
        dDot3 = pGCVectorOfAxis3.Dot(pVectorToOrient)
        If (1 - Abs(dDot3)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusZ
            If dDot3 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusZ
            Orientation_FromCSVectors = True
            Exit Function
        Else
            If bIsMaxSet = False Or Abs(dDot3) - Abs(dDotMax) > EPSILON Then
                dDotMax = dDot3
                bIsMaxSet = True
                eOrientation = eGlobalTrackFlag.GCPlusZ
            End If
        End If
    End If
  
    If dDotMax < -EPSILON Then
        eOrientation = eOrientation + 1
    End If

End Function

Private Function GetGlobalShowDetails() As Long
    Dim oToolshelper As New PORTHELPERLib.ToolsHelper
    GetGlobalShowDetails = oToolshelper.GetGlobalShowDetails()
End Function

Private Function Object_GetPOM(ByVal pObject As IngrGeom3D.IJDObject) As IJDPOM
    On Error Resume Next
    Object_GetPOM = pObject.ResourceManager
    On Error GoTo 0
End Function


Private Function Vector_Normalize(pVector As DVector) As IJDVector
    ' create new vector
    Dim pVectorNew As IJDVector
    Set pVectorNew = pVector.Clone
    Dim dLength As Double
    
    dLength = pVector.Length
    If dLength < EPSILON Then
        Exit Function
    End If
    ' set coordinates
    Call pVectorNew.Set(pVector.x / dLength, pVector.y / dLength, pVector.z / dLength)
    
    ' return result
    Set Vector_Normalize = pVectorNew
End Function
