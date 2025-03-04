VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GussetAtCorner"
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
    COMMON_NODE_NOT_FOUND = 2
    MEMBER_AXIS_TOO_SMALL = 3
    MEMBER_AXES_COLINEAR = 4
    MEMBER_AXES_NOT_COPLANAR = 5
    MEMBER_FACES_NOT_PARALLEL = 6
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
    Let Source = "GCAPSWithTrimBack3.GenericPlateImp"
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

    Dim oAPS As Object:
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Set oAPS = pGC.ControlledInputs("AdvancedPlateSystem")(1)
        Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = oAPS
        pPlateMoldedConventions.plateThicknessDirection = Centered
    End If
    
    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.ProcessError(pGC, Err.Number)
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
' implementation of the IJGCMirror interface
'
Private Sub IJGCMirror_Adapt(ByVal pGCOfOriginalGC As IJGeometricConstruction, ByVal pGCOfMirroredGC As IJGeometricConstruction, _
                             ByVal pPlaneOfMirrorPlane As IJPlane, ByVal pT4x4OfMirrorTransformation As IJDT4x4, _
                             ByVal pElementsOfGCsFromInitalCopySet As IJElements)

    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMirror_Adapt"))

    Call m_oGenericPlateHelper.Adapt(pGCOfOriginalGC, pGCOfMirroredGC, _
                                     pPlaneOfMirrorPlane, pT4x4OfMirrorTransformation, _
                                     pElementsOfGCsFromInitalCopySet, _
                                     "CoordinateSystem", "Location")

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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select a Beam", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile2", "Select a Column", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile3", "Select a Brace", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile3", "BehaviorOfPrimaryProfile3", 1, 0, 0, 0, 0, "MembersBoundBySurfaces+MembersPenetrateByPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_2_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_11_1")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "TrimmingBackBoundary")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddParameter("Cutback", "Cutback", 8, 1, 59, 0, 0, 0.25)
    Call pGeometricConstructionDefinition.AddParameter("Overlap", "Overlap", 8, 1, 59, 0, 0, 0.25)
    Call pGeometricConstructionDefinition.AddParameter("Weldtoe", "Weldtoe", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Angle1", "Delta angle 1", 8, 2, 78, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("Angle2", "Delta angle 2", 8, 2, 78, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("SnipeAngle", "SnipeAngle", 8, 2, 78, 0, 0, 30#)
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "1"
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port_1_1").Clear
    pGeometricConstruction.ControlledInputs("Port_1_1").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile3").Item(1), "1"
    oLineAxisPortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("Port_2_1").Clear
    pGeometricConstruction.ControlledInputs("Port_2_1").Add oLineAxisPortExtractor2.ControlledInputs("Port").Item("1")

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Cutback")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Overlap")
    oParamDistConstant4.Evaluate

    Dim oParamDistMeasureLength5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength5 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "005-ParamDistMeasureLength")
    oParamDistMeasureLength5.Inputs("Curve").Add oLineAxisPortExtractor1
    oParamDistMeasureLength5.Parameter("Value") = 5#
    oParamDistMeasureLength5.Evaluate

    Dim oPointAtCurveMiddle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle6 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "006-PointAtCurveMiddle")
    oPointAtCurveMiddle6.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveMiddle6.Evaluate

    Dim oCSByLines7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines7 = oGCFactory.CreateEntity("CSByLines", pPOM, "007-CSByLines")
    oCSByLines7.Inputs("AxisLine1").Add oLineAxisPortExtractor1
    oCSByLines7.Inputs("AxisLine2").Add oLineAxisPortExtractor2
    oCSByLines7.Parameter("AxesRoles") = 1
    oCSByLines7.Parameter("CSOrientation") = 1
    oCSByLines7.Parameter("TrackFlag") = 1
    oCSByLines7.Evaluate

    Dim oCSByLines8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines8 = oGCFactory.CreateEntity("CSByLines", pPOM, "008-CSByLines")
    oCSByLines8.Inputs("AxisLine1").Add oLineAxisPortExtractor2
    oCSByLines8.Inputs("AxisLine2").Add oLineAxisPortExtractor1
    oCSByLines8.Parameter("AxesRoles") = 1
    oCSByLines8.Parameter("CSOrientation") = 1
    oCSByLines8.Parameter("TrackFlag") = 1
    oCSByLines8.Evaluate

    Dim oPointAtCurveMiddle9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle9 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "009-PointAtCurveMiddle")
    oPointAtCurveMiddle9.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveMiddle9.Evaluate

    Dim oLineOrientate10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineOrientate10 = oGCFactory.CreateEntity("LineOrientate", pPOM, "010-LineOrientate")
    oLineOrientate10.Inputs("Line").Add oLineAxisPortExtractor2
    oLineOrientate10.Inputs("CoordinateSystem").Add oCSByLines8
    oLineOrientate10.Evaluate

    Dim oCSByCS11 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS11 = oGCFactory.CreateEntity("CSByCS", pPOM, "011-CSByCS")
    oCSByCS11.Inputs("CoordinateSystem").Add oCSByLines7
    oCSByCS11.Inputs("Origin").Add oPointAtCurveMiddle6
    oCSByCS11.Parameter("AxesRoles") = 1
    oCSByCS11.Parameter("CSOrientation") = 1
    oCSByCS11.Parameter("TrackFlag") = 1
    oCSByCS11.Evaluate

    Dim oLineFromCS12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS12 = oGCFactory.CreateEntity("LineFromCS", pPOM, "012-LineFromCS")
    oLineFromCS12.Inputs("CoordinateSystem").Add oCSByLines7
    oLineFromCS12.Parameter("LookingAxis") = 3
    oLineFromCS12.Parameter("Length") = -0.5
    oLineFromCS12.Parameter("CSOrientation") = 1
    oLineFromCS12.Parameter("LineJustification") = 1
    oLineFromCS12.Evaluate

    Dim oPlaneFromCS13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS13 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "013-PlaneFromCS")
    oPlaneFromCS13.Inputs("CoordinateSystem").Add oCSByLines7
    oPlaneFromCS13.Parameter("LookingAxis") = 3
    oPlaneFromCS13.Parameter("Offset") = 0#
    oPlaneFromCS13.Parameter("Range") = oParamDistMeasureLength5
    oPlaneFromCS13.Evaluate

    Dim oPointFromCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS14 = oGCFactory.CreateEntity("PointFromCS", pPOM, "014-PointFromCS")
    oPointFromCS14.Inputs("CoordinateSystem").Add oCSByLines7
    oPointFromCS14.Parameter("X") = 0#
    oPointFromCS14.Parameter("Y") = 0#
    oPointFromCS14.Parameter("Z") = 0#
    oPointFromCS14.Evaluate

    Dim oFacePortExtractor15 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor15 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "015-FacePortExtractor")
    oFacePortExtractor15.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
    oFacePortExtractor15.Inputs("CoordinateSystem").Add oCSByCS11
    oFacePortExtractor15.Parameter("GeometrySelector") = 4
    oFacePortExtractor15.Parameter("FacesContext") = 2
    oFacePortExtractor15.Parameter("LookingAxis") = 1
    oFacePortExtractor15.Parameter("IntersectingPlane") = 0
    oFacePortExtractor15.Parameter("SurfaceType") = 1
    oFacePortExtractor15.Parameter("TrackFlag") = 2
    oFacePortExtractor15.Parameter("Offset") = 0#
    oFacePortExtractor15.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort1").Clear
    pGeometricConstruction.ControlledInputs("FacePort1").Add oFacePortExtractor15.ControlledInputs("Port").Item("1")

    Dim oFacePortExtractor16 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "016-FacePortExtractor")
    oFacePortExtractor16.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "2"
    oFacePortExtractor16.Inputs("CoordinateSystem").Add oCSByCS11
    oFacePortExtractor16.Parameter("GeometrySelector") = 4
    oFacePortExtractor16.Parameter("FacesContext") = 2
    oFacePortExtractor16.Parameter("LookingAxis") = 2
    oFacePortExtractor16.Parameter("IntersectingPlane") = 0
    oFacePortExtractor16.Parameter("SurfaceType") = 1
    oFacePortExtractor16.Parameter("TrackFlag") = 2
    oFacePortExtractor16.Parameter("Offset") = 0#
    oFacePortExtractor16.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort2").Clear
    pGeometricConstruction.ControlledInputs("FacePort2").Add oFacePortExtractor16.ControlledInputs("Port").Item("1")

    Dim oCurveByIntersection17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection17 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "017-CurveByIntersection")
    oCurveByIntersection17.Inputs("Surface1").Add oPlaneFromCS13
    oCurveByIntersection17.Inputs("Surface2").Add oFacePortExtractor15
    oCurveByIntersection17.Inputs("TrackPoint").Add oPointAtCurveMiddle9
    oCurveByIntersection17.Parameter("TrackFlag") = 1
    oCurveByIntersection17.Parameter("ConstructionSurface") = 0
    oCurveByIntersection17.Evaluate

    Dim oCurveByIntersection18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection18 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "018-CurveByIntersection")
    oCurveByIntersection18.Inputs("Surface1").Add oPlaneFromCS13
    oCurveByIntersection18.Inputs("Surface2").Add oFacePortExtractor16
    oCurveByIntersection18.Inputs("TrackPoint").Add oPointAtCurveMiddle9
    oCurveByIntersection18.Parameter("TrackFlag") = 1
    oCurveByIntersection18.Parameter("ConstructionSurface") = 0
    oCurveByIntersection18.Evaluate

    Dim oCSByCS19 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS19 = oGCFactory.CreateEntity("CSByCS", pPOM, "019-CSByCS")
    oCSByCS19.Inputs("CoordinateSystem").Add oCSByLines8
    oCSByCS19.Inputs("Origin").Add oPointAtCurveMiddle9
    oCSByCS19.Parameter("AxesRoles") = 1
    oCSByCS19.Parameter("CSOrientation") = 1
    oCSByCS19.Parameter("TrackFlag") = 1
    oCSByCS19.Evaluate

    Dim oFacePortExtractor20 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor20 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "020-FacePortExtractor")
    oFacePortExtractor20.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile3").Item(1), "2"
    oFacePortExtractor20.Inputs("CoordinateSystem").Add oCSByCS19
    oFacePortExtractor20.Parameter("GeometrySelector") = 4
    oFacePortExtractor20.Parameter("FacesContext") = 2
    oFacePortExtractor20.Parameter("LookingAxis") = 2
    oFacePortExtractor20.Parameter("IntersectingPlane") = 0
    oFacePortExtractor20.Parameter("SurfaceType") = 1
    oFacePortExtractor20.Parameter("TrackFlag") = 2
    oFacePortExtractor20.Parameter("Offset") = 0#
    oFacePortExtractor20.Evaluate
    pGeometricConstruction.ControlledInputs("Port_11_1").Clear
    pGeometricConstruction.ControlledInputs("Port_11_1").Add oFacePortExtractor20.ControlledInputs("Port").Item("1")

    Dim oPointByCurves21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves21 = oGCFactory.CreateEntity("PointByCurves", pPOM, "021-PointByCurves")
    oPointByCurves21.Inputs("Curve1").Add oCurveByIntersection18
    oPointByCurves21.Inputs("Curve2").Add oCurveByIntersection17
    oPointByCurves21.Parameter("TrackFlag") = 1
    oPointByCurves21.Evaluate

    Dim oParamDistFunctionWith2Args22 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args22 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "022-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args22.Inputs("Arg1").Add oParamDistConstant3
    oParamDistFunctionWith2Args22.Inputs("Arg2").Add oParamDistConstant4
    oParamDistFunctionWith2Args22.Parameter("Value") = 0.75
    oParamDistFunctionWith2Args22.Parameter("Operation") = 1
    oParamDistFunctionWith2Args22.Evaluate

    Dim oPointAtCurveStart23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart23 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "023-PointAtCurveStart")
    oPointAtCurveStart23.Inputs("Curve").Add oLineOrientate10
    oPointAtCurveStart23.Evaluate

    Dim oPointAlongCurve24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve24 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "024-PointAlongCurve")
    oPointAlongCurve24.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAlongCurve24.Inputs("Point").Add oPointByCurves21
    oPointAlongCurve24.Inputs("TrackPoint").Add oPointAtCurveMiddle9
    oPointAlongCurve24.Parameter("Distance") = oParamDistFunctionWith2Args22
    oPointAlongCurve24.Parameter("TrackFlag") = 1
    oPointAlongCurve24.Evaluate

    Dim oCurveByIntersection25 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection25 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "025-CurveByIntersection")
    oCurveByIntersection25.Inputs("Surface1").Add oPlaneFromCS13
    oCurveByIntersection25.Inputs("Surface2").Add oFacePortExtractor20
    oCurveByIntersection25.Inputs("TrackPoint").Add oPointAtCurveMiddle6
    oCurveByIntersection25.Parameter("TrackFlag") = 1
    oCurveByIntersection25.Parameter("ConstructionSurface") = 0
    oCurveByIntersection25.Evaluate

    Dim oCurveByIntersection26 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection26 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "026-CurveByIntersection")
    oCurveByIntersection26.Inputs("Surface1").Add oPlaneFromCS13
    oCurveByIntersection26.Inputs("Surface2").Add oFacePortExtractor20
    oCurveByIntersection26.Inputs("TrackPoint").Add oPointAtCurveMiddle6
    oCurveByIntersection26.Parameter("TrackFlag") = 2
    oCurveByIntersection26.Parameter("ConstructionSurface") = 0
    oCurveByIntersection26.Evaluate

    Dim oPointAlongCurve27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve27 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "027-PointAlongCurve")
    oPointAlongCurve27.Inputs("Curve").Add oCurveByIntersection25
    oPointAlongCurve27.Inputs("Point").Add oPointAlongCurve24
    oPointAlongCurve27.Parameter("Distance") = 0#
    oPointAlongCurve27.Parameter("TrackFlag") = 1
    oPointAlongCurve27.Evaluate

    Dim oPointAlongCurve28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve28 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "028-PointAlongCurve")
    oPointAlongCurve28.Inputs("Curve").Add oCurveByIntersection26
    oPointAlongCurve28.Inputs("Point").Add oPointAlongCurve24
    oPointAlongCurve28.Parameter("Distance") = 0#
    oPointAlongCurve28.Parameter("TrackFlag") = 1
    oPointAlongCurve28.Evaluate

    Dim oLineByPoints29 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints29 = oGCFactory.CreateEntity("LineByPoints", pPOM, "029-LineByPoints")
    oLineByPoints29.Inputs("StartPoint").Add oPointAlongCurve27
    oLineByPoints29.Inputs("EndPoint").Add oPointAlongCurve28
    oLineByPoints29.Evaluate

    Dim oPointAlongCurve30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve30 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "030-PointAlongCurve")
    oPointAlongCurve30.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAlongCurve30.Inputs("Point").Add oPointByCurves21
    oPointAlongCurve30.Inputs("TrackPoint").Add oPointAlongCurve24
    oPointAlongCurve30.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve30.Parameter("TrackFlag") = 1
    oPointAlongCurve30.Evaluate

    Dim oCSByCS31 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS31 = oGCFactory.CreateEntity("CSByCS", pPOM, "031-CSByCS")
    oCSByCS31.Inputs("CoordinateSystem").Add oCSByLines8
    oCSByCS31.Inputs("Origin").Add oPointAlongCurve30
    oCSByCS31.Parameter("AxesRoles") = 1
    oCSByCS31.Parameter("CSOrientation") = 1
    oCSByCS31.Parameter("TrackFlag") = 1
    oCSByCS31.Evaluate

    Dim oPointAlongCurve32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve32 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "032-PointAlongCurve")
    oPointAlongCurve32.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve32.Inputs("Point").Add oPointAlongCurve27
    oPointAlongCurve32.Inputs("TrackPoint").Add oPointAlongCurve24
    oPointAlongCurve32.Parameter("Distance") = pGeometricConstruction.Parameter("Weldtoe")
    oPointAlongCurve32.Parameter("TrackFlag") = 2
    oPointAlongCurve32.Evaluate

    Dim oPointAlongCurve33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve33 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "033-PointAlongCurve")
    oPointAlongCurve33.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve33.Inputs("Point").Add oPointAlongCurve28
    oPointAlongCurve33.Inputs("TrackPoint").Add oPointAlongCurve24
    oPointAlongCurve33.Parameter("Distance") = pGeometricConstruction.Parameter("Weldtoe")
    oPointAlongCurve33.Parameter("TrackFlag") = 2
    oPointAlongCurve33.Evaluate

    Dim oLineFromCS34 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS34 = oGCFactory.CreateEntity("LineFromCS", pPOM, "034-LineFromCS")
    oLineFromCS34.Inputs("CoordinateSystem").Add oCSByCS31
    oLineFromCS34.Parameter("LookingAxis") = 1
    oLineFromCS34.Parameter("Length") = 1#
    oLineFromCS34.Parameter("CSOrientation") = 1
    oLineFromCS34.Parameter("LineJustification") = 1
    oLineFromCS34.Evaluate

    Dim oPlaneFromCS35 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS35 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "035-PlaneFromCS")
    oPlaneFromCS35.Inputs("CoordinateSystem").Add oCSByCS31
    oPlaneFromCS35.Parameter("LookingAxis") = 2
    oPlaneFromCS35.Parameter("Offset") = 0#
    oPlaneFromCS35.Parameter("Range") = 0.1
    oPlaneFromCS35.Evaluate

    Dim oPointFromCS36 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS36 = oGCFactory.CreateEntity("PointFromCS", pPOM, "036-PointFromCS")
    oPointFromCS36.Inputs("CoordinateSystem").Add oCSByCS31
    oPointFromCS36.Parameter("X") = 0#
    oPointFromCS36.Parameter("Y") = 0#
    oPointFromCS36.Parameter("Z") = 0.1
    oPointFromCS36.Evaluate

    Dim oLineAtAngle37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle37 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "037-LineAtAngle")
    oLineAtAngle37.Inputs("Line").Add oCurveByIntersection18
    oLineAtAngle37.Inputs("Point").Add oPointAlongCurve33
    oLineAtAngle37.Inputs("TrackPoint").Add oPointByCurves21
    oLineAtAngle37.Parameter("Angle") = pGeometricConstruction.Parameter("Angle1")
    oLineAtAngle37.Parameter("TrackFlag") = 1
    oLineAtAngle37.Evaluate

    Dim oLineAtAngle38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle38 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "038-LineAtAngle")
    oLineAtAngle38.Inputs("Line").Add oCurveByIntersection17
    oLineAtAngle38.Inputs("Point").Add oPointAlongCurve32
    oLineAtAngle38.Inputs("TrackPoint").Add oPointByCurves21
    oLineAtAngle38.Parameter("Angle") = pGeometricConstruction.Parameter("Angle2")
    oLineAtAngle38.Parameter("TrackFlag") = 1
    oLineAtAngle38.Evaluate

    Dim oPointByCurves39 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves39 = oGCFactory.CreateEntity("PointByCurves", pPOM, "039-PointByCurves")
    oPointByCurves39.Inputs("Curve1").Add oLineAtAngle38
    oPointByCurves39.Inputs("Curve2").Add oLineAtAngle37
    oPointByCurves39.Parameter("TrackFlag") = 1
    oPointByCurves39.Evaluate

    Dim oLineAtAngle40 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle40 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "040-LineAtAngle")
    oLineAtAngle40.Inputs("Line").Add oLineFromCS34
    oLineAtAngle40.Inputs("Point").Add oPointAlongCurve30
    oLineAtAngle40.Inputs("Plane").Add oPlaneFromCS35
    oLineAtAngle40.Inputs("TrackPoint").Add oPointFromCS36
    oLineAtAngle40.Parameter("Angle") = pGeometricConstruction.Parameter("SnipeAngle")
    oLineAtAngle40.Parameter("TrackFlag") = 1
    oLineAtAngle40.Evaluate

    Dim oLineAtAngle41 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle41 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "041-LineAtAngle")
    oLineAtAngle41.Inputs("Line").Add oLineFromCS34
    oLineAtAngle41.Inputs("Point").Add oPointAlongCurve30
    oLineAtAngle41.Inputs("Plane").Add oPlaneFromCS35
    oLineAtAngle41.Inputs("TrackPoint").Add oPointFromCS36
    oLineAtAngle41.Parameter("Angle") = pGeometricConstruction.Parameter("SnipeAngle")
    oLineAtAngle41.Parameter("TrackFlag") = 2
    oLineAtAngle41.Evaluate

    Dim oCpxStringByCurves42 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves42 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "042-CpxStringByCurves")
    oCpxStringByCurves42.Inputs("Curves").Add oLineAtAngle40
    oCpxStringByCurves42.Inputs("Curves").Add oLineAtAngle41
    oCpxStringByCurves42.Evaluate

    Dim oPointByCurveAndSurface43 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface43 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "043-PointByCurveAndSurface")
    oPointByCurveAndSurface43.Inputs("Curve1").Add oLineAtAngle37
    oPointByCurveAndSurface43.Inputs("Surface2").Add oFacePortExtractor15
    oPointByCurveAndSurface43.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface43.Evaluate

    Dim oPointByCurveAndSurface44 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface44 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "044-PointByCurveAndSurface")
    oPointByCurveAndSurface44.Inputs("Curve1").Add oLineAtAngle38
    oPointByCurveAndSurface44.Inputs("Surface2").Add oFacePortExtractor16
    oPointByCurveAndSurface44.Inputs("TrackPoint").Add oPointAlongCurve32
    oPointByCurveAndSurface44.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface44.Evaluate

    Dim oParamDistMeasureParallel45 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel45 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "045-ParamDistMeasureParallel")
    oParamDistMeasureParallel45.Inputs("Graphics").Add oPointByCurves39
    oParamDistMeasureParallel45.Inputs("Graphics").Add oPointByCurves21
    oParamDistMeasureParallel45.Parameter("Value") = 1.025081
    oParamDistMeasureParallel45.Evaluate

    Dim oLineByPoints46 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints46 = oGCFactory.CreateEntity("LineByPoints", pPOM, "046-LineByPoints")
    oLineByPoints46.Inputs("StartPoint").Add oPointByCurves39
    oLineByPoints46.Inputs("EndPoint").Add oPointByCurves21
    oLineByPoints46.Evaluate

    Dim oPointAlongCurve47 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve47 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "047-PointAlongCurve")
    oPointAlongCurve47.Inputs("Curve").Add oLineAtAngle38
    oPointAlongCurve47.Inputs("Point").Add oPointByCurveAndSurface44
    oPointAlongCurve47.Inputs("TrackPoint").Add oPointAlongCurve32
    oPointAlongCurve47.Parameter("Distance") = 0.1
    oPointAlongCurve47.Parameter("TrackFlag") = 2
    oPointAlongCurve47.Evaluate

    Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "048-PointAlongCurve")
    oPointAlongCurve48.Inputs("Curve").Add oLineAtAngle37
    oPointAlongCurve48.Inputs("Point").Add oPointByCurveAndSurface43
    oPointAlongCurve48.Inputs("TrackPoint").Add oPointAlongCurve33
    oPointAlongCurve48.Parameter("Distance") = 0.1
    oPointAlongCurve48.Parameter("TrackFlag") = 2
    oPointAlongCurve48.Evaluate

    Dim oSurfByRevolution49 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByRevolution49 = oGCFactory.CreateEntity("SurfByRevolution", pPOM, "049-SurfByRevolution")
    oSurfByRevolution49.Inputs("PlanarCrossSection").Add oCpxStringByCurves42
    oSurfByRevolution49.Inputs("RotationAxis").Add oLineFromCS12
    oSurfByRevolution49.Evaluate

    Dim oPointAtCurveMiddle50 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle50 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "050-PointAtCurveMiddle")
    oPointAtCurveMiddle50.Inputs("Curve").Add oLineByPoints46
    oPointAtCurveMiddle50.Evaluate

    Dim oCSByCS51 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS51 = oGCFactory.CreateEntity("CSByCS", pPOM, "051-CSByCS")
    oCSByCS51.Inputs("CoordinateSystem").Add oCSByLines7
    oCSByCS51.Inputs("Origin").Add oPointAtCurveMiddle50
    oCSByCS51.Parameter("AxesRoles") = 1
    oCSByCS51.Parameter("CSOrientation") = 1
    oCSByCS51.Parameter("TrackFlag") = 1
    oCSByCS51.Evaluate

    Dim oLineStringByPoints52 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineStringByPoints52 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "052-LineStringByPoints")
    oLineStringByPoints52.Inputs("Points").Add oPointAlongCurve47
    oLineStringByPoints52.Inputs("Points").Add oPointAlongCurve32
    oLineStringByPoints52.Inputs("Points").Add oPointAlongCurve33
    oLineStringByPoints52.Inputs("Points").Add oPointAlongCurve48
    oLineStringByPoints52.Parameter("ClosedFlag") = 0
    oLineStringByPoints52.Evaluate

    Dim oSurfByLinearExtrusion53 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion53 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "053-SurfByLinearExtrusion")
    oSurfByLinearExtrusion53.Inputs("PlanarCrossSection").Add oLineStringByPoints52
    oSurfByLinearExtrusion53.Inputs("ExtrusionLine").Add oLineFromCS12
    oSurfByLinearExtrusion53.Evaluate

    Dim oPlaneFromCS54 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS54 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "054-PlaneFromCS")
    oPlaneFromCS54.Inputs("CoordinateSystem").Add oCSByCS51
    oPlaneFromCS54.Parameter("LookingAxis") = 3
    oPlaneFromCS54.Parameter("Offset") = 0#
    oPlaneFromCS54.Parameter("Range") = oParamDistMeasureParallel45
    oPlaneFromCS54.Evaluate

    Dim oSurfFromGType55 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType55 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "055-SurfFromGType")
    oSurfFromGType55.Inputs("Surface").Add oPlaneFromCS54
    oSurfFromGType55.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oSurfFromGType55.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion53.Output
    oGCMacro.Output("Boundary", "FacePort1") = oFacePortExtractor15.Output
    oGCMacro.Output("Boundary", "FacePort2") = oFacePortExtractor16.Output
    oGCMacro.Output("TrimmingBackBoundary", 1) = oSurfByRevolution49.Output
    oGCMacro.Output("CoordinateSystem", 1) = oCSByLines7.Output
    oGCMacro.Output("Node", 1) = oPointAtCurveStart23.Output
End Sub
