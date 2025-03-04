VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GussetAtIntersection"
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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select Braces", "ISPSMemberPartPrismatic OR ISPSDesignedMember", 2, 4, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersSplitEachOther+MembersBoundBySurfaces+MembersPenetrateByPlate", 0)
'    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select Braces", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
'    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersSplitEachOther+MembersBoundBySurfaces+MembersPenetrateByPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_2", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_4_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_29_1")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "TrimmingBackBoundary")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddParameter("Cutback", "Cutback", 8, 1, 59, 0, 0, 0.25)
    Call pGeometricConstructionDefinition.AddParameter("Overlap", "Overlap", 8, 1, 59, 0, 0, 0.25)
    Call pGeometricConstructionDefinition.AddParameter("Weldtoe", "Weldtoe", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("SnipeAngle", "SnipeAngle", 8, 2, 78, 0, 0, 30#)
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oOrderAxes21 As SP3DGeometricConstruction.GeometricConstruction
    Set oOrderAxes21 = oGCFactory.CreateEntity("OrderAxes2", pPOM, "001-OrderAxes2")
    
    'Do not use AddElements that keeps keys (migrated profiles will not have good iIndexOfInput)
    Dim pElementsOfPrimaryProfiles As IJElements: Set pElementsOfPrimaryProfiles = pGeometricConstruction.Inputs("PrimaryProfile1")
    Dim i As Integer
    For i = 1 To pElementsOfPrimaryProfiles.Count
        oOrderAxes21.Inputs("PrimaryProfiles").Add pElementsOfPrimaryProfiles(i), CStr(i)
    Next
    oOrderAxes21.Parameter("Support") = 2
    oOrderAxes21.Evaluate
    pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_1").AddElements oOrderAxes21.ControlledInputs("Port1") '.Item("1")
    pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_2").AddElements oOrderAxes21.ControlledInputs("Port2") '("Port1").Item("2")

    Dim pGCMacroOfOrderAxes2 As IJGeometricConstructionMacro: Set pGCMacroOfOrderAxes2 = oOrderAxes21
    Dim pElementsOfAxes  As IJElements: Set pElementsOfAxes = pGCMacroOfOrderAxes2.Outputs("Axis")
    Dim iIndexOfInput(1 To 4) As Integer
    For i = 1 To 4
        Dim sKeyOfAxis As String: Let sKeyOfAxis = pElementsOfAxes.GetKey(pElementsOfAxes(i))
        Debug.Print "KeyOfAxis(" + CStr(i) + ")= " + sKeyOfAxis
        iIndexOfInput(i) = Val(Mid(sKeyOfAxis, 1, 1))
    Next
    
    Dim oCSByLines2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines2 = oGCFactory.CreateEntity("CSByLines", pPOM, "002-CSByLines")
    oCSByLines2.Inputs("AxisLine1").Add oOrderAxes21.Output("Axis", 1)
    oCSByLines2.Inputs("AxisLine2").Add oOrderAxes21.Output("Axis", 2)
    oCSByLines2.Parameter("AxesRoles") = 1
    oCSByLines2.Parameter("CSOrientation") = 1
    oCSByLines2.Parameter("TrackFlag") = 1
    oCSByLines2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Weldtoe")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Overlap")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("Cutback")
    oParamDistConstant5.Evaluate

    Dim oCSByLines6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines6 = oGCFactory.CreateEntity("CSByLines", pPOM, "006-CSByLines")
    oCSByLines6.Inputs("AxisLine1").Add oOrderAxes21.Output("Axis", 2)
    oCSByLines6.Inputs("AxisLine2").Add oOrderAxes21.Output("Axis", 3)
    oCSByLines6.Parameter("AxesRoles") = 1
    oCSByLines6.Parameter("CSOrientation") = 1
    oCSByLines6.Parameter("TrackFlag") = 1
    oCSByLines6.Evaluate

    Dim oCSByLines7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines7 = oGCFactory.CreateEntity("CSByLines", pPOM, "007-CSByLines")
    oCSByLines7.Inputs("AxisLine1").Add oOrderAxes21.Output("Axis", 3)
    oCSByLines7.Inputs("AxisLine2").Add oOrderAxes21.Output("Axis", 4)
    oCSByLines7.Parameter("AxesRoles") = 1
    oCSByLines7.Parameter("CSOrientation") = 1
    oCSByLines7.Parameter("TrackFlag") = 1
    oCSByLines7.Evaluate

    Dim oCSByLines8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines8 = oGCFactory.CreateEntity("CSByLines", pPOM, "008-CSByLines")
    oCSByLines8.Inputs("AxisLine1").Add oOrderAxes21.Output("Axis", 4)
    oCSByLines8.Inputs("AxisLine2").Add oOrderAxes21.Output("Axis", 1)
    oCSByLines8.Parameter("AxesRoles") = 1
    oCSByLines8.Parameter("CSOrientation") = 1
    oCSByLines8.Parameter("TrackFlag") = 1
    oCSByLines8.Evaluate

    Dim oPointAlongCurve9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve9 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "009-PointAlongCurve")
    oPointAlongCurve9.Inputs("Curve").Add oOrderAxes21.Output("Axis", 1)
    oPointAlongCurve9.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve9.Parameter("TrackFlag") = 1
    oPointAlongCurve9.Evaluate

    Dim oLineFromCS10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS10 = oGCFactory.CreateEntity("LineFromCS", pPOM, "010-LineFromCS")
    oLineFromCS10.Inputs("CoordinateSystem").Add oCSByLines2
    oLineFromCS10.Parameter("LookingAxis") = 3
    oLineFromCS10.Parameter("Length") = -0.5
    oLineFromCS10.Parameter("CSOrientation") = 1
    oLineFromCS10.Parameter("LineJustification") = 1
    oLineFromCS10.Evaluate

    Dim oPointAlongCurve11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "011-PointAlongCurve")
    oPointAlongCurve11.Inputs("Curve").Add oOrderAxes21.Output("Axis", 2)
    oPointAlongCurve11.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve11.Parameter("TrackFlag") = 1
    oPointAlongCurve11.Evaluate

    Dim oPointAlongCurve12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve12 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "012-PointAlongCurve")
    oPointAlongCurve12.Inputs("Curve").Add oOrderAxes21.Output("Axis", 3)
    oPointAlongCurve12.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve12.Parameter("TrackFlag") = 1
    oPointAlongCurve12.Evaluate

    Dim oPointAlongCurve13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve13 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "013-PointAlongCurve")
    oPointAlongCurve13.Inputs("Curve").Add oOrderAxes21.Output("Axis", 4)
    oPointAlongCurve13.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve13.Parameter("TrackFlag") = 1
    oPointAlongCurve13.Evaluate

    Dim oCSByCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS14 = oGCFactory.CreateEntity("CSByCS", pPOM, "014-CSByCS")
    oCSByCS14.Inputs("CoordinateSystem").Add oCSByLines2
    oCSByCS14.Inputs("Origin").Add oPointAlongCurve9
    oCSByCS14.Parameter("AxesRoles") = 1
    oCSByCS14.Parameter("CSOrientation") = 1
    oCSByCS14.Parameter("TrackFlag") = 1
    oCSByCS14.Evaluate

    Dim oPointAlongCurve15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve15 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "015-PointAlongCurve")
    oPointAlongCurve15.Inputs("Curve").Add oOrderAxes21.Output("Axis", 1)
    oPointAlongCurve15.Inputs("Point").Add oPointAlongCurve9
    oPointAlongCurve15.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve15.Parameter("TrackFlag") = 1
    oPointAlongCurve15.Evaluate

    Dim oCSByCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS16 = oGCFactory.CreateEntity("CSByCS", pPOM, "016-CSByCS")
    oCSByCS16.Inputs("CoordinateSystem").Add oCSByLines2
    oCSByCS16.Inputs("Origin").Add oPointAlongCurve15
    oCSByCS16.Parameter("AxesRoles") = 1
    oCSByCS16.Parameter("CSOrientation") = 1
    oCSByCS16.Parameter("TrackFlag") = 1
    oCSByCS16.Evaluate

    Dim oFacePortExtractor17 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor17 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "017-FacePortExtractor")
    oFacePortExtractor17.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(iIndexOfInput(1)), "1"
    oFacePortExtractor17.Inputs("CoordinateSystem").Add oCSByCS16
    oFacePortExtractor17.Parameter("GeometrySelector") = 4
    oFacePortExtractor17.Parameter("FacesContext") = 2
    oFacePortExtractor17.Parameter("LookingAxis") = 2
    oFacePortExtractor17.Parameter("IntersectingPlane") = 0
    oFacePortExtractor17.Parameter("SurfaceType") = 1
    oFacePortExtractor17.Parameter("TrackFlag") = 2
    oFacePortExtractor17.Parameter("Offset") = 0#
    oFacePortExtractor17.Evaluate
    pGeometricConstruction.ControlledInputs("Port_4_1").Clear
    pGeometricConstruction.ControlledInputs("Port_4_1").Add oFacePortExtractor17.ControlledInputs("Port").Item("1")

    Dim oLineFromCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS18 = oGCFactory.CreateEntity("LineFromCS", pPOM, "018-LineFromCS")
    oLineFromCS18.Inputs("CoordinateSystem").Add oCSByCS16
    oLineFromCS18.Parameter("LookingAxis") = 2
    oLineFromCS18.Parameter("Length") = 1#
    oLineFromCS18.Parameter("CSOrientation") = 2
    oLineFromCS18.Parameter("LineJustification") = 1
    oLineFromCS18.Evaluate

    Dim oLineFromCS19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS19 = oGCFactory.CreateEntity("LineFromCS", pPOM, "019-LineFromCS")
    oLineFromCS19.Inputs("CoordinateSystem").Add oCSByCS16
    oLineFromCS19.Parameter("LookingAxis") = 2
    oLineFromCS19.Parameter("Length") = 1#
    oLineFromCS19.Parameter("CSOrientation") = 1
    oLineFromCS19.Parameter("LineJustification") = 1
    oLineFromCS19.Evaluate

    Dim oPointAlongCurve20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve20 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "020-PointAlongCurve")
    oPointAlongCurve20.Inputs("Curve").Add oOrderAxes21.Output("Axis", 2)
    oPointAlongCurve20.Inputs("Point").Add oPointAlongCurve11
    oPointAlongCurve20.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve20.Parameter("TrackFlag") = 1
    oPointAlongCurve20.Evaluate

    Dim oPointAlongCurve21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve21 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "021-PointAlongCurve")
    oPointAlongCurve21.Inputs("Curve").Add oOrderAxes21.Output("Axis", 3)
    oPointAlongCurve21.Inputs("Point").Add oPointAlongCurve12
    oPointAlongCurve21.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve21.Parameter("TrackFlag") = 1
    oPointAlongCurve21.Evaluate

    Dim oPointAlongCurve22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve22 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "022-PointAlongCurve")
    oPointAlongCurve22.Inputs("Curve").Add oOrderAxes21.Output("Axis", 4)
    oPointAlongCurve22.Inputs("Point").Add oPointAlongCurve13
    oPointAlongCurve22.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve22.Parameter("TrackFlag") = 1
    oPointAlongCurve22.Evaluate

    Dim oCSByCS23 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS23 = oGCFactory.CreateEntity("CSByCS", pPOM, "023-CSByCS")
    oCSByCS23.Inputs("CoordinateSystem").Add oCSByLines6
    oCSByCS23.Inputs("Origin").Add oPointAlongCurve20
    oCSByCS23.Parameter("AxesRoles") = 1
    oCSByCS23.Parameter("CSOrientation") = 1
    oCSByCS23.Parameter("TrackFlag") = 1
    oCSByCS23.Evaluate

    Dim oCSByCS24 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS24 = oGCFactory.CreateEntity("CSByCS", pPOM, "024-CSByCS")
    oCSByCS24.Inputs("CoordinateSystem").Add oCSByLines7
    oCSByCS24.Inputs("Origin").Add oPointAlongCurve21
    oCSByCS24.Parameter("AxesRoles") = 1
    oCSByCS24.Parameter("CSOrientation") = 1
    oCSByCS24.Parameter("TrackFlag") = 1
    oCSByCS24.Evaluate

    Dim oCSByCS25 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS25 = oGCFactory.CreateEntity("CSByCS", pPOM, "025-CSByCS")
    oCSByCS25.Inputs("CoordinateSystem").Add oCSByLines8
    oCSByCS25.Inputs("Origin").Add oPointAlongCurve22
    oCSByCS25.Parameter("AxesRoles") = 1
    oCSByCS25.Parameter("CSOrientation") = 1
    oCSByCS25.Parameter("TrackFlag") = 1
    oCSByCS25.Evaluate

    Dim oFacePortExtractor26 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor26 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "026-FacePortExtractor")
    oFacePortExtractor26.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(iIndexOfInput(2)), "1"
    oFacePortExtractor26.Inputs("CoordinateSystem").Add oCSByCS23
    oFacePortExtractor26.Parameter("GeometrySelector") = 4
    oFacePortExtractor26.Parameter("FacesContext") = 2
    oFacePortExtractor26.Parameter("LookingAxis") = 2
    oFacePortExtractor26.Parameter("IntersectingPlane") = 0
    oFacePortExtractor26.Parameter("SurfaceType") = 1
    oFacePortExtractor26.Parameter("TrackFlag") = 2
    oFacePortExtractor26.Parameter("Offset") = 0#
    oFacePortExtractor26.Evaluate
    pGeometricConstruction.ControlledInputs("Port_29_1").Clear
    pGeometricConstruction.ControlledInputs("Port_29_1").Add oFacePortExtractor26.ControlledInputs("Port").Item("1")

    Dim oLineFromCS27 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS27 = oGCFactory.CreateEntity("LineFromCS", pPOM, "027-LineFromCS")
    oLineFromCS27.Inputs("CoordinateSystem").Add oCSByCS23
    oLineFromCS27.Parameter("LookingAxis") = 2
    oLineFromCS27.Parameter("Length") = 1#
    oLineFromCS27.Parameter("CSOrientation") = 2
    oLineFromCS27.Parameter("LineJustification") = 1
    oLineFromCS27.Evaluate

    Dim oLineFromCS28 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS28 = oGCFactory.CreateEntity("LineFromCS", pPOM, "028-LineFromCS")
    oLineFromCS28.Inputs("CoordinateSystem").Add oCSByCS23
    oLineFromCS28.Parameter("LookingAxis") = 2
    oLineFromCS28.Parameter("Length") = 1#
    oLineFromCS28.Parameter("CSOrientation") = 1
    oLineFromCS28.Parameter("LineJustification") = 1
    oLineFromCS28.Evaluate

    Dim oFacePortExtractor29 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor29 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "029-FacePortExtractor")
    oFacePortExtractor29.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(iIndexOfInput(3)), "1"
    oFacePortExtractor29.Inputs("CoordinateSystem").Add oCSByCS24
    oFacePortExtractor29.Parameter("GeometrySelector") = 4
    oFacePortExtractor29.Parameter("FacesContext") = 2
    oFacePortExtractor29.Parameter("LookingAxis") = 2
    oFacePortExtractor29.Parameter("IntersectingPlane") = 0
    oFacePortExtractor29.Parameter("SurfaceType") = 1
    oFacePortExtractor29.Parameter("TrackFlag") = 2
    oFacePortExtractor29.Parameter("Offset") = 0#
    oFacePortExtractor29.Evaluate
    pGeometricConstruction.ControlledInputs("Port_4_1").Clear
    pGeometricConstruction.ControlledInputs("Port_4_1").Add oFacePortExtractor29.ControlledInputs("Port").Item("1")

    Dim oLineFromCS30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS30 = oGCFactory.CreateEntity("LineFromCS", pPOM, "030-LineFromCS")
    oLineFromCS30.Inputs("CoordinateSystem").Add oCSByCS24
    oLineFromCS30.Parameter("LookingAxis") = 2
    oLineFromCS30.Parameter("Length") = 1#
    oLineFromCS30.Parameter("CSOrientation") = 2
    oLineFromCS30.Parameter("LineJustification") = 1
    oLineFromCS30.Evaluate

    Dim oLineFromCS31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS31 = oGCFactory.CreateEntity("LineFromCS", pPOM, "031-LineFromCS")
    oLineFromCS31.Inputs("CoordinateSystem").Add oCSByCS24
    oLineFromCS31.Parameter("LookingAxis") = 2
    oLineFromCS31.Parameter("Length") = 1#
    oLineFromCS31.Parameter("CSOrientation") = 1
    oLineFromCS31.Parameter("LineJustification") = 1
    oLineFromCS31.Evaluate

    Dim oFacePortExtractor32 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor32 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "032-FacePortExtractor")
    oFacePortExtractor32.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(iIndexOfInput(4)), "1"
    oFacePortExtractor32.Inputs("CoordinateSystem").Add oCSByCS25
    oFacePortExtractor32.Parameter("GeometrySelector") = 4
    oFacePortExtractor32.Parameter("FacesContext") = 2
    oFacePortExtractor32.Parameter("LookingAxis") = 2
    oFacePortExtractor32.Parameter("IntersectingPlane") = 0
    oFacePortExtractor32.Parameter("SurfaceType") = 1
    oFacePortExtractor32.Parameter("TrackFlag") = 2
    oFacePortExtractor32.Parameter("Offset") = 0#
    oFacePortExtractor32.Evaluate
    pGeometricConstruction.ControlledInputs("Port_29_1").Clear
    pGeometricConstruction.ControlledInputs("Port_29_1").Add oFacePortExtractor32.ControlledInputs("Port").Item("1")

    Dim oLineFromCS33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS33 = oGCFactory.CreateEntity("LineFromCS", pPOM, "033-LineFromCS")
    oLineFromCS33.Inputs("CoordinateSystem").Add oCSByCS25
    oLineFromCS33.Parameter("LookingAxis") = 2
    oLineFromCS33.Parameter("Length") = 1#
    oLineFromCS33.Parameter("CSOrientation") = 2
    oLineFromCS33.Parameter("LineJustification") = 1
    oLineFromCS33.Evaluate

    Dim oLineFromCS34 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS34 = oGCFactory.CreateEntity("LineFromCS", pPOM, "034-LineFromCS")
    oLineFromCS34.Inputs("CoordinateSystem").Add oCSByCS25
    oLineFromCS34.Parameter("LookingAxis") = 2
    oLineFromCS34.Parameter("Length") = 1#
    oLineFromCS34.Parameter("CSOrientation") = 1
    oLineFromCS34.Parameter("LineJustification") = 1
    oLineFromCS34.Evaluate

    Dim oPointByCurveAndSurface35 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface35 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "035-PointByCurveAndSurface")
    oPointByCurveAndSurface35.Inputs("Curve1").Add oLineFromCS27
    oPointByCurveAndSurface35.Inputs("Surface2").Add oFacePortExtractor26
    oPointByCurveAndSurface35.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface35.Evaluate

    Dim oPointByCurveAndSurface36 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface36 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "036-PointByCurveAndSurface")
    oPointByCurveAndSurface36.Inputs("Curve1").Add oLineFromCS28
    oPointByCurveAndSurface36.Inputs("Surface2").Add oFacePortExtractor26
    oPointByCurveAndSurface36.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface36.Evaluate

    Dim oPointByCurveAndSurface37 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface37 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "037-PointByCurveAndSurface")
    oPointByCurveAndSurface37.Inputs("Curve1").Add oLineFromCS33
    oPointByCurveAndSurface37.Inputs("Surface2").Add oFacePortExtractor32
    oPointByCurveAndSurface37.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface37.Evaluate

    Dim oPointByCurveAndSurface38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface38 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "038-PointByCurveAndSurface")
    oPointByCurveAndSurface38.Inputs("Curve1").Add oLineFromCS34
    oPointByCurveAndSurface38.Inputs("Surface2").Add oFacePortExtractor32
    oPointByCurveAndSurface38.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface38.Evaluate

    Dim oPointByCurveAndSurface39 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface39 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "039-PointByCurveAndSurface")
    oPointByCurveAndSurface39.Inputs("Curve1").Add oLineFromCS30
    oPointByCurveAndSurface39.Inputs("Surface2").Add oFacePortExtractor29
    oPointByCurveAndSurface39.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface39.Evaluate

    Dim oPointByCurveAndSurface40 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface40 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "040-PointByCurveAndSurface")
    oPointByCurveAndSurface40.Inputs("Curve1").Add oLineFromCS31
    oPointByCurveAndSurface40.Inputs("Surface2").Add oFacePortExtractor29
    oPointByCurveAndSurface40.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface40.Evaluate

    Dim oPlaneFromCS41 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS41 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "041-PlaneFromCS")
    oPlaneFromCS41.Inputs("CoordinateSystem").Add oCSByCS14
    oPlaneFromCS41.Parameter("LookingAxis") = 2
    oPlaneFromCS41.Parameter("Offset") = 0#
    oPlaneFromCS41.Parameter("Range") = 0.1
    oPlaneFromCS41.Evaluate

    Dim oLineFromCS42 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS42 = oGCFactory.CreateEntity("LineFromCS", pPOM, "042-LineFromCS")
    oLineFromCS42.Inputs("CoordinateSystem").Add oCSByCS14
    oLineFromCS42.Parameter("LookingAxis") = 1
    oLineFromCS42.Parameter("Length") = 1#
    oLineFromCS42.Parameter("CSOrientation") = 1
    oLineFromCS42.Parameter("LineJustification") = 0
    oLineFromCS42.Evaluate

    Dim oPointFromCS43 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS43 = oGCFactory.CreateEntity("PointFromCS", pPOM, "043-PointFromCS")
    oPointFromCS43.Inputs("CoordinateSystem").Add oCSByCS14
    oPointFromCS43.Parameter("X") = 0#
    oPointFromCS43.Parameter("Y") = 0#
    oPointFromCS43.Parameter("Z") = 0.1
    oPointFromCS43.Evaluate

    Dim oLineAtAngle44 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle44 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "044-LineAtAngle")
    oLineAtAngle44.Inputs("Line").Add oLineFromCS42
    oLineAtAngle44.Inputs("Point").Add oPointAlongCurve9
    oLineAtAngle44.Inputs("TrackPoint").Add oPointFromCS43
    oLineAtAngle44.Parameter("Angle") = pGeometricConstruction.Parameter("SnipeAngle")
    oLineAtAngle44.Parameter("TrackFlag") = 2
    oLineAtAngle44.Evaluate

    Dim oLineAtAngle45 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle45 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "045-LineAtAngle")
    oLineAtAngle45.Inputs("Line").Add oLineFromCS42
    oLineAtAngle45.Inputs("Point").Add oPointAlongCurve9
    oLineAtAngle45.Inputs("TrackPoint").Add oPointFromCS43
    oLineAtAngle45.Parameter("Angle") = pGeometricConstruction.Parameter("SnipeAngle")
    oLineAtAngle45.Parameter("TrackFlag") = 1
    oLineAtAngle45.Evaluate

    Dim oPointByCurveAndSurface46 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface46 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "046-PointByCurveAndSurface")
    oPointByCurveAndSurface46.Inputs("Curve1").Add oLineFromCS18
    oPointByCurveAndSurface46.Inputs("Surface2").Add oFacePortExtractor17
    oPointByCurveAndSurface46.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface46.Evaluate

    Dim oPointByCurveAndSurface47 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface47 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "047-PointByCurveAndSurface")
    oPointByCurveAndSurface47.Inputs("Curve1").Add oLineFromCS19
    oPointByCurveAndSurface47.Inputs("Surface2").Add oFacePortExtractor17
    oPointByCurveAndSurface47.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface47.Evaluate

    Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "048-PointAlongCurve")
    oPointAlongCurve48.Inputs("Curve").Add oLineFromCS18
    oPointAlongCurve48.Inputs("Point").Add oPointByCurveAndSurface46
    oPointAlongCurve48.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve48.Parameter("TrackFlag") = 1
    oPointAlongCurve48.Evaluate

    Dim oPointAlongCurve49 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
    oPointAlongCurve49.Inputs("Curve").Add oLineFromCS19
    oPointAlongCurve49.Inputs("Point").Add oPointByCurveAndSurface47
    oPointAlongCurve49.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve49.Parameter("TrackFlag") = 1
    oPointAlongCurve49.Evaluate

    Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
    oPointAlongCurve50.Inputs("Curve").Add oLineFromCS27
    oPointAlongCurve50.Inputs("Point").Add oPointByCurveAndSurface35
    oPointAlongCurve50.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve50.Parameter("TrackFlag") = 1
    oPointAlongCurve50.Evaluate

    Dim oPointAlongCurve51 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve51 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "051-PointAlongCurve")
    oPointAlongCurve51.Inputs("Curve").Add oLineFromCS28
    oPointAlongCurve51.Inputs("Point").Add oPointByCurveAndSurface36
    oPointAlongCurve51.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve51.Parameter("TrackFlag") = 1
    oPointAlongCurve51.Evaluate

    Dim oPointAlongCurve52 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve52 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "052-PointAlongCurve")
    oPointAlongCurve52.Inputs("Curve").Add oLineFromCS30
    oPointAlongCurve52.Inputs("Point").Add oPointByCurveAndSurface39
    oPointAlongCurve52.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve52.Parameter("TrackFlag") = 1
    oPointAlongCurve52.Evaluate

    Dim oPointAlongCurve53 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "053-PointAlongCurve")
    oPointAlongCurve53.Inputs("Curve").Add oLineFromCS31
    oPointAlongCurve53.Inputs("Point").Add oPointByCurveAndSurface40
    oPointAlongCurve53.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve53.Parameter("TrackFlag") = 1
    oPointAlongCurve53.Evaluate

    Dim oPointAlongCurve54 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve54 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "054-PointAlongCurve")
    oPointAlongCurve54.Inputs("Curve").Add oLineFromCS33
    oPointAlongCurve54.Inputs("Point").Add oPointByCurveAndSurface37
    oPointAlongCurve54.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve54.Parameter("TrackFlag") = 1
    oPointAlongCurve54.Evaluate

    Dim oPointAlongCurve55 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve55 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "055-PointAlongCurve")
    oPointAlongCurve55.Inputs("Curve").Add oLineFromCS34
    oPointAlongCurve55.Inputs("Point").Add oPointByCurveAndSurface38
    oPointAlongCurve55.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve55.Parameter("TrackFlag") = 1
    oPointAlongCurve55.Evaluate

    Dim oCpxStringByCurves56 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves56 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "056-CpxStringByCurves")
    oCpxStringByCurves56.Inputs("Curves").Add oLineAtAngle44
    oCpxStringByCurves56.Inputs("Curves").Add oLineAtAngle45
    oCpxStringByCurves56.Evaluate

    Dim oSurfByRevolution57 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByRevolution57 = oGCFactory.CreateEntity("SurfByRevolution", pPOM, "057-SurfByRevolution")
    oSurfByRevolution57.Inputs("PlanarCrossSection").Add oCpxStringByCurves56
    oSurfByRevolution57.Inputs("RotationAxis").Add oLineFromCS10
    oSurfByRevolution57.Evaluate

    Dim oLineStringByPoints58 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineStringByPoints58 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "058-LineStringByPoints")
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve48
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve49
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve50
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve51
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve52
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve53
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve54
    oLineStringByPoints58.Inputs("Points").Add oPointAlongCurve55
    oLineStringByPoints58.Parameter("ClosedFlag") = 1
    oLineStringByPoints58.Evaluate

    Dim oSurfByLinearExtrusion59 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion59 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "059-SurfByLinearExtrusion")
    oSurfByLinearExtrusion59.Inputs("PlanarCrossSection").Add oLineStringByPoints58
    oSurfByLinearExtrusion59.Inputs("ExtrusionLine").Add oLineFromCS10
    oSurfByLinearExtrusion59.Evaluate

    Dim oCpxStringByCurves60 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves60 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "060-CpxStringByCurves")
    oCpxStringByCurves60.Inputs("Curves").Add oLineStringByPoints58
    oCpxStringByCurves60.Evaluate

    Dim oPlaneByBoundaries61 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByBoundaries61 = oGCFactory.CreateEntity("PlaneByBoundaries", pPOM, "061-PlaneByBoundaries")
    oPlaneByBoundaries61.Inputs("OuterBoundary").Add oCpxStringByCurves60
    oPlaneByBoundaries61.Evaluate

    Dim oSurfFromGType62 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType62 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "062-SurfFromGType")
    oSurfFromGType62.Inputs("Surface").Add oPlaneByBoundaries61
    oSurfFromGType62.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oSurfFromGType62.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion59.Output
    oGCMacro.Output("TrimmingBackBoundary", 1) = oSurfByRevolution57.Output
    oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
    oGCMacro.Output("CoordinateSystem", 1) = oOrderAxes21.Output("CoordinateSystem", 1)
End Sub
