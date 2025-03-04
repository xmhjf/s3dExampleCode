VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlateOneSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
'
'   File:           BucklPlateOneSide.cls
'   ProgID:         GCBucklingPlate.BucklPlateOneSide
'   Author:         Ramesh Babu Kandukuri
'   Creation Date:  May 25, 2012
'   Description:
'       This macro places an advanced plate on either a standard
'       member or on a built up member.
'       Generate connectoions automatically.
'
'   History:
'   Ramesh              May 25, 2012        Created
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oRegularAPSHelper As Object
Private Const TOLERANCE As Double = 0.0001
'
' definition of the error codes
'
Private Enum Errors
    MISSING_MEMBER = 1
    INVALID_OFFSET = 2
    INVALID_HEIGHT = 3
    INVALID_MEMBER = 4
    INVALID_SIDE = 5
    INVALID_WELDTOE1 = 6
    INVALID_WELDTOE2 = 7
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
    Let Source = "GCBucklingPlate.BucklingPlateOneSide"
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
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a MemberPart", "ISPSMemberPartLinear or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("Port", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_Web")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_TopFlange")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_BottomFlange")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 0.2)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe1", "WeldToe1", 8, 1, 59, 0, 0, 0.04)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe2", "WeldToe2", 8, 1, 59, 0, 0, 0.015)
    Call pGeometricConstructionDefinition.AddParameter("Height", "Height", 8, 1, 59, 0, 0, 0.3)
    Call pGeometricConstructionDefinition.AddParameter("TrackFlag", "Near/Far", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Near", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Far", 2)
    pGeometricConstructionDefinition.Behavior = 0
    
    'Add Errors
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER, "Missing Member", "Mising member part for the buckling plate.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_OFFSET, "Invalid Offset", "Offset should be within the member part length")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_HEIGHT, "Invalid Height", "Height should be within the member part depth")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_MEMBER, "Invalid Member", "The selected member part is not a valid type for this buckling plate.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_SIDE, "Invalid Side", "Can not place buckling plate on this side of the member cross-section.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_WELDTOE1, "Invalid WeldToe1", "WeldToe1 should be within the member part width.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_WELDTOE2, "Invalid WeldToe2", "WeldToe2 should be within the member part width.")

    
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    
    Dim lErrorNumber As Long
    'validate member part
    If pGeometricConstruction.Inputs("MemberPart").Count <> 1 Then Err.Raise MISSING_MEMBER
    Dim oSPSMemberPartPrismatic As ISPSMemberPartPrismatic
    Dim oISPSDesignedMember As ISPSDesignedMember
    On Error Resume Next
    Set oSPSMemberPartPrismatic = pGeometricConstruction.Inputs("MemberPart").Item(1)
    Set oISPSDesignedMember = pGeometricConstruction.Inputs("MemberPart").Item(1)
    On Error GoTo 0
    If (oSPSMemberPartPrismatic Is Nothing) And (oISPSDesignedMember Is Nothing) Then Err.Raise INVALID_MEMBER
    
    'Check any parameter is -ve
    If (pGeometricConstruction.Parameter("Offset") < 0) Then Err.Raise INVALID_OFFSET
    If (pGeometricConstruction.Parameter("WeldToe1") < 0) Then Err.Raise INVALID_WELDTOE1
    If (pGeometricConstruction.Parameter("WeldToe2") < 0) Then Err.Raise INVALID_WELDTOE2
    If (pGeometricConstruction.Parameter("Height") < 0) Then Err.Raise INVALID_HEIGHT
    
    'Construct GC macro
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port").Clear
    pGeometricConstruction.ControlledInputs("Port").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")
    
    'Validate the Offset againest the member length
    Dim oLine As IJLine
    On Error Resume Next
    Set oLine = oLineAxisPortExtractor1.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Offset") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Offset") > oLine.Length Then Err.Raise INVALID_OFFSET
    End If
    
    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Height")
    oParamDistConstant2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe1")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe2")
    oParamDistConstant5.Evaluate
    
    Dim bWeldToe2MorethanWeldToe1 As Boolean
    Dim dWeldToe2 As Double
    Dim dWeldToe1 As Double
    bWeldToe2MorethanWeldToe1 = False
    dWeldToe1 = pGeometricConstruction.Parameter("WeldToe1")
    dWeldToe2 = pGeometricConstruction.Parameter("WeldToe2")
    If dWeldToe2 > dWeldToe1 Then
        bWeldToe2MorethanWeldToe1 = True
    End If

    Dim oPointAtCurveMiddle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle6 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0006-PointAtCurveMiddle")
    oPointAtCurveMiddle6.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
    oPointAtCurveMiddle6.Evaluate

    Dim oPointAtCurveStart7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart7 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0007-PointAtCurveStart")
    oPointAtCurveStart7.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
    oPointAtCurveStart7.Evaluate

    Dim oPointAtCurveEnd8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd8 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0008-PointAtCurveEnd")
    oPointAtCurveEnd8.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
    oPointAtCurveEnd8.Evaluate

    Dim oCSFromMember9 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember9 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0009-CSFromMember")
    oCSFromMember9.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oCSFromMember9.Inputs("Point").Add oPointAtCurveStart7, "1"
    oCSFromMember9.Evaluate

    Dim oFacePortExtractor10 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor10 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0010-FacePortExtractor")
    oFacePortExtractor10.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oFacePortExtractor10.Inputs("CoordinateSystem").Add oCSFromMember9, "1"
    oFacePortExtractor10.Parameter("GeometrySelector") = 4
    oFacePortExtractor10.Parameter("FacesContext") = 5
    oFacePortExtractor10.Parameter("LookingAxis") = 2
    oFacePortExtractor10.Parameter("IntersectingPlane") = 0
    oFacePortExtractor10.Parameter("SurfaceType") = 1
    oFacePortExtractor10.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFacePortExtractor10.Parameter("Offset") = 0#
    oFacePortExtractor10.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort_Web").Clear
    pGeometricConstruction.ControlledInputs("FacePort_Web").Add oFacePortExtractor10.ControlledInputs("Port").Item("1")
    
    Dim oPointAlongCurve11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve11 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0011-PointAlongCurve")
    oPointAlongCurve11.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
    oPointAlongCurve11.Inputs("Point").Add oPointAtCurveStart7, "1"
    oPointAlongCurve11.Inputs("TrackPoint").Add oPointAtCurveMiddle6, "1"
    oPointAlongCurve11.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve11.Parameter("TrackFlag") = 1
    oPointAlongCurve11.Evaluate


    Dim oPlaneByPointNormal12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal12 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0012-PlaneByPointNormal")
    oPlaneByPointNormal12.Inputs("Point").Add oPointAlongCurve11, "1"
    oPlaneByPointNormal12.Inputs("Line").Add oLineAxisPortExtractor1, "1"
    oPlaneByPointNormal12.Parameter("Range") = 2#
    oPlaneByPointNormal12.Evaluate

    Dim oCSByPlane13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane13 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0013-CSByPlane")
    oCSByPlane13.Inputs("Plane").Add oFacePortExtractor10, "1"
    oCSByPlane13.Inputs("Point").Add oPointAtCurveStart7, "1"
    oCSByPlane13.Evaluate

    Dim oLineFromCrossSection14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection14 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0014-LineFromCrossSection")
    oLineFromCrossSection14.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineFromCrossSection14.Inputs("Point").Add oPointAlongCurve11, "2"
    oLineFromCrossSection14.Parameter("EdgeSelector") = 514
    oLineFromCrossSection14.Evaluate
    
    'Validate the weldToe1
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection14.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe1") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe1") > oLine.Length Then Err.Raise INVALID_WELDTOE1
    End If


    Dim oLineFromCrossSection15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection15 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0015-LineFromCrossSection")
    oLineFromCrossSection15.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineFromCrossSection15.Inputs("Point").Add oPointAlongCurve11, "2"
    oLineFromCrossSection15.Parameter("EdgeSelector") = 513
    oLineFromCrossSection15.Evaluate
    
    'Validate the weldToe2
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection15.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe2") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe2") > oLine.Length Then Err.Raise INVALID_WELDTOE2
    End If


    Dim oLineFromCrossSection16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection16 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0016-LineFromCrossSection")
    oLineFromCrossSection16.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineFromCrossSection16.Inputs("Point").Add oPointAlongCurve11, "2"
    oLineFromCrossSection16.Parameter("EdgeSelector") = 258
    oLineFromCrossSection16.Evaluate
    
    'Validate the height
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection16.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Height") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Height") > oLine.Length Then Err.Raise INVALID_HEIGHT
    End If


    Dim oFacePortExtractor117 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor117 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "0017-FacePortExtractor1")
    oFacePortExtractor117.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oFacePortExtractor117.Inputs("Face").Add oFacePortExtractor10, "1"
    oFacePortExtractor117.Inputs("CoordinateSystem").Add oCSFromMember9, "1"
    oFacePortExtractor117.Parameter("GeometrySelector") = 4
    oFacePortExtractor117.Parameter("FacesContext") = 4
    oFacePortExtractor117.Parameter("LookingAxis") = 3
    oFacePortExtractor117.Parameter("SurfaceType") = 1
    oFacePortExtractor117.Parameter("TrackFlag") = 2
    oFacePortExtractor117.Parameter("Offset") = 0#
    
    'Validate the side
    On Error Resume Next
    lErrorNumber = 0
    oFacePortExtractor117.Evaluate
    lErrorNumber = Err.Number
    On Error GoTo 0
    If lErrorNumber <> 0 Then Err.Raise INVALID_SIDE
    
    pGeometricConstruction.ControlledInputs("FacePort_TopFlange").Clear
    pGeometricConstruction.ControlledInputs("FacePort_TopFlange").Add oFacePortExtractor117.ControlledInputs("Port").Item("1")
    
    Dim oFacePortExtractor118 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor118 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "0018-FacePortExtractor1")
    oFacePortExtractor118.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oFacePortExtractor118.Inputs("Face").Add oFacePortExtractor10, "1"
    oFacePortExtractor118.Inputs("CoordinateSystem").Add oCSFromMember9, "1"
    oFacePortExtractor118.Parameter("GeometrySelector") = 4
    oFacePortExtractor118.Parameter("FacesContext") = 4
    oFacePortExtractor118.Parameter("LookingAxis") = 3
    oFacePortExtractor118.Parameter("SurfaceType") = 1
    oFacePortExtractor118.Parameter("TrackFlag") = 1
    oFacePortExtractor118.Parameter("Offset") = 0#
    
    'Validate the side
    On Error Resume Next
    lErrorNumber = 0
    oFacePortExtractor118.Evaluate
    lErrorNumber = Err.Number
    On Error GoTo 0
    If lErrorNumber <> 0 Then Err.Raise INVALID_SIDE
    
    pGeometricConstruction.ControlledInputs("FacePort_BottomFlange").Clear
    pGeometricConstruction.ControlledInputs("FacePort_BottomFlange").Add oFacePortExtractor118.ControlledInputs("Port").Item("1")
    
    Dim oSurfFromGType19 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType19 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0019-SurfFromGType")
    oSurfFromGType19.Inputs("Surface").Add oPlaneByPointNormal12, "1"
    oSurfFromGType19.Evaluate

    Dim oPointAtCurveStart20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart20 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0020-PointAtCurveStart")
    oPointAtCurveStart20.Inputs("Curve").Add oLineFromCrossSection14, "1"
    oPointAtCurveStart20.Evaluate

    Dim oPointAtCurveEnd21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd21 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0021-PointAtCurveEnd")
    oPointAtCurveEnd21.Inputs("Curve").Add oLineFromCrossSection14, "1"
    oPointAtCurveEnd21.Evaluate

    Dim oPointAtCurveMiddle22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle22 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0022-PointAtCurveMiddle")
    oPointAtCurveMiddle22.Inputs("Curve").Add oLineFromCrossSection14, "1"
    oPointAtCurveMiddle22.Evaluate

    Dim oCurveAtOffset23 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveAtOffset23 = oGCFactory.CreateEntity("CurveAtOffset", pPOM, "0023-CurveAtOffset")
    oCurveAtOffset23.Inputs("Curve").Add oLineFromCrossSection14, "1"
    oCurveAtOffset23.Inputs("Surface").Add oSurfFromGType19, "1"
    oCurveAtOffset23.Parameter("Offset") = 0.015
    oCurveAtOffset23.Parameter("Orientation") = 1
    oCurveAtOffset23.Parameter("TrackFlag") = 1
    oCurveAtOffset23.Evaluate

    Dim oCurveAtOffset24 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveAtOffset24 = oGCFactory.CreateEntity("CurveAtOffset", pPOM, "0024-CurveAtOffset")
    oCurveAtOffset24.Inputs("Curve").Add oLineFromCrossSection15, "1"
    oCurveAtOffset24.Inputs("Surface").Add oSurfFromGType19, "1"
    oCurveAtOffset24.Parameter("Offset") = 0.015
    oCurveAtOffset24.Parameter("Orientation") = 1
    oCurveAtOffset24.Parameter("TrackFlag") = 1
    oCurveAtOffset24.Evaluate

    Dim oCurveByIntersection25 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection25 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0025-CurveByIntersection")
    oCurveByIntersection25.Inputs("Surface1").Add oSurfFromGType19, "1"
    oCurveByIntersection25.Inputs("Surface2").Add oFacePortExtractor117, "1"
    oCurveByIntersection25.Inputs("CoordinateSystem").Add oCSByPlane13, "2"
    oCurveByIntersection25.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oCurveByIntersection25.Parameter("ConstructionSurface") = 0
    oCurveByIntersection25.Evaluate

    Dim oCurveByIntersection26 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection26 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0026-CurveByIntersection")
    oCurveByIntersection26.Inputs("Surface1").Add oSurfFromGType19, "1"
    oCurveByIntersection26.Inputs("Surface2").Add oFacePortExtractor118, "2"
    oCurveByIntersection26.Inputs("CoordinateSystem").Add oCSByPlane13, "2"
    oCurveByIntersection26.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oCurveByIntersection26.Parameter("ConstructionSurface") = 0
    oCurveByIntersection26.Evaluate

    Dim oPointAtCurveEnd27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd27 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0027-PointAtCurveEnd")
    oPointAtCurveEnd27.Inputs("Curve").Add oCurveByIntersection25, "1"
    oPointAtCurveEnd27.Evaluate

    Dim oPointAtCurveMiddle28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle28 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0028-PointAtCurveMiddle")
    oPointAtCurveMiddle28.Inputs("Curve").Add oCurveByIntersection25, "1"
    oPointAtCurveMiddle28.Evaluate

    Dim oPointAtCurveEnd29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd29 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0029-PointAtCurveEnd")
    oPointAtCurveEnd29.Inputs("Curve").Add oCurveByIntersection26, "2"
    oPointAtCurveEnd29.Evaluate

    Dim oPointAtCurveMiddle30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle30 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0030-PointAtCurveMiddle")
    oPointAtCurveMiddle30.Inputs("Curve").Add oCurveByIntersection26, "2"
    oPointAtCurveMiddle30.Evaluate

    Dim oPointAlongCurve31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve31 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0031-PointAlongCurve")
    oPointAlongCurve31.Inputs("Curve").Add oCurveByIntersection25, "1"
    oPointAlongCurve31.Inputs("Point").Add oPointAtCurveEnd27, "2"
    oPointAlongCurve31.Inputs("TrackPoint").Add oPointAtCurveMiddle28, "1"
    oPointAlongCurve31.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve31.Parameter("TrackFlag") = 1
    oPointAlongCurve31.Evaluate

    Dim oPointAlongCurve32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve32 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0032-PointAlongCurve")
    oPointAlongCurve32.Inputs("Curve").Add oCurveByIntersection26, "1"
    oPointAlongCurve32.Inputs("Point").Add oPointAtCurveEnd29, "2"
    oPointAlongCurve32.Inputs("TrackPoint").Add oPointAtCurveMiddle30, "1"
    oPointAlongCurve32.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve32.Parameter("TrackFlag") = 1
    oPointAlongCurve32.Evaluate
    
    Dim oSurfByLinearExtrusion47 As SP3DGeometricConstruction.GeometricConstruction
    If bWeldToe2MorethanWeldToe1 = False Then
        Dim oVectorNormalToSurface34 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorNormalToSurface34 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "0034-VectorNormalToSurface")
        oVectorNormalToSurface34.Inputs("Surface").Add oFacePortExtractor118, "1"
        oVectorNormalToSurface34.Inputs("Point").Add oPointAlongCurve32, "1"
        oVectorNormalToSurface34.Parameter("Range") = 1#
        oVectorNormalToSurface34.Parameter("Orientation") = 1
        oVectorNormalToSurface34.Parameter("TrackFlag") = 1
        oVectorNormalToSurface34.Evaluate

        Dim oPointAlongCurve36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve36 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0036-PointAlongCurve")
        oPointAlongCurve36.Inputs("Curve").Add oVectorNormalToSurface34, "1"
        oPointAlongCurve36.Inputs("Point").Add oPointAlongCurve32, "1"
        oPointAlongCurve36.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve36.Parameter("TrackFlag") = 1
        oPointAlongCurve36.Evaluate
    
        Dim oLineByPoints37 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints37 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0037-LineByPoints")
        oLineByPoints37.Inputs("StartPoint").Add oPointAlongCurve36, "1"
        oLineByPoints37.Inputs("EndPoint").Add oPointAlongCurve31, "1"
        oLineByPoints37.Evaluate
    
        Dim oLineByPoints40 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints40 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0040-LineByPoints")
        oLineByPoints40.Inputs("StartPoint").Add oPointAlongCurve32, "1"
        oLineByPoints40.Inputs("EndPoint").Add oPointAlongCurve36, "1"
        oLineByPoints40.Evaluate
        
        
        Dim oLineExtend41 As SP3DGeometricConstruction.GeometricConstruction
        If Abs(dWeldToe1 - dWeldToe2) <= TOLERANCE Then
            'Handle extend
            Set oLineExtend41 = oGCFactory.CreateEntity("LineExtend", pPOM, "0041-LineExtend")
            oLineExtend41.Inputs("Line").Add oLineByPoints37, "1"
            oLineExtend41.Inputs("Curve1").Add oCurveAtOffset23, "1"
            oLineExtend41.Inputs("Curve2").Add oCurveAtOffset24, "1"
            oLineExtend41.Parameter("Context1") = 1
            oLineExtend41.Parameter("Context2") = 3
            oLineExtend41.Evaluate
        
            Set oSurfByLinearExtrusion47 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0047-SurfByLinearExtrusion")
            oSurfByLinearExtrusion47.Inputs("PlanarCrossSection").Add oLineExtend41, "1"
            oSurfByLinearExtrusion47.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1, "1"
            oSurfByLinearExtrusion47.Evaluate
        Else
            Set oLineExtend41 = oGCFactory.CreateEntity("LineExtend", pPOM, "0041-LineExtend")
            oLineExtend41.Inputs("Line").Add oLineByPoints37, "1"
            oLineExtend41.Inputs("Curve1").Add oCurveAtOffset23, "1"
            oLineExtend41.Inputs("Curve2").Add oLineByPoints40, "1"
            oLineExtend41.Parameter("Context1") = 1
            oLineExtend41.Parameter("Context2") = 3
            oLineExtend41.Evaluate
        
            Dim oLineExtend42 As SP3DGeometricConstruction.GeometricConstruction
            Set oLineExtend42 = oGCFactory.CreateEntity("LineExtend", pPOM, "0042-LineExtend")
            oLineExtend42.Inputs("Line").Add oLineByPoints40, "1"
            oLineExtend42.Inputs("Curve1").Add oCurveAtOffset24, "1"
            oLineExtend42.Parameter("Context1") = 1
            oLineExtend42.Parameter("Context2") = 3
            oLineExtend42.Evaluate
        
            Dim oCpxStringByCurves45 As SP3DGeometricConstruction.GeometricConstruction
            Set oCpxStringByCurves45 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0045-CpxStringByCurves")
            oCpxStringByCurves45.Inputs("Curves").Add oLineExtend42, "1"
            oCpxStringByCurves45.Inputs("Curves").Add oLineExtend41, "2"
            oCpxStringByCurves45.Evaluate
        
            Set oSurfByLinearExtrusion47 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0047-SurfByLinearExtrusion")
            oSurfByLinearExtrusion47.Inputs("PlanarCrossSection").Add oCpxStringByCurves45, "1"
            oSurfByLinearExtrusion47.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1, "1"
            oSurfByLinearExtrusion47.Evaluate
        End If
    Else

        Dim oVectorNormalToSurface33 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorNormalToSurface33 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "0033-VectorNormalToSurface")
        oVectorNormalToSurface33.Inputs("Surface").Add oFacePortExtractor117, "1"
        oVectorNormalToSurface33.Inputs("Point").Add oPointAlongCurve31, "1"
        oVectorNormalToSurface33.Parameter("Range") = 1#
        oVectorNormalToSurface33.Parameter("Orientation") = 1
        oVectorNormalToSurface33.Parameter("TrackFlag") = 1
        oVectorNormalToSurface33.Evaluate

        Dim oPointAlongCurve35 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve35 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0035-PointAlongCurve")
        oPointAlongCurve35.Inputs("Curve").Add oVectorNormalToSurface33, "1"
        oPointAlongCurve35.Inputs("Point").Add oPointAlongCurve31, "1"
        oPointAlongCurve35.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve35.Parameter("TrackFlag") = 1
        oPointAlongCurve35.Evaluate
    
        Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0038-LineByPoints")
        oLineByPoints38.Inputs("StartPoint").Add oPointAlongCurve31, "1"
        oLineByPoints38.Inputs("EndPoint").Add oPointAlongCurve35, "1"
        oLineByPoints38.Evaluate
    
        Dim oLineByPoints39 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints39 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0039-LineByPoints")
        oLineByPoints39.Inputs("StartPoint").Add oPointAlongCurve35, "1"
        oLineByPoints39.Inputs("EndPoint").Add oPointAlongCurve32, "1"
        oLineByPoints39.Evaluate
    
        Dim oLineExtend43 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend43 = oGCFactory.CreateEntity("LineExtend", pPOM, "0043-LineExtend")
        oLineExtend43.Inputs("Line").Add oLineByPoints38, "1"
        oLineExtend43.Inputs("Curve1").Add oCurveAtOffset23, "1"
        oLineExtend43.Parameter("Context1") = 1
        oLineExtend43.Parameter("Context2") = 3
        oLineExtend43.Evaluate
    
        Dim oLineExtend44 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend44 = oGCFactory.CreateEntity("LineExtend", pPOM, "0044-LineExtend")
        oLineExtend44.Inputs("Line").Add oLineByPoints39, "1"
        oLineExtend44.Inputs("Curve1").Add oCurveAtOffset24, "1"
        oLineExtend44.Inputs("Curve2").Add oLineByPoints38, "1"
        oLineExtend44.Parameter("Context1") = 1
        oLineExtend44.Parameter("Context2") = 3
        oLineExtend44.Evaluate
    
        Dim oCpxStringByCurves46 As SP3DGeometricConstruction.GeometricConstruction
        Set oCpxStringByCurves46 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0046-CpxStringByCurves")
        oCpxStringByCurves46.Inputs("Curves").Add oLineExtend43, "1"
        oCpxStringByCurves46.Inputs("Curves").Add oLineExtend44, "2"
        oCpxStringByCurves46.Evaluate
    
        Set oSurfByLinearExtrusion47 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0048-SurfByLinearExtrusion")
        oSurfByLinearExtrusion47.Inputs("PlanarCrossSection").Add oCpxStringByCurves46, "1"
        oSurfByLinearExtrusion47.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1, "1"
        oSurfByLinearExtrusion47.Evaluate
    End If

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    
    oGCMacro.Output("Support", 1) = oSurfFromGType19.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion47.Output
    
    'Add faceport boundary output with name same as the controlled input i.e. FacePort_Web
    oGCMacro.Output("Boundary", "FacePort_Web") = oFacePortExtractor10.Output
    
    If TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSDesignedMember Then
        'Add faceport boundary output with name same as the controlled input i.e. FacePort_TopFlange
        oGCMacro.Output("Boundary", "FacePort_TopFlange") = oFacePortExtractor117.Output
        'Add faceport boundary output with name same as the controlled input i.e. FacePort_BottomFlange
        oGCMacro.Output("Boundary", "FacePort_BottomFlange") = oFacePortExtractor118.Output
    End If
    
End Sub
