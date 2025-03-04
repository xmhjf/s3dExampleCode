VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklPlateOneSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   File:           BucklPlateOneSide.cls
'   ProgID:         GCBucklPlateOneSide.BucklPlateOneSide
'   Author:         Alligators Team(India)
'   Creation Date:  Tuesday, Aug 10 2010
'   Description:
'      This macro places an advanced plate on either a standard
'      member or on a built up member.

'      Two sub-routines are used in 'Evaluate' method (one for standard member and one
'      for built-up).
'
'   Change History:
'   dd.mmm.yyyy     who              change description
'   -----------     ---              ------------------
'   10.Aug.2010      CM              CR-CP-181617 Creation.
'   2.Dec.2010     svsmylav          TK-79111: Added regular APS template; Also modified to make
'                                    the GC placeable on standard member.
'   18.July.2011     CM              TR-CP-200358 Changed the Output Support type from Plane to
'                                      Gtype Surface Body.(This got impacted due to change in
'                                       middle tier)
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCToDoDelegate

Private m_oRegularAPSHelper As Object
Private m_oAdvancedDebug As New AdvancedDebug

Private Property Get Source() As String
    Let Source = "GCBucklingPlatePartial.Partial" 'GCCustomPlate.CustomPlate"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property

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
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oRegularAPSHelper.PreInitialize(pGeometricConstructionDefinition)
    Call Initialize(pGeometricConstructionDefinition)
    Call m_oRegularAPSHelper.PostInitialize(pGeometricConstructionDefinition)
End Sub

Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call m_oRegularAPSHelper.PreEvaluate(pGeometricConstruction, pPOM)
    ' customized evaluation
    Call Evaluate(pGeometricConstruction, pPOM)
    Call m_oRegularAPSHelper.PostEvaluate(pGeometricConstruction, pPOM)
    
    Exit Sub
ErrorHandler:
    Call m_oRegularAPSHelper.ProcessError(pGeometricConstruction, Err.Number)
End Sub

' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)

    Call m_oRegularAPSHelper.PostConnectionAdded(oRelationship)
    
End Sub

Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    
    Call m_oRegularAPSHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

End Sub '
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
' to be overriden by print graph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a RolledMember or a DesignedMember", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("Port", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort11")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort12")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort13")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_11")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_21")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_31")
    
    'For standard member
    Call pGeometricConstructionDefinition.AddControlledInput("Port_41", "IJGeometry")
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe1", "WeldToe1", 8, 1, 59, 0, 0, 0.04)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe2", "WeldToe2", 8, 1, 59, 0, 0, 0.015)
    Call pGeometricConstructionDefinition.AddParameter("Height", "Height", 8, 1, 59, 0, 0, 0.6)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 0.02)
    Call pGeometricConstructionDefinition.AddParameter("TrackFlag", "Near/Far", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Near", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Far", 2)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)

    If TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSMemberPartPrismatic Then
        'StandardMember, code is received from Lee
        CreateAPSForStandardMember pGeometricConstruction, pPOM
    ElseIf TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSDesignedMember Then
        'BuiltUp
        CreateAPSForBuiltUpMember pGeometricConstruction, pPOM
    End If

End Sub
Private Sub CreateAPSForStandardMember(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
   Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oFacePortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor1 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "001-FacePortExtractor")
    oFacePortExtractor1.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor1.Parameter("GeometrySelector") = 2
    oFacePortExtractor1.Parameter("FacesContext") = 2
    oFacePortExtractor1.Parameter("LookingAxis") = 2
    oFacePortExtractor1.Parameter("SurfaceType") = 1
    oFacePortExtractor1.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFacePortExtractor1.Parameter("Offset") = 0#
    oFacePortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port_11").Clear
    pGeometricConstruction.ControlledInputs("Port_11").AddElements oFacePortExtractor1.ControlledInputs("Port")

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("Port_41").Clear
    pGeometricConstruction.ControlledInputs("Port_41").AddElements oLineAxisPortExtractor2.ControlledInputs("Port")

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Height")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("Radius")
    oParamDistConstant5.Evaluate
    
    Dim bRisZero As Boolean
    bRisZero = False
    If oParamDistConstant5.Parameter("Value") = 0 Then bRisZero = True

    Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "006-ParamDistConstant")
    oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe1")
    oParamDistConstant6.Evaluate

    Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "007-ParamDistConstant")
    oParamDistConstant7.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe2")
    oParamDistConstant7.Evaluate

    Dim oPointAtCurveStart8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart8 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "008-PointAtCurveStart")
    oPointAtCurveStart8.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveStart8.Evaluate

    Dim oPointAtCurveMiddle9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle9 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "009-PointAtCurveMiddle")
    oPointAtCurveMiddle9.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveMiddle9.Evaluate

    Dim oPointAlongCurve10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve10 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "010-PointAlongCurve")
    oPointAlongCurve10.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAlongCurve10.Inputs("Point").Add oPointAtCurveStart8
    oPointAlongCurve10.Inputs("TrackPoint").Add oPointAtCurveMiddle9
    oPointAlongCurve10.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve10.Parameter("TrackFlag") = 1
    oPointAlongCurve10.Evaluate

    Dim oPlaneByPointNormal11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal11 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "011-PlaneByPointNormal")
    oPlaneByPointNormal11.Inputs("Point").Add oPointAlongCurve10
    oPlaneByPointNormal11.Inputs("Line").Add oLineAxisPortExtractor2
    oPlaneByPointNormal11.Parameter("Range") = 5#
    oPlaneByPointNormal11.Evaluate

    Dim oCSByPlane12 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane12 = oGCFactory.CreateEntity("CSByPlane", pPOM, "012-CSByPlane")
    oCSByPlane12.Inputs("Plane").Add oFacePortExtractor1
    oCSByPlane12.Inputs("Point").Add oPointAtCurveStart8
    oCSByPlane12.Evaluate

    Dim oCSFromMember13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember13 = oGCFactory.CreateEntity("CSFromMember", pPOM, "013-CSFromMember")
    oCSFromMember13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember13.Inputs("Point").Add oPointAtCurveStart8
    oCSFromMember13.Evaluate

    Dim oCurveByIntersection14 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection14 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "014-CurveByIntersection")
    oCurveByIntersection14.Inputs("Surface1").Add oPlaneByPointNormal11
    oCurveByIntersection14.Inputs("Surface2").Add oFacePortExtractor1
    oCurveByIntersection14.Inputs("CoordinateSystem").Add oCSByPlane12
    oCurveByIntersection14.Parameter("TrackFlag") = 1
    oCurveByIntersection14.Evaluate

    Dim oFacePortExtractor15 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor15 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "015-FacePortExtractor1")
    oFacePortExtractor15.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor15.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor15.Inputs("Face").Add oFacePortExtractor1
    oFacePortExtractor15.Parameter("GeometrySelector") = 2
    oFacePortExtractor15.Parameter("FacesContext") = 4
    oFacePortExtractor15.Parameter("LookingAxis") = 3
    oFacePortExtractor15.Parameter("SurfaceType") = 1
    oFacePortExtractor15.Parameter("TrackFlag") = 1
    oFacePortExtractor15.Parameter("Offset") = 0#
    oFacePortExtractor15.Evaluate
    pGeometricConstruction.ControlledInputs("Port_21").Clear
    pGeometricConstruction.ControlledInputs("Port_21").AddElements oFacePortExtractor15.ControlledInputs("Port")

    Dim oFacePortExtractor16 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "016-FacePortExtractor1")
    oFacePortExtractor16.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor16.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor16.Inputs("Face").Add oFacePortExtractor1
    oFacePortExtractor16.Parameter("GeometrySelector") = 2
    oFacePortExtractor16.Parameter("FacesContext") = 4
    oFacePortExtractor16.Parameter("LookingAxis") = 3
    oFacePortExtractor16.Parameter("SurfaceType") = 1
    oFacePortExtractor16.Parameter("TrackFlag") = 2
    oFacePortExtractor16.Parameter("Offset") = 0#
    oFacePortExtractor16.Evaluate
    pGeometricConstruction.ControlledInputs("Port_31").Clear
    pGeometricConstruction.ControlledInputs("Port_31").AddElements oFacePortExtractor16.ControlledInputs("Port")

    Dim oCurveByIntersection17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection17 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "017-CurveByIntersection")
    oCurveByIntersection17.Inputs("Surface1").Add oPlaneByPointNormal11
    oCurveByIntersection17.Inputs("Surface2").Add oFacePortExtractor15
    oCurveByIntersection17.Inputs("CoordinateSystem").Add oCSByPlane12
    oCurveByIntersection17.Parameter("TrackFlag") = 1
    oCurveByIntersection17.Evaluate

    Dim oCurveByIntersection18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection18 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "018-CurveByIntersection")
    oCurveByIntersection18.Inputs("Surface1").Add oPlaneByPointNormal11
    oCurveByIntersection18.Inputs("Surface2").Add oFacePortExtractor16
    oCurveByIntersection18.Inputs("CoordinateSystem").Add oCSByPlane12
    oCurveByIntersection18.Parameter("TrackFlag") = 1
    oCurveByIntersection18.Evaluate
    
    'For Tracking Point
    Dim oPointAtCurveMiddle19_1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle19_1 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "019_1-PointAtCurveMiddle")
    oPointAtCurveMiddle19_1.Inputs("Curve").Add oCurveByIntersection18
    oPointAtCurveMiddle19_1.Evaluate

    If oParamDistConstant3.Parameter("Value") > 0 Then 'H1 is more than zero
        Dim oLineParallel19 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel19 = oGCFactory.CreateEntity("LineParallel", pPOM, "019-LineParallel")
        oLineParallel19.Inputs("Line").Add oCurveByIntersection17
        oLineParallel19.Inputs("Plane").Add oPlaneByPointNormal11
        oLineParallel19.Inputs("TrackPoint").Add oPointAtCurveMiddle19_1
        oLineParallel19.Parameter("Distance") = oParamDistConstant3
        oLineParallel19.Parameter("TrackFlag") = 1
        oLineParallel19.Evaluate
   End If

    Dim oPointByCurves20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves20 = oGCFactory.CreateEntity("PointByCurves", pPOM, "020-PointByCurves")
    oPointByCurves20.Inputs("Curve1").Add oCurveByIntersection14
    oPointByCurves20.Inputs("Curve2").Add oCurveByIntersection17
    oPointByCurves20.Parameter("TrackFlag") = 1
    oPointByCurves20.Evaluate

    Dim oPointAtCurveEnd21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd21 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "021-PointAtCurveEnd")
    oPointAtCurveEnd21.Inputs("Curve").Add oCurveByIntersection17
    oPointAtCurveEnd21.Evaluate

    Dim oPointByCurves22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves22 = oGCFactory.CreateEntity("PointByCurves", pPOM, "022-PointByCurves")
    oPointByCurves22.Inputs("Curve1").Add oCurveByIntersection14
    oPointByCurves22.Inputs("Curve2").Add oCurveByIntersection18
    oPointByCurves22.Parameter("TrackFlag") = 1
    oPointByCurves22.Evaluate

    Dim oPointAtCurveEnd23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd23 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "023-PointAtCurveEnd")
    oPointAtCurveEnd23.Inputs("Curve").Add oCurveByIntersection18
    oPointAtCurveEnd23.Evaluate

    Dim oLineByPoints24 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints24 = oGCFactory.CreateEntity("LineByPoints", pPOM, "024-LineByPoints")
    oLineByPoints24.Inputs("StartPoint").Add oPointByCurves20
    oLineByPoints24.Inputs("EndPoint").Add oPointAtCurveEnd21
    oLineByPoints24.Evaluate

    Dim oLineByPoints25 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints25 = oGCFactory.CreateEntity("LineByPoints", pPOM, "025-LineByPoints")
    oLineByPoints25.Inputs("StartPoint").Add oPointByCurves22
    oLineByPoints25.Inputs("EndPoint").Add oPointByCurves20
    oLineByPoints25.Evaluate

    Dim oLineByPoints26 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints26 = oGCFactory.CreateEntity("LineByPoints", pPOM, "026-LineByPoints")
    oLineByPoints26.Inputs("StartPoint").Add oPointByCurves20
    oLineByPoints26.Inputs("EndPoint").Add oPointByCurves22
    oLineByPoints26.Evaluate

    Dim oLineByPoints27 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints27 = oGCFactory.CreateEntity("LineByPoints", pPOM, "027-LineByPoints")
    oLineByPoints27.Inputs("StartPoint").Add oPointByCurves22
    oLineByPoints27.Inputs("EndPoint").Add oPointAtCurveEnd23
    oLineByPoints27.Evaluate

    Dim oPointAlongCurve31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve31 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "031-PointAlongCurve")
    oPointAlongCurve31.Inputs("Curve").Add oCurveByIntersection17
    oPointAlongCurve31.Inputs("Point").Add oPointAtCurveEnd21
    oPointAlongCurve31.Inputs("TrackPoint").Add oPointAtCurveMiddle19_1
    oPointAlongCurve31.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve31.Parameter("TrackFlag") = 1
    oPointAlongCurve31.Evaluate

    Dim oPointAlongCurve32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve32 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "032-PointAlongCurve")
    oPointAlongCurve32.Inputs("Curve").Add oCurveByIntersection18
    oPointAlongCurve32.Inputs("Point").Add oPointAtCurveEnd23
    oPointAlongCurve32.Inputs("TrackPoint").Add oPointAtCurveMiddle19_1
    oPointAlongCurve32.Parameter("Distance") = oParamDistConstant6
    oPointAlongCurve32.Parameter("TrackFlag") = 1
    oPointAlongCurve32.Evaluate

    If oParamDistConstant3.Parameter("Value") > 0 Then
    
        Dim oPointAtCurveEnd28 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveEnd28 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "028-PointAtCurveEnd")
        oPointAtCurveEnd28.Inputs("Curve").Add oLineParallel19
        oPointAtCurveEnd28.Evaluate
        
        Dim oPointAlongCurve33 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve33 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "033-PointAlongCurve")
        oPointAlongCurve33.Inputs("Curve").Add oLineParallel19
        oPointAlongCurve33.Inputs("Point").Add oPointAtCurveEnd28
        oPointAlongCurve33.Inputs("TrackPoint").Add oPointAtCurveMiddle19_1
        oPointAlongCurve33.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve33.Parameter("TrackFlag") = 1
        oPointAlongCurve33.Evaluate
    End If
    
    If bRisZero = False Then

        Dim oArcByFillet29 As SP3DGeometricConstruction.GeometricConstruction
        Set oArcByFillet29 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "029-ArcByFillet")
        oArcByFillet29.Inputs("Curve1").Add oLineByPoints24
        oArcByFillet29.Inputs("Curve2").Add oLineByPoints26
        oArcByFillet29.Parameter("Radius") = oParamDistConstant5
        oArcByFillet29.Evaluate
    
        Dim oArcByFillet30 As SP3DGeometricConstruction.GeometricConstruction
        Set oArcByFillet30 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "030-ArcByFillet")
        oArcByFillet30.Inputs("Curve1").Add oLineByPoints25
        oArcByFillet30.Inputs("Curve2").Add oLineByPoints27
        oArcByFillet30.Parameter("Radius") = oParamDistConstant5
        oArcByFillet30.Evaluate
        
        Dim oPointByCurves34 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves34 = oGCFactory.CreateEntity("PointByCurves", pPOM, "034-PointByCurves")
        oPointByCurves34.Inputs("Curve1").Add oArcByFillet29
        oPointByCurves34.Inputs("Curve2").Add oLineByPoints26
        oPointByCurves34.Parameter("TrackFlag") = 1
        oPointByCurves34.Evaluate
        
        Dim oPointByCurves34_1 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves34_1 = oGCFactory.CreateEntity("PointByCurves", pPOM, "034_1-PointByCurves")
        oPointByCurves34_1.Inputs("Curve1").Add oArcByFillet29
        oPointByCurves34_1.Inputs("Curve2").Add oLineByPoints24
        oPointByCurves34_1.Parameter("TrackFlag") = 1
        oPointByCurves34_1.Evaluate
        
        Dim oPointByCurves36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves36 = oGCFactory.CreateEntity("PointByCurves", pPOM, "036-PointByCurves")
        oPointByCurves36.Inputs("Curve1").Add oLineByPoints25
        oPointByCurves36.Inputs("Curve2").Add oArcByFillet30
        oPointByCurves36.Parameter("TrackFlag") = 1
        oPointByCurves36.Evaluate
    
        Dim oPointByCurves37 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves37 = oGCFactory.CreateEntity("PointByCurves", pPOM, "037-PointByCurves")
        oPointByCurves37.Inputs("Curve1").Add oArcByFillet30
        oPointByCurves37.Inputs("Curve2").Add oCurveByIntersection18
        oPointByCurves37.Parameter("TrackFlag") = 1
        oPointByCurves37.Evaluate
        
        'Boundary List
        Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "038-LineByPoints")
        oLineByPoints38.Inputs("StartPoint").Add oPointByCurves37
        oLineByPoints38.Inputs("EndPoint").Add oPointAlongCurve32
        oLineByPoints38.Evaluate
        
        Dim oLineByPoints39 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints39 = oGCFactory.CreateEntity("LineByPoints", pPOM, "039-LineByPoints")
        oLineByPoints39.Inputs("StartPoint").Add oPointByCurves36
        oLineByPoints39.Inputs("EndPoint").Add oPointByCurves34
        oLineByPoints39.Evaluate
        
        Dim oLineByPoints40 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints40 = oGCFactory.CreateEntity("LineByPoints", pPOM, "040-LineByPoints")
        oLineByPoints40.Inputs("StartPoint").Add oPointByCurves34_1
        oLineByPoints40.Inputs("EndPoint").Add oPointAlongCurve31
        oLineByPoints40.Evaluate
        
        Dim oSurfByLinearExtrusion41 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion41 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "041-SurfByLinearExtrusion")
        oSurfByLinearExtrusion41.Inputs("PlanarCrossSection").Add oArcByFillet29
        oSurfByLinearExtrusion41.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion41.Evaluate
        
        Dim oSurfByLinearExtrusion42 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion42 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "042-SurfByLinearExtrusion")
        oSurfByLinearExtrusion42.Inputs("PlanarCrossSection").Add oArcByFillet30
        oSurfByLinearExtrusion42.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion42.Evaluate
        
        Dim oSurfByLinearExtrusion43 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion43 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "043-SurfByLinearExtrusion")
        oSurfByLinearExtrusion43.Inputs("PlanarCrossSection").Add oLineByPoints38
        oSurfByLinearExtrusion43.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion43.Evaluate
        
        Dim oSurfByLinearExtrusion44 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion44 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "044-SurfByLinearExtrusion")
        oSurfByLinearExtrusion44.Inputs("PlanarCrossSection").Add oLineByPoints39
        oSurfByLinearExtrusion44.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion44.Evaluate
        
        Dim oSurfByLinearExtrusion45 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion45 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "045-SurfByLinearExtrusion")
        oSurfByLinearExtrusion45.Inputs("PlanarCrossSection").Add oLineByPoints40
        oSurfByLinearExtrusion45.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion45.Evaluate

    End If

    If oParamDistConstant3.Parameter("Value") > 0 Then 'H1 >0
        Dim oLineByPoints46 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints46 = oGCFactory.CreateEntity("LineByPoints", pPOM, "046-LineByPoints")
        oLineByPoints46.Inputs("StartPoint").Add oPointAlongCurve31
        oLineByPoints46.Inputs("EndPoint").Add oPointAlongCurve33
        oLineByPoints46.Evaluate
    
        Dim oLineByPoints47 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints47 = oGCFactory.CreateEntity("LineByPoints", pPOM, "047-LineByPoints")
        oLineByPoints47.Inputs("StartPoint").Add oPointAlongCurve33
        oLineByPoints47.Inputs("EndPoint").Add oPointAlongCurve32
        oLineByPoints47.Evaluate
        
        Dim oSurfByLinearExtrusion48 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion48 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "048-SurfByLinearExtrusion")
        oSurfByLinearExtrusion48.Inputs("PlanarCrossSection").Add oLineByPoints46
        oSurfByLinearExtrusion48.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion48.Evaluate

    ElseIf oParamDistConstant3.Parameter("Value") = 0 Then
    
        Dim oLineByPoints47_1 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints47_1 = oGCFactory.CreateEntity("LineByPoints", pPOM, "047_1-LineByPoints")
        oLineByPoints47_1.Inputs("StartPoint").Add oPointAlongCurve31
        oLineByPoints47_1.Inputs("EndPoint").Add oPointAlongCurve32
        oLineByPoints47_1.Evaluate
 
    End If
    
        Dim oSurfByLinearExtrusion49 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion49 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "049-SurfByLinearExtrusion")
        If oParamDistConstant3.Parameter("Value") > 0 Then 'H1 >0
            oSurfByLinearExtrusion49.Inputs("PlanarCrossSection").Add oLineByPoints47
        ElseIf oParamDistConstant3.Parameter("Value") = 0 Then
            oSurfByLinearExtrusion49.Inputs("PlanarCrossSection").Add oLineByPoints47_1
        End If
        oSurfByLinearExtrusion49.Inputs("ExtrusionLine").Add oLineAxisPortExtractor2
        oSurfByLinearExtrusion49.Evaluate
    
    Dim oSurfFromGType50 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType50 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "050-SurfFromGType")
    oSurfFromGType50.Inputs("Surface").Add oPlaneByPointNormal11
    oSurfFromGType50.Evaluate
        
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    If bRisZero = False Then ' R>0
        oGCMacro.Output("Support", 1) = oSurfFromGType50.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion41.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion42.Output
        oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion43.Output
        oGCMacro.Output("Boundary", 4) = oSurfByLinearExtrusion44.Output
        oGCMacro.Output("Boundary", 5) = oSurfByLinearExtrusion45.Output
        oGCMacro.Output("Boundary", 6) = oSurfByLinearExtrusion49.Output
        If oParamDistConstant3.Parameter("Value") > 0 Then 'H1 >0
            oGCMacro.Output("Boundary", 7) = oSurfByLinearExtrusion48.Output
        End If
    Else 'R is zero
    
        If oGCMacro.Outputs("Boundary").Count = 7 Then
            oGCMacro.Outputs("Boundary").Remove (7)
            oGCMacro.Outputs("Boundary").Remove (6)
        ElseIf oGCMacro.Outputs("Boundary").Count = 6 Then
            oGCMacro.Outputs("Boundary").Remove (6)
        End If
        oGCMacro.Output("Support", 1) = oSurfFromGType50.Output
        oGCMacro.Output("Boundary", 1) = oFacePortExtractor1.Output
        oGCMacro.Output("Boundary", 2) = oFacePortExtractor15.Output
        oGCMacro.Output("Boundary", 3) = oFacePortExtractor16.Output
        oGCMacro.Output("Boundary", 4) = oSurfByLinearExtrusion49.Output
        If oParamDistConstant3.Parameter("Value") > 0 Then 'H1 >0
            oGCMacro.Output("Boundary", 5) = oSurfByLinearExtrusion48.Output
        End If
    End If
End Sub

Private Sub CreateAPSForBuiltUpMember(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    'Assumption: Radius value less than 0.1 mm is considered to be zero radius.
    Dim bCreateArcOnAPS As Boolean
    Dim dArcRadius As Double
    dArcRadius = pGeometricConstruction.Parameter("Radius")
    If dArcRadius < 0.0001 Then
        dArcRadius = 0.0001 'A small value is used so that existing code need not be changed to a large extent
        bCreateArcOnAPS = False 'Do not create arc
    Else
        bCreateArcOnAPS = True  'Create arc
    End If

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port").Clear
    pGeometricConstruction.ControlledInputs("Port").AddElements oLineAxisPortExtractor1.ControlledInputs("Port")

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Height")
    oParamDistConstant2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = dArcRadius 'pGeometricConstruction.Parameter("Radius")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe1")
    oParamDistConstant5.Evaluate

    Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "006-ParamDistConstant")
    oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe2")
    oParamDistConstant6.Evaluate

    Dim oPointAtCurveStart7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart7 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "007-PointAtCurveStart")
    oPointAtCurveStart7.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveStart7.Evaluate

    Dim oPointAtCurveEnd8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd8 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "008-PointAtCurveEnd")
    oPointAtCurveEnd8.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveEnd8.Evaluate

    Dim oLineFromCrossSection9 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection9 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "009-LineFromCrossSection")
    oLineFromCrossSection9.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection9.Inputs("Point").Add oPointAtCurveStart7
    oLineFromCrossSection9.Parameter("EdgeSelector") = 514
    oLineFromCrossSection9.Evaluate

    Dim oLineFromCrossSection10 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection10 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "010-LineFromCrossSection")
    oLineFromCrossSection10.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection10.Inputs("Point").Add oPointAtCurveStart7
    oLineFromCrossSection10.Parameter("EdgeSelector") = 513
    oLineFromCrossSection10.Evaluate

    Dim oLineFromCrossSection11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection11 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "011-LineFromCrossSection")
    oLineFromCrossSection11.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection11.Inputs("Point").Add oPointAtCurveStart7
    oLineFromCrossSection11.Parameter("EdgeSelector") = 258
    oLineFromCrossSection11.Evaluate

    Dim oCSFromMember12 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember12 = oGCFactory.CreateEntity("CSFromMember", pPOM, "012-CSFromMember")
    oCSFromMember12.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember12.Inputs("Point").Add oPointAtCurveStart7
    oCSFromMember12.Evaluate

    Dim oExtractPorts13 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts13 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "013-ExtractPorts")
    oExtractPorts13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oExtractPorts13.Inputs("ZAxis").Add oLineFromCrossSection9
    oExtractPorts13.Parameter("Support") = 1
    oExtractPorts13.Parameter("Debug") = 0
    oExtractPorts13.Evaluate
    pGeometricConstruction.ControlledInputs("Port").Clear
    pGeometricConstruction.ControlledInputs("Port").AddElements oExtractPorts13.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort13").Clear
    pGeometricConstruction.ControlledInputs("FacePort13").AddElements oExtractPorts13.ControlledInputs("FacePort")

    Dim oExtractPorts14 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts14 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "014-ExtractPorts")
    oExtractPorts14.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oExtractPorts14.Inputs("ZAxis").Add oLineFromCrossSection11
    oExtractPorts14.Parameter("Support") = 2
    oExtractPorts14.Parameter("Debug") = 0
    oExtractPorts14.Evaluate
    pGeometricConstruction.ControlledInputs("Port").Clear
    pGeometricConstruction.ControlledInputs("Port").AddElements oExtractPorts14.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort11").Clear
    pGeometricConstruction.ControlledInputs("FacePort11").AddElements oExtractPorts14.ControlledInputs("FacePort")

    Dim oExtractPorts15 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts15 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "015-ExtractPorts")
    oExtractPorts15.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oExtractPorts15.Inputs("ZAxis").Add oLineFromCrossSection11
    oExtractPorts15.Parameter("Support") = 1
    oExtractPorts15.Parameter("Debug") = 0
    oExtractPorts15.Evaluate
    pGeometricConstruction.ControlledInputs("Port").Clear
    pGeometricConstruction.ControlledInputs("Port").AddElements oExtractPorts15.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort12").Clear
    pGeometricConstruction.ControlledInputs("FacePort12").AddElements oExtractPorts15.ControlledInputs("FacePort")

    Dim oFacePortExtractor16 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "016-FacePortExtractor")
    oFacePortExtractor16.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor16.Inputs("CoordinateSystem").Add oCSFromMember12
    oFacePortExtractor16.Parameter("GeometrySelector") = 4
    oFacePortExtractor16.Parameter("FacesContext") = 4
    oFacePortExtractor16.Parameter("LookingAxis") = 3
    oFacePortExtractor16.Parameter("SurfaceType") = 1
    oFacePortExtractor16.Parameter("TrackFlag") = 2
    oFacePortExtractor16.Parameter("Offset") = 0#
    oFacePortExtractor16.Evaluate
    pGeometricConstruction.ControlledInputs("Port_11").Clear
    pGeometricConstruction.ControlledInputs("Port_11").AddElements oFacePortExtractor16.ControlledInputs("Port")

    Dim oFacePortExtractor17 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor17 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "017-FacePortExtractor")
    oFacePortExtractor17.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor17.Inputs("CoordinateSystem").Add oCSFromMember12
    oFacePortExtractor17.Parameter("GeometrySelector") = 4
    oFacePortExtractor17.Parameter("FacesContext") = 4
    oFacePortExtractor17.Parameter("LookingAxis") = 3
    oFacePortExtractor17.Parameter("SurfaceType") = 1
    oFacePortExtractor17.Parameter("TrackFlag") = 1
    oFacePortExtractor17.Parameter("Offset") = 0#
    oFacePortExtractor17.Evaluate
    pGeometricConstruction.ControlledInputs("Port_21").Clear
    pGeometricConstruction.ControlledInputs("Port_21").AddElements oFacePortExtractor17.ControlledInputs("Port")

    Dim oFaceFromCrossSection18 As SP3DGeometricConstruction.GeometricConstruction
    Set oFaceFromCrossSection18 = oGCFactory.CreateEntity("FaceFromCrossSection", pPOM, "018-FaceFromCrossSection")
    oFaceFromCrossSection18.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFaceFromCrossSection18.Inputs("CoordinateSystem").Add oCSFromMember12
    oFaceFromCrossSection18.Parameter("LookingAxis") = 2
    oFaceFromCrossSection18.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFaceFromCrossSection18.Parameter("Offset") = 0#
    oFaceFromCrossSection18.Evaluate

    Dim oFacePortExtractor19 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor19 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "019-FacePortExtractor")
    oFacePortExtractor19.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor19.Inputs("CoordinateSystem").Add oCSFromMember12
    oFacePortExtractor19.Parameter("GeometrySelector") = 4
    oFacePortExtractor19.Parameter("FacesContext") = 4
    oFacePortExtractor19.Parameter("LookingAxis") = 2
    oFacePortExtractor19.Parameter("SurfaceType") = 1
    oFacePortExtractor19.Parameter("TrackFlag") = 2
    oFacePortExtractor19.Parameter("Offset") = 0#
    oFacePortExtractor19.Evaluate
    pGeometricConstruction.ControlledInputs("Port_31").Clear
    pGeometricConstruction.ControlledInputs("Port_31").AddElements oFacePortExtractor19.ControlledInputs("Port")

    Dim oLineParallel20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel20 = oGCFactory.CreateEntity("LineParallel", pPOM, "020-LineParallel")
    oLineParallel20.Inputs("Line").Add oLineFromCrossSection9
    oLineParallel20.Inputs("Plane").Add oExtractPorts14.Output("Support", 1)
    oLineParallel20.Inputs("TrackPoint").Add oPointAtCurveEnd8
    oLineParallel20.Parameter("Distance") = oParamDistConstant3
    oLineParallel20.Parameter("TrackFlag") = 1
    oLineParallel20.Evaluate

    Dim oLineParallel21 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel21 = oGCFactory.CreateEntity("LineParallel", pPOM, "021-LineParallel")
    oLineParallel21.Inputs("Line").Add oLineFromCrossSection10
    oLineParallel21.Inputs("Plane").Add oExtractPorts15.Output("Support", 1)
    oLineParallel21.Inputs("TrackPoint").Add oPointAtCurveEnd8
    oLineParallel21.Parameter("Distance") = oParamDistConstant3
    oLineParallel21.Parameter("TrackFlag") = 1
    oLineParallel21.Evaluate

    Dim oPointAtCurveMiddle22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle22 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "022-PointAtCurveMiddle")
    oPointAtCurveMiddle22.Inputs("Curve").Add oLineParallel20
    oPointAtCurveMiddle22.Evaluate

    Dim oPointAtCurveStart23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart23 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "023-PointAtCurveStart")
    oPointAtCurveStart23.Inputs("Curve").Add oLineParallel20
    oPointAtCurveStart23.Inputs("Surface").Add oFacePortExtractor16
    oPointAtCurveStart23.Evaluate

    Dim oPointAtCurveEnd24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd24 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "024-PointAtCurveEnd")
    oPointAtCurveEnd24.Inputs("Curve").Add oLineParallel20
    oPointAtCurveEnd24.Inputs("Surface").Add oFacePortExtractor16
    oPointAtCurveEnd24.Evaluate

    Dim oPointAtCurveMiddle25 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle25 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "025-PointAtCurveMiddle")
    oPointAtCurveMiddle25.Inputs("Curve").Add oLineParallel21
    oPointAtCurveMiddle25.Evaluate

    Dim oPointAtCurveStart26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart26 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "026-PointAtCurveStart")
    oPointAtCurveStart26.Inputs("Curve").Add oLineParallel21
    oPointAtCurveStart26.Inputs("Surface").Add oFacePortExtractor17
    oPointAtCurveStart26.Evaluate

    Dim oPointAtCurveEnd27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd27 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "027-PointAtCurveEnd")
    oPointAtCurveEnd27.Inputs("Curve").Add oLineParallel21
    oPointAtCurveEnd27.Inputs("Surface").Add oFacePortExtractor17
    oPointAtCurveEnd27.Evaluate

    Dim oPlaneByPointNormal28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal28 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "028-PlaneByPointNormal")
    oPlaneByPointNormal28.Inputs("Point").Add oPointAtCurveMiddle22
    oPlaneByPointNormal28.Inputs("Line").Add oExtractPorts14.Output("MemberAxis", 1)
    oPlaneByPointNormal28.Parameter("Range") = 1#
    oPlaneByPointNormal28.Evaluate

    Dim oLineByPoints29 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints29 = oGCFactory.CreateEntity("LineByPoints", pPOM, "029-LineByPoints")
    oLineByPoints29.Inputs("StartPoint").Add oPointAtCurveStart23
    oLineByPoints29.Inputs("EndPoint").Add oPointAtCurveEnd24
    oLineByPoints29.Evaluate

    Dim oLineByPoints30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints30 = oGCFactory.CreateEntity("LineByPoints", pPOM, "030-LineByPoints")
    oLineByPoints30.Inputs("StartPoint").Add oPointAtCurveStart26
    oLineByPoints30.Inputs("EndPoint").Add oPointAtCurveEnd27
    oLineByPoints30.Evaluate

    Dim oPointAtCurveMiddle31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle31 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "031-PointAtCurveMiddle")
    oPointAtCurveMiddle31.Inputs("Curve").Add oLineByPoints29
    oPointAtCurveMiddle31.Evaluate

    Dim oLineParallel32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel32 = oGCFactory.CreateEntity("LineParallel", pPOM, "032-LineParallel")
    oLineParallel32.Inputs("Line").Add oLineParallel20
    oLineParallel32.Inputs("Plane").Add oPlaneByPointNormal28
    oLineParallel32.Inputs("TrackPoint").Add oPointAtCurveMiddle25
    oLineParallel32.Parameter("Distance") = 0.015
    oLineParallel32.Parameter("TrackFlag") = 2
    oLineParallel32.Evaluate

    Dim oLineParallel33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel33 = oGCFactory.CreateEntity("LineParallel", pPOM, "033-LineParallel")
    oLineParallel33.Inputs("Line").Add oLineParallel21
    oLineParallel33.Inputs("Plane").Add oPlaneByPointNormal28
    oLineParallel33.Inputs("TrackPoint").Add oPointAtCurveMiddle22
    oLineParallel33.Parameter("Distance") = 0.015
    oLineParallel33.Parameter("TrackFlag") = 2
    oLineParallel33.Evaluate

    Dim oPointByCurveAndSurface34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface34 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "034-PointByCurveAndSurface")
    oPointByCurveAndSurface34.Inputs("Curve1").Add oLineByPoints29
    oPointByCurveAndSurface34.Inputs("Surface2").Add oFaceFromCrossSection18
    oPointByCurveAndSurface34.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface34.Evaluate

    Dim oPointByCurveAndSurface35 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface35 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "035-PointByCurveAndSurface")
    oPointByCurveAndSurface35.Inputs("Curve1").Add oLineByPoints30
    oPointByCurveAndSurface35.Inputs("Surface2").Add oFaceFromCrossSection18
    oPointByCurveAndSurface35.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface35.Evaluate

    Dim oPointByCurveAndSurface36 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface36 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "036-PointByCurveAndSurface")
    oPointByCurveAndSurface36.Inputs("Curve1").Add oLineByPoints29
    oPointByCurveAndSurface36.Inputs("Surface2").Add oFacePortExtractor19
    oPointByCurveAndSurface36.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface36.Evaluate

    Dim oParamDistMeasureParallel37 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel37 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "037-ParamDistMeasureParallel")
    oParamDistMeasureParallel37.Inputs("Graphics").Add oPointAtCurveMiddle31
    oParamDistMeasureParallel37.Inputs("Graphics").Add oPointByCurveAndSurface36
    oParamDistMeasureParallel37.Parameter("Value") = 0.0075
    oParamDistMeasureParallel37.Evaluate

    Dim oPointAlongCurve38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve38 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "038-PointAlongCurve")
    oPointAlongCurve38.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve38.Inputs("Point").Add oPointByCurveAndSurface34
    oPointAlongCurve38.Inputs("TrackPoint").Add oPointAtCurveMiddle22
    oPointAlongCurve38.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve38.Parameter("TrackFlag") = 1
    oPointAlongCurve38.Evaluate

    Dim oPointAlongCurve39 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve39 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "039-PointAlongCurve")
    oPointAlongCurve39.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve39.Inputs("Point").Add oPointByCurveAndSurface34
    oPointAlongCurve39.Inputs("TrackPoint").Add oPointAtCurveMiddle22
    oPointAlongCurve39.Parameter("Distance") = oParamDistConstant6
    oPointAlongCurve39.Parameter("TrackFlag") = 1
    oPointAlongCurve39.Evaluate

    Dim oPointAlongCurve40 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve40 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "040-PointAlongCurve")
    oPointAlongCurve40.Inputs("Curve").Add oLineByPoints30
    oPointAlongCurve40.Inputs("Point").Add oPointByCurveAndSurface35
    oPointAlongCurve40.Inputs("TrackPoint").Add oPointAtCurveMiddle25
    oPointAlongCurve40.Parameter("Distance") = oParamDistConstant6
    oPointAlongCurve40.Parameter("TrackFlag") = 1
    oPointAlongCurve40.Evaluate

    Dim oLineByPoints41 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints41 = oGCFactory.CreateEntity("LineByPoints", pPOM, "041-LineByPoints")
    oLineByPoints41.Inputs("StartPoint").Add oPointAlongCurve40
    oLineByPoints41.Inputs("EndPoint").Add oPointAlongCurve39
    oLineByPoints41.Evaluate

    Dim oPointAlongCurve42 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve42 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "042-PointAlongCurve")
    oPointAlongCurve42.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve42.Inputs("Point").Add oPointAtCurveMiddle22
    oPointAlongCurve42.Inputs("TrackPoint").Add oPointByCurveAndSurface34
    oPointAlongCurve42.Parameter("Distance") = oParamDistMeasureParallel37
    oPointAlongCurve42.Parameter("TrackFlag") = 1
    oPointAlongCurve42.Evaluate

    Dim oPointAlongCurve43 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve43 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "043-PointAlongCurve")
    oPointAlongCurve43.Inputs("Curve").Add oLineByPoints30
    oPointAlongCurve43.Inputs("Point").Add oPointAtCurveMiddle25
    oPointAlongCurve43.Inputs("TrackPoint").Add oPointByCurveAndSurface35
    oPointAlongCurve43.Parameter("Distance") = oParamDistMeasureParallel37
    oPointAlongCurve43.Parameter("TrackFlag") = 1
    oPointAlongCurve43.Evaluate

    Dim oLineByPoints44 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints44 = oGCFactory.CreateEntity("LineByPoints", pPOM, "044-LineByPoints")
    oLineByPoints44.Inputs("StartPoint").Add oPointAlongCurve42
    oLineByPoints44.Inputs("EndPoint").Add oPointAlongCurve43
    oLineByPoints44.Evaluate

    Dim oSurfByLinearExtrusion45 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion45 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "045-SurfByLinearExtrusion")
    oSurfByLinearExtrusion45.Inputs("PlanarCrossSection").Add oLineByPoints44
    oSurfByLinearExtrusion45.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion45.Evaluate

    Dim oPointAlongCurve46 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve46 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "046-PointAlongCurve")
    oPointAlongCurve46.Inputs("Curve").Add oLineByPoints41
    oPointAlongCurve46.Inputs("Point").Add oPointAlongCurve40
    oPointAlongCurve46.Inputs("TrackPoint").Add oPointAlongCurve39
    oPointAlongCurve46.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve46.Parameter("TrackFlag") = 1
    oPointAlongCurve46.Evaluate

    Dim oPointAlongCurve47 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve47 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "047-PointAlongCurve")
    oPointAlongCurve47.Inputs("Curve").Add oLineByPoints44
    oPointAlongCurve47.Inputs("Point").Add oPointAlongCurve42
    oPointAlongCurve47.Inputs("TrackPoint").Add oPointAlongCurve43
    oPointAlongCurve47.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve47.Parameter("TrackFlag") = 1
    oPointAlongCurve47.Evaluate

    Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "048-PointAlongCurve")
    oPointAlongCurve48.Inputs("Curve").Add oLineByPoints29
    oPointAlongCurve48.Inputs("Point").Add oPointAlongCurve42
    oPointAlongCurve48.Inputs("TrackPoint").Add oPointByCurveAndSurface34
    oPointAlongCurve48.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve48.Parameter("TrackFlag") = 1
    oPointAlongCurve48.Evaluate

    Dim oPointAlongCurve49 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
    oPointAlongCurve49.Inputs("Curve").Add oLineByPoints44
    oPointAlongCurve49.Inputs("Point").Add oPointAlongCurve43
    oPointAlongCurve49.Inputs("TrackPoint").Add oPointAlongCurve42
    oPointAlongCurve49.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve49.Parameter("TrackFlag") = 1
    oPointAlongCurve49.Evaluate

    Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
    oPointAlongCurve50.Inputs("Curve").Add oLineByPoints30
    oPointAlongCurve50.Inputs("Point").Add oPointAlongCurve43
    oPointAlongCurve50.Inputs("TrackPoint").Add oPointByCurveAndSurface35
    oPointAlongCurve50.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve50.Parameter("TrackFlag") = 1
    oPointAlongCurve50.Evaluate

    Dim oVectorNormalToSurface51 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface51 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "051-VectorNormalToSurface")
    oVectorNormalToSurface51.Inputs("Surface").Add oFacePortExtractor16
    oVectorNormalToSurface51.Inputs("Point").Add oPointAlongCurve48
    oVectorNormalToSurface51.Parameter("Range") = 0.2
    oVectorNormalToSurface51.Parameter("Orientation") = 2
    oVectorNormalToSurface51.Parameter("TrackFlag") = 2
    oVectorNormalToSurface51.Evaluate

    Dim oVectorNormalToSurface52 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface52 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "052-VectorNormalToSurface")
    oVectorNormalToSurface52.Inputs("Surface").Add oFacePortExtractor17
    oVectorNormalToSurface52.Inputs("Point").Add oPointAlongCurve50
    oVectorNormalToSurface52.Parameter("Range") = 0.2
    oVectorNormalToSurface52.Parameter("Orientation") = 2
    oVectorNormalToSurface52.Parameter("TrackFlag") = 2
    oVectorNormalToSurface52.Evaluate

    Dim oLineAtAngle53 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle53 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "053-LineAtAngle")
    oLineAtAngle53.Inputs("Line").Add oVectorNormalToSurface51
    oLineAtAngle53.Inputs("Point").Add oPointAlongCurve48
    oLineAtAngle53.Inputs("Plane").Add oPlaneByPointNormal28
    oLineAtAngle53.Inputs("TrackPoint").Add oPointAlongCurve47
    oLineAtAngle53.Parameter("Angle") = 0.7853982
    oLineAtAngle53.Parameter("TrackFlag") = 1
    oLineAtAngle53.Evaluate

    Dim oLineAtAngle54 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle54 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "054-LineAtAngle")
    oLineAtAngle54.Inputs("Line").Add oVectorNormalToSurface52
    oLineAtAngle54.Inputs("Point").Add oPointAlongCurve50
    oLineAtAngle54.Inputs("Plane").Add oPlaneByPointNormal28
    oLineAtAngle54.Inputs("TrackPoint").Add oPointAlongCurve49
    oLineAtAngle54.Parameter("Angle") = 0.7853982
    oLineAtAngle54.Parameter("TrackFlag") = 1
    oLineAtAngle54.Evaluate

    Dim oVectorNormalToSurface55 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface55 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "055-VectorNormalToSurface")
    oVectorNormalToSurface55.Inputs("Surface").Add oSurfByLinearExtrusion45
    oVectorNormalToSurface55.Inputs("Point").Add oPointAlongCurve47
    oVectorNormalToSurface55.Inputs("TrackPoint").Add oPointByCurveAndSurface34
    oVectorNormalToSurface55.Parameter("Range") = 0.1
    oVectorNormalToSurface55.Parameter("Orientation") = 1
    oVectorNormalToSurface55.Parameter("TrackFlag") = 2
    oVectorNormalToSurface55.Evaluate

    Dim oVectorNormalToSurface56 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface56 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "056-VectorNormalToSurface")
    oVectorNormalToSurface56.Inputs("Surface").Add oSurfByLinearExtrusion45
    oVectorNormalToSurface56.Inputs("Point").Add oPointAlongCurve49
    oVectorNormalToSurface56.Inputs("TrackPoint").Add oPointByCurveAndSurface35
    oVectorNormalToSurface56.Parameter("Range") = 0.1
    oVectorNormalToSurface56.Parameter("Orientation") = 1
    oVectorNormalToSurface56.Parameter("TrackFlag") = 2
    oVectorNormalToSurface56.Evaluate

    Dim oLineByPoints57 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints57 = oGCFactory.CreateEntity("LineByPoints", pPOM, "057-LineByPoints")
    oLineByPoints57.Inputs("StartPoint").Add oPointAlongCurve38
    oLineByPoints57.Inputs("EndPoint").Add oPointAlongCurve46
    oLineByPoints57.Evaluate

    Dim oLineByPoints58 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints58 = oGCFactory.CreateEntity("LineByPoints", pPOM, "058-LineByPoints")
    oLineByPoints58.Inputs("StartPoint").Add oPointAlongCurve46
    oLineByPoints58.Inputs("EndPoint").Add oPointAlongCurve40
    oLineByPoints58.Evaluate

    Dim oLineByPoints59 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints59 = oGCFactory.CreateEntity("LineByPoints", pPOM, "059-LineByPoints")
    oLineByPoints59.Inputs("StartPoint").Add oPointAlongCurve48
    oLineByPoints59.Inputs("EndPoint").Add oPointAlongCurve50
    oLineByPoints59.Evaluate

    Dim oPointAlongCurve60 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve60 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "060-PointAlongCurve")
    oPointAlongCurve60.Inputs("Curve").Add oLineByPoints59
    oPointAlongCurve60.Inputs("Point").Add oPointAlongCurve50
    oPointAlongCurve60.Inputs("TrackPoint").Add oPointAlongCurve48
    oPointAlongCurve60.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve60.Parameter("TrackFlag") = 1
    oPointAlongCurve60.Evaluate

    Dim oPointAlongCurve61 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve61 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "061-PointAlongCurve")
    oPointAlongCurve61.Inputs("Curve").Add oLineByPoints59
    oPointAlongCurve61.Inputs("Point").Add oPointAlongCurve48
    oPointAlongCurve61.Inputs("TrackPoint").Add oPointAlongCurve50
    oPointAlongCurve61.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve61.Parameter("TrackFlag") = 1
    oPointAlongCurve61.Evaluate

    Dim oLineAtAngle62 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle62 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "062-LineAtAngle")
    oLineAtAngle62.Inputs("Line").Add oVectorNormalToSurface55
    oLineAtAngle62.Inputs("Point").Add oPointAlongCurve47
    oLineAtAngle62.Inputs("Plane").Add oPlaneByPointNormal28
    oLineAtAngle62.Inputs("TrackPoint").Add oPointAlongCurve48
    oLineAtAngle62.Parameter("Angle") = 0.7853982
    oLineAtAngle62.Parameter("TrackFlag") = 1
    oLineAtAngle62.Evaluate

    Dim oLineAtAngle63 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle63 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "063-LineAtAngle")
    oLineAtAngle63.Inputs("Line").Add oVectorNormalToSurface56
    oLineAtAngle63.Inputs("Point").Add oPointAlongCurve49
    oLineAtAngle63.Inputs("Plane").Add oPlaneByPointNormal28
    oLineAtAngle63.Inputs("TrackPoint").Add oPointAlongCurve50
    oLineAtAngle63.Parameter("Angle") = 0.7853982
    oLineAtAngle63.Parameter("TrackFlag") = 1
    oLineAtAngle63.Evaluate

    Dim oLineExtend64 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend64 = oGCFactory.CreateEntity("LineExtend", pPOM, "064-LineExtend")
    oLineExtend64.Inputs("Line").Add oLineByPoints57
    oLineExtend64.Inputs("Curve1").Add oLineParallel32
    oLineExtend64.Inputs("Curve2").Add oLineByPoints41
    oLineExtend64.Parameter("Context1") = 1
    oLineExtend64.Parameter("Context2") = 3
    oLineExtend64.Evaluate

    Dim oLineExtend65 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend65 = oGCFactory.CreateEntity("LineExtend", pPOM, "065-LineExtend")
    oLineExtend65.Inputs("Line").Add oLineByPoints58
    oLineExtend65.Inputs("Curve1").Add oLineParallel33
    oLineExtend65.Inputs("Curve2").Add oLineByPoints57
    oLineExtend65.Parameter("Context1") = 1
    oLineExtend65.Parameter("Context2") = 3
    oLineExtend65.Evaluate
    
    Dim oArcByCenter66 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCenter66 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "066-ArcByCenter")
    oArcByCenter66.Inputs("Center").Add oPointAlongCurve61
    oArcByCenter66.Inputs("StartPoint").Add oPointAlongCurve48
    oArcByCenter66.Inputs("EndPoint").Add oPointAlongCurve47
    oArcByCenter66.Parameter("SweepAngle") = 1
    oArcByCenter66.Parameter("TrackFlag") = 1
    oArcByCenter66.Evaluate

    Dim oArcByCenter67 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByCenter67 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "067-ArcByCenter")
    oArcByCenter67.Inputs("Center").Add oPointAlongCurve60
    oArcByCenter67.Inputs("StartPoint").Add oPointAlongCurve50
    oArcByCenter67.Inputs("EndPoint").Add oPointAlongCurve49
    oArcByCenter67.Parameter("SweepAngle") = 1
    oArcByCenter67.Parameter("TrackFlag") = 1
    oArcByCenter67.Evaluate

    Dim oCpxStringByCurves68 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves68 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "068-CpxStringByCurves")
    oCpxStringByCurves68.Inputs("Curves").Add oLineExtend64
    oCpxStringByCurves68.Inputs("Curves").Add oLineExtend65
    oCpxStringByCurves68.Evaluate

    Dim oCpxStringByCurves69 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves69 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "069-CpxStringByCurves")
    oCpxStringByCurves69.Inputs("Curves").Add oLineAtAngle62
    oCpxStringByCurves69.Inputs("Curves").Add oArcByCenter66
    oCpxStringByCurves69.Inputs("Curves").Add oLineAtAngle53
    oCpxStringByCurves69.Evaluate

    Dim oCpxStringByCurves70 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves70 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "070-CpxStringByCurves")
    oCpxStringByCurves70.Inputs("Curves").Add oLineAtAngle63
    oCpxStringByCurves70.Inputs("Curves").Add oArcByCenter67
    oCpxStringByCurves70.Inputs("Curves").Add oLineAtAngle54
    oCpxStringByCurves70.Evaluate

    Dim oSurfByLinearExtrusion71 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion71 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "071-SurfByLinearExtrusion")
    oSurfByLinearExtrusion71.Inputs("PlanarCrossSection").Add oCpxStringByCurves69
    oSurfByLinearExtrusion71.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion71.Evaluate

    Dim oSurfByLinearExtrusion72 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion72 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "072-SurfByLinearExtrusion")
    oSurfByLinearExtrusion72.Inputs("PlanarCrossSection").Add oCpxStringByCurves70
    oSurfByLinearExtrusion72.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion72.Evaluate
    
    Dim oSurfByLinearExtrusion73 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion73 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "073-SurfByLinearExtrusion")
    oSurfByLinearExtrusion73.Inputs("PlanarCrossSection").Add oCpxStringByCurves68
    oSurfByLinearExtrusion73.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion73.Evaluate
    
    Dim oSurfFromGType74 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType74 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "074-SurfFromGType")
    oSurfFromGType74.Inputs("Surface").Add oPlaneByPointNormal28
    oSurfFromGType74.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Outputs("Boundary").Clear
    oGCMacro.Output("Support", 1) = oSurfFromGType74.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion73.Output
    If bCreateArcOnAPS = True Then
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion71.Output
        oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion72.Output
    End If
    oGCMacro.Output("Boundary", "FacePort12") = oExtractPorts15.Output("Support", 1)
    oGCMacro.Output("Boundary", "FacePort13") = oExtractPorts13.Output("Support", 1)
    oGCMacro.Output("Boundary", "FacePort11") = oExtractPorts14.Output("Support", 1)
        
End Sub

