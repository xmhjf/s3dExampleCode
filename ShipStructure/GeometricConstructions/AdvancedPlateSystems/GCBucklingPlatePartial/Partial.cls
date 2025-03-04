VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Partial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010-13, Intergraph Corporation. All rights reserved.
'
'   File:           Partial.cls
'   ProgID:         GCBucklingPlatePartial.Partial
'   Author:         Alligators Team(India)
'   Creation Date:  Tuesday, Sep 29 2010
'   Description:
'      This macro places an advanced plate on either a standard
'      member or on a built up member. User can select 'Near/Far' to place APS
'      on one side (either on web-left or web-right)
'      respectively. Also, user can select top/bottom to attach APS to either the
'      top flange or bottom flange.

'      Two sub-routines are used in 'Evaluate' mehtod (one for standard member and one
'      for built-up)
'
' *** NOTE *** : while placing this APS on standard member user may specify radius value
'       (i) Very small value 0.001 mm for zero radius APS
'                         or
'       (ii) Specify radius value greater than or equal to radius of the member.
'
'   Change History:
'   dd.mmm.yyyy     who              change description
'   -----------     ---              ------------------
'
'   29.Sep.2010     svsmylav         CR-CP-181618 Creation.
'   20.Feb.2013     svsmylav         TR-CP-223312 Following changes are done: For BUHaunch designed member Faceport extractor GC
'                                    did not work correctly, so replaced it with 'FacePortExtractor0' GC using PlateSystem as input.
'                                    Also created two planes 'oPlaneByPointNormal24b' and 'oPlaneByPointNormal24c' which are used by
'                                    by downstream GCs for construction(instead of curved surfaces from face ports mentioned above).
'                                    Cleanup:
'                                    Deleted 'Both' option for TrackFlag from Initialize method. So deleted 'Evaluate_BothSidedPlateOnStdMember'
'                                    and Evaluate_BothSidedPlateOnBUI methods and the method calls from 'Evaluate' method.
'   01.Mar.2013     svsmylav         CR-219579: Changes are done to to create LCs automatically: Details:- in subroutine for std. member
'                                    oLineByPoints72 and oLineByPoints75 are extended beyond the member port and FACEPORT_WEB is added as boundary.
'                                    Also, PORT, FACEPORT_WEB and FACEPORT_FLANGE constants are defined for better readability.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGCToDoDelegate

Private m_bCreateArcOnAPS As Boolean
Private m_oRegularAPSHelper As Object
Private m_oAdvancedDebug As New AdvancedDebug
Private Const TOLERANCE As Double = 0.0001
'
' definition of the error codes
'
Private Enum Errors
    MISSING_MEMBER = 1
    INVALID_OFFSET = 2
    INVALID_HEIGHT1 = 3
    INVALID_MEMBER = 4
    INVALID_SIDE = 5
    INVALID_WELDTOE1 = 6
    INVALID_WELDTOE2 = 7
    INVALID_HEIGHT2 = 8
End Enum
'Defined below constants to improve readability
Const PORT = "Port_1_1"
Const FACEPORT_WEB = "FacePort12"
Const FACEPORT_FLANGE = "Port11"

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
    Call m_oAdvancedDebug.ExitMethod
End Sub

Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call m_oRegularAPSHelper.PreEvaluate(pGeometricConstruction, pPOM)
    ' customized evaluation
    Call Evaluate(pGeometricConstruction, pPOM)
    Call m_oRegularAPSHelper.PostEvaluate(pGeometricConstruction, pPOM)
    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oRegularAPSHelper.ProcessError(pGeometricConstruction, Err.Number)
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
' to be oveeriden by print graph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a RolledMember or a DesignedMember", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")

    Call pGeometricConstructionDefinition.AddControlledInput(PORT, "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput(FACEPORT_FLANGE)
    Call pGeometricConstructionDefinition.AddControlledInput(FACEPORT_WEB)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    
    'Inputs
    Call pGeometricConstructionDefinition.AddParameter("Offset", "Offset", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Height1", "Height1", 8, 1, 59, 0, 0, 0.15)
    Call pGeometricConstructionDefinition.AddParameter("Height2", "Height2", 8, 1, 59, 0, 0, 0.06)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe1", "WeldToe1", 8, 1, 59, 0, 0, 0.01)
    Call pGeometricConstructionDefinition.AddParameter("WeldToe2", "WeldToe2", 8, 1, 59, 0, 0, 0.015)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 0.01)
    Call pGeometricConstructionDefinition.AddParameter("CSOrientation", "Top/Bottom Flange", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("CSOrientation", "Top", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("CSOrientation", "Bottom", 2)
    Call pGeometricConstructionDefinition.AddParameter("TrackFlag", "Near/Far", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Near", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("TrackFlag", "Far", 2)
    
    pGeometricConstructionDefinition.Behavior = 0
    'The error message that should be displayed is defined when invalid parameters are given as input
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_MEMBER, "Missing Member", "Mising member part for the buckling plate.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_OFFSET, "Invalid Offset", "Offset should be within the member part length")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_HEIGHT1, "Invalid Height1", "Height should be within the member part depth")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_MEMBER, "Invalid Member", "The selected member part is not a valid type for this buckling plate.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_SIDE, "Invalid Side", "Can not place buckling plate on this side of the member cross-section.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_WELDTOE1, "Invalid WeldToe1", "WeldToe1 should be within the member part width.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_WELDTOE2, "Invalid WeldToe2", "WeldToe2 should be within the member part width.")
    Call pGeometricConstructionDefinition.AddErrorValue(INVALID_HEIGHT2, "Invalid Height2", "Height should be within the member part depth")
End Sub

Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    'Assumption: Radius value less than 0.1 mm is considered to be zero radius.
    If pGeometricConstruction.Parameter("Radius") < 0.0001 Then
        m_bCreateArcOnAPS = False 'Do not create arc
    Else
        m_bCreateArcOnAPS = True  'Create arc
    End If
    If TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSMemberPartPrismatic Then
        'Standard member
        Evaluate_OneSidePlateOnStdMember pGeometricConstruction, pPOM
    ElseIf TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSDesignedMember Then
        'Built-up
         Evaluate_OneSidePlateOnBUI pGeometricConstruction, pPOM
    End If
End Sub

Private Sub Evaluate_OneSidePlateOnBUI(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)

    Dim lErrorNumber As Long
    If pGeometricConstruction.Inputs("MemberPart").Count <> 1 Then Err.Raise MISSING_MEMBER
    Dim oISPSDesignedMember As ISPSDesignedMember
    On Error Resume Next
    Set oISPSDesignedMember = pGeometricConstruction.Inputs("MemberPart").Item(1)
    On Error GoTo 0
    If (oISPSDesignedMember Is Nothing) Then Err.Raise INVALID_MEMBER
    
    'Check any parameter is -ve
    If (pGeometricConstruction.Parameter("Offset") < 0) Then Err.Raise INVALID_OFFSET
    If (pGeometricConstruction.Parameter("WeldToe1") < 0) Then Err.Raise INVALID_WELDTOE1
    If (pGeometricConstruction.Parameter("WeldToe2") < 0) Then Err.Raise INVALID_WELDTOE2
    If (pGeometricConstruction.Parameter("Height1") < 0) Then Err.Raise INVALID_HEIGHT1
    If (pGeometricConstruction.Parameter("Height2") < 0) Then Err.Raise INVALID_HEIGHT2

    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    'Need to get top and bottom flange ports
    Dim oDP As IJDesignParent
    Dim oChildren As IJDObjectCollection
    Set oDP = pGeometricConstruction.Inputs("MemberPart").Item(1)
    oDP.GetChildren oChildren 'Got plate systems
    Dim oPlateSystem As Object
    Dim oPlate As IJPlate
       
    Dim oReqFlgPltSys1 As Object
    Dim oReqFlgPltSys2 As Object
    Dim bIsBottom As Boolean
    bIsBottom = False
    For Each oPlateSystem In oChildren
        If TypeOf oPlateSystem Is IJPlate Then
            Set oPlate = oPlateSystem
            If oPlate.plateType = 15 Then 'Ensure that current plate system is Flange Plate System, type = 15
                If pGeometricConstruction.Parameter("CSOrientation") = 2 Then
                    bIsBottom = True
                    '"CSOrientation" - "Bottom" Case
                    If oPlate.NamingCategory = 2 Then       'NameCatTFlange = 2 Top-Flange
                        Set oReqFlgPltSys1 = oPlateSystem
                    Else                                    'NameCatFlange = 0 'Bottom-Flange
                        Set oReqFlgPltSys2 = oPlateSystem
                    End If
                Else
                    '"CSOrientation" - "Top" Case: ***swap*** systems 1 and 2
                    If oPlate.NamingCategory = 2 Then
                        Set oReqFlgPltSys2 = oPlateSystem
                    Else
                        Set oReqFlgPltSys1 = oPlateSystem
                    End If
                End If
            End If
        End If
    Next oPlateSystem
    Set oPlateSystem = Nothing
    Set oChildren = Nothing
    Set oDP = Nothing

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs(PORT).Clear
    pGeometricConstruction.ControlledInputs(PORT).AddElements oLineAxisPortExtractor1.ControlledInputs("Port")

    Dim oLine As IJLine
    On Error Resume Next
    Set oLine = oLineAxisPortExtractor1.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Offset") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Offset") > oLine.Length Then Err.Raise INVALID_OFFSET
    End If

    Dim oParamDouble2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble2 = oGCFactory.CreateEntity("ParamDouble", pPOM, "002-ParamDouble")
    oParamDouble2.Parameter("Value") = 0.5
    oParamDouble2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Height1")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Height2")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = 0.015
    oParamDistConstant5.Evaluate

    Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "006-ParamDistConstant")
    oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant6.Evaluate

    Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "007-ParamDistConstant")
    oParamDistConstant7.Parameter("Value") = pGeometricConstruction.Parameter("Radius")
    oParamDistConstant7.Evaluate

    Dim oParamDistConstant8 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant8 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "008-ParamDistConstant")
    oParamDistConstant8.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe1")
    oParamDistConstant8.Evaluate

    Dim oParamDistConstant9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant9 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "009-ParamDistConstant")
    oParamDistConstant9.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe2")
    oParamDistConstant9.Evaluate

    Dim oPointAtCurveStart10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart10 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "010-PointAtCurveStart")
    oPointAtCurveStart10.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveStart10.Evaluate

    Dim oPointAtCurveEnd11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd11 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "011-PointAtCurveEnd")
    oPointAtCurveEnd11.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveEnd11.Evaluate

    Dim oLineFromCrossSection12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection12 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "012-LineFromCrossSection")
    oLineFromCrossSection12.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection12.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection12.Parameter("EdgeSelector") = 513
    oLineFromCrossSection12.Evaluate
    
    'Validate the weldToe2
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection12.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe2") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe2") > oLine.Length Then Err.Raise INVALID_WELDTOE2
    End If


    Dim oCSFromMember13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember13 = oGCFactory.CreateEntity("CSFromMember", pPOM, "013-CSFromMember")
    oCSFromMember13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember13.Inputs("Point").Add oPointAtCurveStart10
    oCSFromMember13.Evaluate

    Dim oLineFromCrossSection14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection14 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "014-LineFromCrossSection")
    oLineFromCrossSection14.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection14.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection14.Parameter("EdgeSelector") = 258
    oLineFromCrossSection14.Evaluate

    'Validate the height1
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection14.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Height1") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Height1") > oLine.Length Then Err.Raise INVALID_HEIGHT1
    End If

    'Validate the height2
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection14.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Height2") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Height2") > oLine.Length Then Err.Raise INVALID_HEIGHT2
    End If

    Dim oLineFromCrossSection15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection15 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "015-LineFromCrossSection")
    oLineFromCrossSection15.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection15.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection15.Parameter("EdgeSelector") = 514
    oLineFromCrossSection15.Evaluate
    
    'Validate the weldToe1
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection15.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe1") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe1") > oLine.Length Then Err.Raise INVALID_WELDTOE1
    End If

    Dim oFacePortExtractor16 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "016-FacePortExtractor")
    oFacePortExtractor16.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor16.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor16.Parameter("GeometrySelector") = 4
    oFacePortExtractor16.Parameter("FacesContext") = 5
    oFacePortExtractor16.Parameter("LookingAxis") = 2
    oFacePortExtractor16.Parameter("SurfaceType") = 1
    oFacePortExtractor16.Parameter("TrackFlag") = 1
    oFacePortExtractor16.Parameter("Offset") = 0#
    oFacePortExtractor16.Evaluate

    Dim oFacePortExtractor17 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor17 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "017-FacePortExtractor")
    oFacePortExtractor17.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor17.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor17.Parameter("GeometrySelector") = 4
    oFacePortExtractor17.Parameter("FacesContext") = 5
    oFacePortExtractor17.Parameter("LookingAxis") = 2
    oFacePortExtractor17.Parameter("SurfaceType") = 1
    oFacePortExtractor17.Parameter("TrackFlag") = 2
    oFacePortExtractor17.Parameter("Offset") = 0#
    oFacePortExtractor17.Evaluate

    Dim oCSByCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS18 = oGCFactory.CreateEntity("CSByCS", pPOM, "018-CSByCS")
    oCSByCS18.Inputs("CoordinateSystem").Add oCSFromMember13
    oCSByCS18.Inputs("Origin").Add oPointAtCurveStart10
    oCSByCS18.Parameter("AxesRoles") = 1
    oCSByCS18.Parameter("CSOrientation") = pGeometricConstruction.Parameter("CSOrientation")
    oCSByCS18.Parameter("TrackFlag") = 1
    oCSByCS18.Evaluate

    Dim oFaceFromCrossSection19 As SP3DGeometricConstruction.GeometricConstruction
    Set oFaceFromCrossSection19 = oGCFactory.CreateEntity("FaceFromCrossSection", pPOM, "019-FaceFromCrossSection")
    oFaceFromCrossSection19.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFaceFromCrossSection19.Inputs("CoordinateSystem").Add oCSFromMember13
    oFaceFromCrossSection19.Parameter("LookingAxis") = 2
    oFaceFromCrossSection19.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFaceFromCrossSection19.Parameter("Offset") = 0#
    oFaceFromCrossSection19.Evaluate

    Dim oFacePortExtractor20 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor20 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "020-FacePortExtractor0")
    oFacePortExtractor20.Inputs("Connectable").Add oReqFlgPltSys1
    oFacePortExtractor20.Parameter("Offset") = 0#
    oFacePortExtractor20.Evaluate
    Set oReqFlgPltSys1 = Nothing
    
    Dim oFacePortExtractor21 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor21 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "021-FacePortExtractor0")
    oFacePortExtractor21.Inputs("Connectable").Add oReqFlgPltSys2
    oFacePortExtractor21.Parameter("Offset") = 0#
    oFacePortExtractor21.Evaluate
        'Validate the side
    On Error Resume Next
    lErrorNumber = 0
    oFacePortExtractor21.Evaluate
    lErrorNumber = Err.Number
    On Error GoTo 0
    If lErrorNumber <> 0 Then Err.Raise INVALID_SIDE

    pGeometricConstruction.ControlledInputs(FACEPORT_FLANGE).Clear
    pGeometricConstruction.ControlledInputs(FACEPORT_FLANGE).AddElements oFacePortExtractor21.ControlledInputs("Port")
    Set oReqFlgPltSys2 = Nothing

    Dim oFacePortExtractor16b As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16b = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "016b-FacePortExtractor")
    oFacePortExtractor16b.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor16b.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor16b.Parameter("GeometrySelector") = 4
    oFacePortExtractor16b.Parameter("FacesContext") = 2
    oFacePortExtractor16b.Parameter("LookingAxis") = 2
    oFacePortExtractor16b.Parameter("SurfaceType") = 1
    oFacePortExtractor16b.Parameter("TrackFlag") = 1
    oFacePortExtractor16b.Parameter("Offset") = 0#
    oFacePortExtractor16b.Evaluate
    
    Dim oFacePortExtractor22 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor22 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "022-FacePortExtractor1")
    oFacePortExtractor22.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor22.Inputs("CoordinateSystem").Add oCSByCS18
    oFacePortExtractor22.Inputs("Face").Add oFacePortExtractor16b
    oFacePortExtractor22.Parameter("GeometrySelector") = 4
    oFacePortExtractor22.Parameter("FacesContext") = 4
    oFacePortExtractor22.Parameter("LookingAxis") = 3
    oFacePortExtractor22.Parameter("SurfaceType") = 1
    oFacePortExtractor22.Parameter("TrackFlag") = 2
    oFacePortExtractor22.Parameter("Offset") = 0#
    oFacePortExtractor22.Evaluate

    Dim oPointByCurveAndSurface23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface23 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "023-PointByCurveAndSurface")
    oPointByCurveAndSurface23.Inputs("Curve1").Add oLineFromCrossSection14
    oPointByCurveAndSurface23.Inputs("Surface2").Add oFacePortExtractor20
    oPointByCurveAndSurface23.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface23.Evaluate

    Dim oPointByCurveAndSurface24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface24 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "024-PointByCurveAndSurface")
    oPointByCurveAndSurface24.Inputs("Curve1").Add oLineFromCrossSection14
    oPointByCurveAndSurface24.Inputs("Surface2").Add oFacePortExtractor21
    oPointByCurveAndSurface24.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface24.Evaluate
    
'   For BUHaunch Face ports 20/21 extracted above could not be used as planes - so below 'oPlaneByPointNormal24b'
'   'oPlaneByPointNormal24c' are used instead of above two Face-ports.
    Dim oPlaneByPointNormal24b As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal24b = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "026b-PlaneByPointNormal")
    oPlaneByPointNormal24b.Inputs("Point").Add oPointByCurveAndSurface23
    oPlaneByPointNormal24b.Inputs("Line").Add oLineFromCrossSection14
    oPlaneByPointNormal24b.Parameter("Range") = 0.999998
    oPlaneByPointNormal24b.Evaluate
    
    Dim oPlaneByPointNormal24c As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal24c = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "026b-PlaneByPointNormal")
    oPlaneByPointNormal24c.Inputs("Point").Add oPointByCurveAndSurface24
    oPlaneByPointNormal24c.Inputs("Line").Add oLineFromCrossSection14
    oPlaneByPointNormal24c.Parameter("Range") = 0.999998
    oPlaneByPointNormal24c.Evaluate

    Dim oLineAtAngle25 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle25 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "025-LineAtAngle")
    oLineAtAngle25.Inputs("Line").Add oLineFromCrossSection15
    oLineAtAngle25.Inputs("Point").Add oPointByCurveAndSurface23
    oLineAtAngle25.Inputs("Plane").Add oPlaneByPointNormal24b
    oLineAtAngle25.Parameter("Angle") = 0#
    oLineAtAngle25.Parameter("TrackFlag") = 1
    oLineAtAngle25.Evaluate

    Dim oLineAtAngle26 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle26 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "026-LineAtAngle")
    oLineAtAngle26.Inputs("Line").Add oLineAtAngle25
    oLineAtAngle26.Inputs("Point").Add oPointByCurveAndSurface24
    oLineAtAngle26.Inputs("Plane").Add oPlaneByPointNormal24c
    oLineAtAngle26.Parameter("Angle") = 0#
    oLineAtAngle26.Parameter("TrackFlag") = 1
    oLineAtAngle26.Evaluate

    Dim oLineParallel27 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel27 = oGCFactory.CreateEntity("LineParallel", pPOM, "027-LineParallel")
    oLineParallel27.Inputs("Line").Add oLineAtAngle26
    oLineParallel27.Inputs("Plane").Add oPlaneByPointNormal24c
    oLineParallel27.Inputs("TrackPoint").Add oPointAtCurveEnd11
    oLineParallel27.Parameter("Distance") = oParamDistConstant6
    oLineParallel27.Parameter("TrackFlag") = 1
    oLineParallel27.Evaluate

    Dim oLineParallel28 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel28 = oGCFactory.CreateEntity("LineParallel", pPOM, "028-LineParallel")
    oLineParallel28.Inputs("Line").Add oLineAtAngle25
    oLineParallel28.Inputs("Plane").Add oPlaneByPointNormal24b
    oLineParallel28.Inputs("TrackPoint").Add oPointAtCurveEnd11
    oLineParallel28.Parameter("Distance") = oParamDistConstant6
    oLineParallel28.Parameter("TrackFlag") = 1
    oLineParallel28.Evaluate

    Dim oPointByCurveAndSurface29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface29 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "029-PointByCurveAndSurface")
    oPointByCurveAndSurface29.Inputs("Curve1").Add oLineParallel28
    oPointByCurveAndSurface29.Inputs("Surface2").Add oFaceFromCrossSection19
    oPointByCurveAndSurface29.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface29.Evaluate

    Dim oPointAtCurveMiddle30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle30 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "030-PointAtCurveMiddle")
    oPointAtCurveMiddle30.Inputs("Curve").Add oLineParallel28
    oPointAtCurveMiddle30.Evaluate

    Dim oParamDistMeasureLength31 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength31 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "031-ParamDistMeasureLength")
    oParamDistMeasureLength31.Inputs("Curve").Add oLineParallel27
    oParamDistMeasureLength31.Parameter("Value") = 0.2
    oParamDistMeasureLength31.Evaluate

    Dim oPointByCurveAndSurface32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface32 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "032-PointByCurveAndSurface")
    oPointByCurveAndSurface32.Inputs("Curve1").Add oLineParallel27
    oPointByCurveAndSurface32.Inputs("Surface2").Add oFaceFromCrossSection19
    oPointByCurveAndSurface32.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface32.Evaluate

    Dim oPointAtCurveMiddle33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle33 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "033-PointAtCurveMiddle")
    oPointAtCurveMiddle33.Inputs("Curve").Add oLineParallel27
    oPointAtCurveMiddle33.Evaluate

    Dim oPointAlongCurve34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve34 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "034-PointAlongCurve")
    oPointAlongCurve34.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve34.Inputs("Point").Add oPointByCurveAndSurface32
    oPointAlongCurve34.Inputs("TrackPoint").Add oPointAtCurveMiddle33
    oPointAlongCurve34.Parameter("Distance") = oParamDistConstant8
    oPointAlongCurve34.Parameter("TrackFlag") = 1
    oPointAlongCurve34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointAtCurveMiddle33
    oLineByPoints35.Inputs("EndPoint").Add oPointAtCurveMiddle30
    oLineByPoints35.Evaluate

    Dim oLineByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints36 = oGCFactory.CreateEntity("LineByPoints", pPOM, "036-LineByPoints")
    oLineByPoints36.Inputs("StartPoint").Add oPointByCurveAndSurface32
    oLineByPoints36.Inputs("EndPoint").Add oPointByCurveAndSurface29
    oLineByPoints36.Evaluate

    Dim oParamDistFunctionWith2Args37 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args37 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "037-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args37.Inputs("Arg1").Add oParamDistMeasureLength31
    oParamDistFunctionWith2Args37.Inputs("Arg2").Add oParamDouble2
    oParamDistFunctionWith2Args37.Parameter("Value") = 0.1
    oParamDistFunctionWith2Args37.Parameter("Operation") = 3
    oParamDistFunctionWith2Args37.Evaluate

    Dim oPointByCurveAndSurface38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface38 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "038-PointByCurveAndSurface")
    oPointByCurveAndSurface38.Inputs("Curve1").Add oLineByPoints35
    oPointByCurveAndSurface38.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface38.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface38.Evaluate

    Dim oPointByCurveAndSurface39 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface39 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "039-PointByCurveAndSurface")
    oPointByCurveAndSurface39.Inputs("Curve1").Add oLineByPoints36
    oPointByCurveAndSurface39.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface39.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface39.Evaluate

    Dim oLineByPoints40 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints40 = oGCFactory.CreateEntity("LineByPoints", pPOM, "040-LineByPoints")
    oLineByPoints40.Inputs("StartPoint").Add oPointByCurveAndSurface39
    oLineByPoints40.Inputs("EndPoint").Add oPointByCurveAndSurface38
    oLineByPoints40.Evaluate

    Dim oPointAlongCurve41 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve41 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "041-PointAlongCurve")
    oPointAlongCurve41.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve41.Inputs("Point").Add oPointByCurveAndSurface38
    oPointAlongCurve41.Inputs("TrackPoint").Add oPointAtCurveMiddle30
    oPointAlongCurve41.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve41.Parameter("TrackFlag") = 1
    oPointAlongCurve41.Evaluate

    Dim oPointAlongCurve42 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve42 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "042-PointAlongCurve")
    oPointAlongCurve42.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve42.Inputs("Point").Add oPointAlongCurve41
    oPointAlongCurve42.Inputs("TrackPoint").Add oPointAtCurveMiddle30
    oPointAlongCurve42.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve42.Parameter("TrackFlag") = 1
    oPointAlongCurve42.Evaluate

    Dim oPointAlongCurve43 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve43 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "043-PointAlongCurve")
    oPointAlongCurve43.Inputs("Curve").Add oLineByPoints36
    oPointAlongCurve43.Inputs("Point").Add oPointByCurveAndSurface39
    oPointAlongCurve43.Inputs("TrackPoint").Add oPointByCurveAndSurface29
    oPointAlongCurve43.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve43.Parameter("TrackFlag") = 1
    oPointAlongCurve43.Evaluate

    Dim oPointAlongCurve44 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve44 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "044-PointAlongCurve")
    oPointAlongCurve44.Inputs("Curve").Add oLineByPoints36
    oPointAlongCurve44.Inputs("Point").Add oPointAlongCurve43
    oPointAlongCurve44.Inputs("TrackPoint").Add oPointByCurveAndSurface29
    oPointAlongCurve44.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve44.Parameter("TrackFlag") = 1
    oPointAlongCurve44.Evaluate

    Dim oLineByPoints45 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints45 = oGCFactory.CreateEntity("LineByPoints", pPOM, "045-LineByPoints")
    oLineByPoints45.Inputs("StartPoint").Add oPointAlongCurve43
    oLineByPoints45.Inputs("EndPoint").Add oPointAlongCurve41
    oLineByPoints45.Evaluate
    
    Dim oFacePortExtractor46 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor46 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "046-FacePortExtractor")
    oFacePortExtractor46.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor46.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor46.Parameter("GeometrySelector") = 4
    oFacePortExtractor46.Parameter("FacesContext") = 2
    oFacePortExtractor46.Parameter("LookingAxis") = 2
    oFacePortExtractor46.Parameter("SurfaceType") = 1
    oFacePortExtractor46.Parameter("TrackFlag") = 2
    oFacePortExtractor46.Parameter("Offset") = 0#
    oFacePortExtractor46.Evaluate
    pGeometricConstruction.ControlledInputs(FACEPORT_WEB).Clear
    pGeometricConstruction.ControlledInputs(FACEPORT_WEB).AddElements oFacePortExtractor46.ControlledInputs("Port")

    Dim oLineByPoints47 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints47 = oGCFactory.CreateEntity("LineByPoints", pPOM, "047-LineByPoints")
    oLineByPoints47.Inputs("StartPoint").Add oPointAlongCurve42
    oLineByPoints47.Inputs("EndPoint").Add oPointAlongCurve44
    oLineByPoints47.Evaluate

    Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "048-PointAlongCurve")
    oPointAlongCurve48.Inputs("Curve").Add oLineByPoints47
    oPointAlongCurve48.Inputs("Point").Add oPointAlongCurve42
    oPointAlongCurve48.Inputs("TrackPoint").Add oPointAlongCurve44
    oPointAlongCurve48.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve48.Parameter("TrackFlag") = 2
    oPointAlongCurve48.Evaluate

    Dim oPointAlongCurve49 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
    oPointAlongCurve49.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve49.Inputs("Point").Add oPointByCurveAndSurface38
    oPointAlongCurve49.Inputs("TrackPoint").Add oPointAlongCurve41
    oPointAlongCurve49.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve49.Parameter("TrackFlag") = 1
    oPointAlongCurve49.Evaluate

    Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
    oPointAlongCurve50.Inputs("Curve").Add oLineByPoints45
    oPointAlongCurve50.Inputs("Point").Add oPointAlongCurve43
    oPointAlongCurve50.Inputs("TrackPoint").Add oPointAlongCurve41
    oPointAlongCurve50.Parameter("Distance") = oParamDistConstant8
    oPointAlongCurve50.Parameter("TrackFlag") = 1
    oPointAlongCurve50.Evaluate

    Dim oPointAlongCurve51 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve51 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "051-PointAlongCurve")
    oPointAlongCurve51.Inputs("Curve").Add oLineByPoints45
    oPointAlongCurve51.Inputs("Point").Add oPointAlongCurve41
    oPointAlongCurve51.Inputs("TrackPoint").Add oPointAlongCurve43
    oPointAlongCurve51.Parameter("Distance") = oParamDistFunctionWith2Args37
    oPointAlongCurve51.Parameter("TrackFlag") = 2
    oPointAlongCurve51.Evaluate

    Dim oLineByPoints52 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints52 = oGCFactory.CreateEntity("LineByPoints", pPOM, "052-LineByPoints")
    oLineByPoints52.Inputs("StartPoint").Add oPointAlongCurve43
    oLineByPoints52.Inputs("EndPoint").Add oPointAlongCurve51
    oLineByPoints52.Evaluate

    Dim oPointByCurveAndSurface53 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface53 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "053-PointByCurveAndSurface")
    oPointByCurveAndSurface53.Inputs("Curve1").Add oLineByPoints52
    oPointByCurveAndSurface53.Inputs("Surface2").Add oFacePortExtractor16
    oPointByCurveAndSurface53.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface53.Evaluate

    Dim oPointByCurveAndSurface54 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface54 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "054-PointByCurveAndSurface")
    oPointByCurveAndSurface54.Inputs("Curve1").Add oLineByPoints52
    oPointByCurveAndSurface54.Inputs("Surface2").Add oFacePortExtractor17
    oPointByCurveAndSurface54.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface54.Evaluate

    Dim oLineByPoints55 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints55 = oGCFactory.CreateEntity("LineByPoints", pPOM, "055-LineByPoints")
    oLineByPoints55.Inputs("StartPoint").Add oPointAlongCurve34
    oLineByPoints55.Inputs("EndPoint").Add oPointAlongCurve50
    oLineByPoints55.Evaluate

    Dim oLineByPoints56 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints56 = oGCFactory.CreateEntity("LineByPoints", pPOM, "056-LineByPoints")
    oLineByPoints56.Inputs("StartPoint").Add oPointByCurveAndSurface54
    oLineByPoints56.Inputs("EndPoint").Add oPointByCurveAndSurface53
    oLineByPoints56.Evaluate

    Dim oParamDistMeasureLength57 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength57 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "057-ParamDistMeasureLength")
    oParamDistMeasureLength57.Inputs("Curve").Add oLineByPoints56
    oParamDistMeasureLength57.Parameter("Value") = 0.015
    oParamDistMeasureLength57.Evaluate

    Dim oParamDistFunctionWith2Args58 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args58 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "058-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args58.Inputs("Arg1").Add oParamDistMeasureLength57
    oParamDistFunctionWith2Args58.Inputs("Arg2").Add oParamDouble2
    oParamDistFunctionWith2Args58.Parameter("Value") = 0.0075
    oParamDistFunctionWith2Args58.Parameter("Operation") = 3
    oParamDistFunctionWith2Args58.Evaluate

    Dim oPointAlongCurve59 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve59 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "059-PointAlongCurve")
    oPointAlongCurve59.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve59.Inputs("Point").Add oPointAtCurveMiddle33
    oPointAlongCurve59.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve59.Parameter("Distance") = oParamDistFunctionWith2Args58
    oPointAlongCurve59.Parameter("TrackFlag") = 1
    oPointAlongCurve59.Evaluate

    Dim oPointAlongCurve60 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve60 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "060-PointAlongCurve")
    oPointAlongCurve60.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve60.Inputs("Point").Add oPointAlongCurve59
    oPointAlongCurve60.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve60.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve60.Parameter("TrackFlag") = 1
    oPointAlongCurve60.Evaluate

    Dim oPointByCurveAndSurface61 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface61 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "061-PointByCurveAndSurface")
    oPointByCurveAndSurface61.Inputs("Curve1").Add oLineByPoints55
    oPointByCurveAndSurface61.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface61.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface61.Evaluate

    Dim oPointAlongCurve62 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve62 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "062-PointAlongCurve")
    oPointAlongCurve62.Inputs("Curve").Add oLineByPoints55
    oPointAlongCurve62.Inputs("Point").Add oPointAlongCurve34
    oPointAlongCurve62.Inputs("TrackPoint").Add oPointAlongCurve50
    oPointAlongCurve62.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve62.Parameter("TrackFlag") = 2
    oPointAlongCurve62.Evaluate

    Dim oPointAlongCurve63 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve63 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "063-PointAlongCurve")
    oPointAlongCurve63.Inputs("Curve").Add oLineByPoints40
    oPointAlongCurve63.Inputs("Point").Add oPointByCurveAndSurface38
    oPointAlongCurve63.Inputs("TrackPoint").Add oPointByCurveAndSurface39
    oPointAlongCurve63.Parameter("Distance") = oParamDistFunctionWith2Args58
    oPointAlongCurve63.Parameter("TrackFlag") = 1
    oPointAlongCurve63.Evaluate

    Dim oPointAlongCurve64 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve64 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "064-PointAlongCurve")
    oPointAlongCurve64.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve64.Inputs("Point").Add oPointAlongCurve49
    oPointAlongCurve64.Inputs("TrackPoint").Add oPointAlongCurve41
    oPointAlongCurve64.Parameter("Distance") = oParamDistFunctionWith2Args58
    oPointAlongCurve64.Parameter("TrackFlag") = 1
    oPointAlongCurve64.Evaluate

    Dim oPointAlongCurve65 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve65 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "065-PointAlongCurve")
    oPointAlongCurve65.Inputs("Curve").Add oLineByPoints40
    oPointAlongCurve65.Inputs("Point").Add oPointAlongCurve63
    oPointAlongCurve65.Inputs("TrackPoint").Add oPointByCurveAndSurface39
    oPointAlongCurve65.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve65.Parameter("TrackFlag") = 1
    oPointAlongCurve65.Evaluate

    Dim oLineByPoints66 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints66 = oGCFactory.CreateEntity("LineByPoints", pPOM, "066-LineByPoints")
    oLineByPoints66.Inputs("StartPoint").Add oPointAlongCurve62
    oLineByPoints66.Inputs("EndPoint").Add oPointAlongCurve50
    oLineByPoints66.Evaluate

    Dim oPointAlongCurve67 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve67 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "067-PointAlongCurve")
    oPointAlongCurve67.Inputs("Curve").Add oLineByPoints47
    oPointAlongCurve67.Inputs("Point").Add oPointAlongCurve42
    oPointAlongCurve67.Inputs("TrackPoint").Add oPointAlongCurve44
    oPointAlongCurve67.Parameter("Distance") = oParamDistFunctionWith2Args58
    oPointAlongCurve67.Parameter("TrackFlag") = 1
    oPointAlongCurve67.Evaluate

    Dim oPointAlongCurve68 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve68 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "068-PointAlongCurve")
    oPointAlongCurve68.Inputs("Curve").Add oLineByPoints47
    oPointAlongCurve68.Inputs("Point").Add oPointAlongCurve67
    oPointAlongCurve68.Inputs("TrackPoint").Add oPointAlongCurve44
    oPointAlongCurve68.Parameter("Distance") = oParamDistConstant9
    oPointAlongCurve68.Parameter("TrackFlag") = 1
    oPointAlongCurve68.Evaluate

    Dim oLineByPoints69 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints69 = oGCFactory.CreateEntity("LineByPoints", pPOM, "069-LineByPoints")
    oLineByPoints69.Inputs("StartPoint").Add oPointAlongCurve68
    oLineByPoints69.Inputs("EndPoint").Add oPointAlongCurve48
    oLineByPoints69.Evaluate

    Dim oPlaneByPointNormal70 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal70 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "070-PlaneByPointNormal")
    oPlaneByPointNormal70.Inputs("Point").Add oPointAlongCurve59
    oPlaneByPointNormal70.Inputs("Line").Add oLineAxisPortExtractor1
    oPlaneByPointNormal70.Parameter("Range") = 0.999998
    oPlaneByPointNormal70.Evaluate

    Dim oPointAlongCurve71 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve71 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "071-PointAlongCurve")
    oPointAlongCurve71.Inputs("Curve").Add oLineByPoints55
    oPointAlongCurve71.Inputs("Point").Add oPointByCurveAndSurface61
    oPointAlongCurve71.Inputs("TrackPoint").Add oPointAlongCurve50
    oPointAlongCurve71.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve71.Parameter("TrackFlag") = 1
    oPointAlongCurve71.Evaluate

    Dim oLineByPoints72 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints72 = oGCFactory.CreateEntity("LineByPoints", pPOM, "072-LineByPoints")
    oLineByPoints72.Inputs("StartPoint").Add oPointByCurveAndSurface61
    oLineByPoints72.Inputs("EndPoint").Add oPointAlongCurve34
    oLineByPoints72.Evaluate

    Dim oLineByPoints73 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints73 = oGCFactory.CreateEntity("LineByPoints", pPOM, "073-LineByPoints")
    oLineByPoints73.Inputs("StartPoint").Add oPointAlongCurve68
    oLineByPoints73.Inputs("EndPoint").Add oPointAlongCurve50
    oLineByPoints73.Evaluate

    Dim oParamDistMeasureLength74 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength74 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "074-ParamDistMeasureLength")
    oParamDistMeasureLength74.Inputs("Curve").Add oLineByPoints72
    oParamDistMeasureLength74.Parameter("Value") = 0.025
    oParamDistMeasureLength74.Evaluate

    Dim oLineByPoints75 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints75 = oGCFactory.CreateEntity("LineByPoints", pPOM, "075-LineByPoints")
    oLineByPoints75.Inputs("StartPoint").Add oPointAlongCurve63
    oLineByPoints75.Inputs("EndPoint").Add oPointAlongCurve67
    oLineByPoints75.Evaluate

    Dim oPointAlongCurve76 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve76 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "076-PointAlongCurve")
    oPointAlongCurve76.Inputs("Curve").Add oLineByPoints75
    oPointAlongCurve76.Inputs("Point").Add oPointAlongCurve63
    oPointAlongCurve76.Inputs("TrackPoint").Add oPointByCurveAndSurface53
    oPointAlongCurve76.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve76.Parameter("TrackFlag") = 1
    oPointAlongCurve76.Evaluate

    Dim oLineByPoints77 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints77 = oGCFactory.CreateEntity("LineByPoints", pPOM, "077-LineByPoints")
    oLineByPoints77.Inputs("StartPoint").Add oPointAlongCurve76
    oLineByPoints77.Inputs("EndPoint").Add oPointAlongCurve64
    oLineByPoints77.Evaluate

    Dim oPointAlongCurve78 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve78 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "078-PointAlongCurve")
    oPointAlongCurve78.Inputs("Curve").Add oLineByPoints77
    oPointAlongCurve78.Inputs("Point").Add oPointAlongCurve64
    oPointAlongCurve78.Inputs("TrackPoint").Add oPointAlongCurve76
    oPointAlongCurve78.Parameter("Distance") = oParamDistFunctionWith2Args58
    oPointAlongCurve78.Parameter("TrackFlag") = 2
    oPointAlongCurve78.Evaluate

    Dim oLineByPoints79 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints79 = oGCFactory.CreateEntity("LineByPoints", pPOM, "079-LineByPoints")
    oLineByPoints79.Inputs("StartPoint").Add oPointAlongCurve76
    oLineByPoints79.Inputs("EndPoint").Add oPointAlongCurve78
    oLineByPoints79.Evaluate

    Dim oCpxStringByCurves80 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves80 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "080-CpxStringByCurves")
    oCpxStringByCurves80.Inputs("Curves").Add oLineByPoints66
    oCpxStringByCurves80.Inputs("Curves").Add oLineByPoints73
    oCpxStringByCurves80.Inputs("Curves").Add oLineByPoints69
    oCpxStringByCurves80.Evaluate

    Dim oSurfFromGType81 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType81 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "081-SurfFromGType")
    oSurfFromGType81.Inputs("Surface").Add oPlaneByPointNormal70
    oSurfFromGType81.Evaluate

    Dim oSurfByLinearExtrusion82 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion82 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "082-SurfByLinearExtrusion")
    oSurfByLinearExtrusion82.Inputs("PlanarCrossSection").Add oCpxStringByCurves80
    oSurfByLinearExtrusion82.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion82.Evaluate

    Dim oPointAlongCurve83 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve83 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "083-PointAlongCurve")
    oPointAlongCurve83.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve83.Inputs("Point").Add oPointAlongCurve60
    oPointAlongCurve83.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve83.Parameter("Distance") = oParamDistMeasureLength74
    oPointAlongCurve83.Parameter("TrackFlag") = 1
    oPointAlongCurve83.Evaluate

    Dim oLineByPoints84 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints84 = oGCFactory.CreateEntity("LineByPoints", pPOM, "084-LineByPoints")
    oLineByPoints84.Inputs("StartPoint").Add oPointAlongCurve83
    oLineByPoints84.Inputs("EndPoint").Add oPointAlongCurve65
    oLineByPoints84.Evaluate

    Dim oPointAlongCurve85 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve85 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "085-PointAlongCurve")
    oPointAlongCurve85.Inputs("Curve").Add oLineByPoints84
    oPointAlongCurve85.Inputs("Point").Add oPointAlongCurve83
    oPointAlongCurve85.Inputs("TrackPoint").Add oPointAlongCurve65
    oPointAlongCurve85.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve85.Parameter("TrackFlag") = 2
    oPointAlongCurve85.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    
    Dim oLineByPoints88 As SP3DGeometricConstruction.GeometricConstruction
    Dim oCpxStringByCurves90 As SP3DGeometricConstruction.GeometricConstruction
    Dim oSurfByLinearExtrusion91 As SP3DGeometricConstruction.GeometricConstruction

    If m_bCreateArcOnAPS = True Then
        Dim oLineByPoints86 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints86 = oGCFactory.CreateEntity("LineByPoints", pPOM, "086-LineByPoints")
        oLineByPoints86.Inputs("StartPoint").Add oPointAlongCurve76
        oLineByPoints86.Inputs("EndPoint").Add oPointAlongCurve71
        oLineByPoints86.Evaluate
        
        Dim oPointAlongCurve87 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve87 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "087-PointAlongCurve")
        oPointAlongCurve87.Inputs("Curve").Add oLineByPoints86
        oPointAlongCurve87.Inputs("Point").Add oPointAlongCurve76
        oPointAlongCurve87.Inputs("TrackPoint").Add oPointAlongCurve71
        oPointAlongCurve87.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve87.Parameter("TrackFlag") = 1
        oPointAlongCurve87.Evaluate
    
        Set oLineByPoints88 = oGCFactory.CreateEntity("LineByPoints", pPOM, "088-LineByPoints")
        oLineByPoints88.Inputs("StartPoint").Add oPointAlongCurve85
        oLineByPoints88.Inputs("EndPoint").Add oPointAlongCurve65
        oLineByPoints88.Evaluate
        
        Dim oArcByCenter89 As SP3DGeometricConstruction.GeometricConstruction
        Set oArcByCenter89 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "089-ArcByCenter")
        oArcByCenter89.Inputs("Center").Add oPointAlongCurve87
        oArcByCenter89.Inputs("StartPoint").Add oPointAlongCurve65
        oArcByCenter89.Inputs("EndPoint").Add oPointAlongCurve76
        oArcByCenter89.Inputs("NormalLine").Add oLineAxisPortExtractor1
        oArcByCenter89.Inputs("TrackPoint").Add oPointAlongCurve63
        oArcByCenter89.Parameter("SweepAngle") = 1
        oArcByCenter89.Parameter("TrackFlag") = 1
        oArcByCenter89.Evaluate
        
        Set oCpxStringByCurves90 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "090-CpxStringByCurves")
        oCpxStringByCurves90.Inputs("Curves").Add oLineByPoints88
        oCpxStringByCurves90.Inputs("Curves").Add oArcByCenter89
        oCpxStringByCurves90.Inputs("Curves").Add oLineByPoints79
        oCpxStringByCurves90.Evaluate
    
        Set oSurfByLinearExtrusion91 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "091-SurfByLinearExtrusion")
        oSurfByLinearExtrusion91.Inputs("PlanarCrossSection").Add oCpxStringByCurves90
        oSurfByLinearExtrusion91.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
        oSurfByLinearExtrusion91.Evaluate
    
        oGCMacro.Outputs("Boundary").Clear
        oGCMacro.Output("Support", 1) = oSurfFromGType81.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion91.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion82.Output
        oGCMacro.Output("Boundary", FACEPORT_FLANGE) = oFacePortExtractor21.Output
        oGCMacro.Output("Boundary", FACEPORT_WEB) = oFacePortExtractor46.Output
    Else
        oGCMacro.Outputs("Boundary").Clear
        oGCMacro.Output("Support", 1) = oSurfFromGType81.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion82.Output
        oGCMacro.Output("Boundary", FACEPORT_FLANGE) = oFacePortExtractor21.Output
        oGCMacro.Output("Boundary", FACEPORT_WEB) = oFacePortExtractor46.Output
    End If
End Sub

Private Sub Evaluate_OneSidePlateOnStdMember(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    If pGeometricConstruction.Inputs("MemberPart").Count <> 1 Then Err.Raise MISSING_MEMBER
    Dim oSPSMemberPartPrismatic As ISPSMemberPartPrismatic
    On Error Resume Next
    Set oSPSMemberPartPrismatic = pGeometricConstruction.Inputs("MemberPart").Item(1)
    On Error GoTo 0
    If (oSPSMemberPartPrismatic Is Nothing) Then Err.Raise INVALID_MEMBER
    
    'Check any parameter is -ve
    If (pGeometricConstruction.Parameter("Offset") < 0) Then Err.Raise INVALID_OFFSET
    If (pGeometricConstruction.Parameter("WeldToe1") < 0) Then Err.Raise INVALID_WELDTOE1
    If (pGeometricConstruction.Parameter("WeldToe2") < 0) Then Err.Raise INVALID_WELDTOE2
    If (pGeometricConstruction.Parameter("Height1") < 0) Then Err.Raise INVALID_HEIGHT1
    If (pGeometricConstruction.Parameter("Height2") < 0) Then Err.Raise INVALID_HEIGHT2

    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs(PORT).Clear
    pGeometricConstruction.ControlledInputs(PORT).AddElements oLineAxisPortExtractor1.ControlledInputs("Port")
    
        Dim oLine As IJLine
    On Error Resume Next
    Set oLine = oLineAxisPortExtractor1.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Offset") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Offset") > oLine.Length Then Err.Raise INVALID_OFFSET
    End If

    If TypeOf pGeometricConstruction.Inputs("MemberPart").Item(1) Is ISPSMemberPartCommon Then
        Dim oSPSCommon As ISPSMemberPartCommon
        Dim OSPSxSec As ISPSCrossSection
        Dim oSectionAttrbs As IJDAttributes
        Dim dKDetail As Double
        Dim dFlangeThickness As Double
        Dim dStdMemberRadius As Double
        
        Set oSPSCommon = pGeometricConstruction.Inputs("MemberPart").Item(1)
        Set OSPSxSec = oSPSCommon.CrossSection
        Set oSectionAttrbs = OSPSxSec.Definition
         'Calculate Member Radius
        dKDetail = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("kdetail").Value
        dFlangeThickness = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
        dStdMemberRadius = dKDetail - dFlangeThickness
    End If
        
    Dim oParamDouble2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble2 = oGCFactory.CreateEntity("ParamDouble", pPOM, "002-ParamDouble")
    oParamDouble2.Parameter("Value") = 0.5
    oParamDouble2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Height1")
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Height2")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = 0.015
    oParamDistConstant5.Evaluate

    Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "006-ParamDistConstant")
    oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("Offset")
    oParamDistConstant6.Evaluate

    Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "007-ParamDistConstant")
    Dim dGCRadius As Double
    dGCRadius = pGeometricConstruction.Parameter("Radius")
    'set the GCRadius to member radius
    If dGCRadius > 0.0001 And Abs(dGCRadius - dStdMemberRadius) > 0.0001 Then
        dGCRadius = dStdMemberRadius
        pGeometricConstruction.Parameter("Radius") = dGCRadius
    End If
    oParamDistConstant7.Parameter("Value") = dGCRadius
    oParamDistConstant7.Evaluate

    Dim oParamDistConstant8 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant8 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "008-ParamDistConstant")
    oParamDistConstant8.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe1")
    oParamDistConstant8.Evaluate

    Dim oParamDistConstant9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant9 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "009-ParamDistConstant")
    oParamDistConstant9.Parameter("Value") = pGeometricConstruction.Parameter("WeldToe2")
    oParamDistConstant9.Evaluate

    Dim oPointAtCurveStart10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart10 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "010-PointAtCurveStart")
    oPointAtCurveStart10.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveStart10.Evaluate

    Dim oPointAtCurveEnd11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd11 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "011-PointAtCurveEnd")
    oPointAtCurveEnd11.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveEnd11.Evaluate

    Dim oLineFromCrossSection12 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection12 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "012-LineFromCrossSection")
    oLineFromCrossSection12.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection12.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection12.Parameter("EdgeSelector") = 513
    oLineFromCrossSection12.Evaluate

    'Validate the weldToe2
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection12.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe2") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe2") > oLine.Length Then Err.Raise INVALID_WELDTOE2
    End If


    Dim oCSFromMember13 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember13 = oGCFactory.CreateEntity("CSFromMember", pPOM, "013-CSFromMember")
    oCSFromMember13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember13.Inputs("Point").Add oPointAtCurveStart10
    oCSFromMember13.Evaluate

    Dim oLineFromCrossSection14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection14 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "014-LineFromCrossSection")
    oLineFromCrossSection14.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection14.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection14.Parameter("EdgeSelector") = 258
    oLineFromCrossSection14.Evaluate

    'Validate the height
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection14.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Height1") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Height1") > oLine.Length Then Err.Raise INVALID_HEIGHT1
    End If

    'Validate the height
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection14.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("Height2") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("Height2") > oLine.Length Then Err.Raise INVALID_HEIGHT2
    End If

    Dim oLineFromCrossSection15 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCrossSection15 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "015-LineFromCrossSection")
    oLineFromCrossSection15.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineFromCrossSection15.Inputs("Point").Add oPointAtCurveStart10
    oLineFromCrossSection15.Parameter("EdgeSelector") = 514
    oLineFromCrossSection15.Evaluate

    'Validate the weldToe1
    Set oLine = Nothing
    On Error Resume Next
    Set oLine = oLineFromCrossSection15.Output
    On Error GoTo 0
    If Not oLine Is Nothing Then
        If Abs(pGeometricConstruction.Parameter("WeldToe1") - oLine.Length) > TOLERANCE And pGeometricConstruction.Parameter("WeldToe1") > oLine.Length Then Err.Raise INVALID_WELDTOE1
    End If

    Dim oFacePortExtractor16 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor16 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "016-FacePortExtractor")
    oFacePortExtractor16.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor16.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor16.Parameter("GeometrySelector") = 4
    oFacePortExtractor16.Parameter("FacesContext") = 5
    oFacePortExtractor16.Parameter("LookingAxis") = 2
    oFacePortExtractor16.Parameter("SurfaceType") = 1
    oFacePortExtractor16.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFacePortExtractor16.Parameter("Offset") = 0#
    oFacePortExtractor16.Evaluate
    pGeometricConstruction.ControlledInputs(FACEPORT_WEB).Clear
    pGeometricConstruction.ControlledInputs(FACEPORT_WEB).AddElements oFacePortExtractor16.ControlledInputs("Port")

    Dim oFacePortExtractor17 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor17 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "017-FacePortExtractor")
    oFacePortExtractor17.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor17.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor17.Parameter("GeometrySelector") = 4
    oFacePortExtractor17.Parameter("FacesContext") = 5
    oFacePortExtractor17.Parameter("LookingAxis") = 2
    oFacePortExtractor17.Parameter("SurfaceType") = 1
    oFacePortExtractor17.Parameter("TrackFlag") = 2 / pGeometricConstruction.Parameter("TrackFlag")
    oFacePortExtractor17.Parameter("Offset") = 0#
    oFacePortExtractor17.Evaluate

    Dim oCSByCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS18 = oGCFactory.CreateEntity("CSByCS", pPOM, "018-CSByCS")
    oCSByCS18.Inputs("CoordinateSystem").Add oCSFromMember13
    oCSByCS18.Inputs("Origin").Add oPointAtCurveStart10
    oCSByCS18.Parameter("AxesRoles") = 1
    oCSByCS18.Parameter("CSOrientation") = pGeometricConstruction.Parameter("CSOrientation")
    oCSByCS18.Parameter("TrackFlag") = 1
    oCSByCS18.Evaluate

    Dim oFaceFromCrossSection19 As SP3DGeometricConstruction.GeometricConstruction
    Set oFaceFromCrossSection19 = oGCFactory.CreateEntity("FaceFromCrossSection", pPOM, "019-FaceFromCrossSection")
    oFaceFromCrossSection19.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFaceFromCrossSection19.Inputs("CoordinateSystem").Add oCSFromMember13
    oFaceFromCrossSection19.Parameter("LookingAxis") = 2
    oFaceFromCrossSection19.Parameter("TrackFlag") = pGeometricConstruction.Parameter("TrackFlag")
    oFaceFromCrossSection19.Parameter("Offset") = 0#
    oFaceFromCrossSection19.Evaluate

    Dim oFacePortExtractor20 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor20 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "020-FacePortExtractor")
    oFacePortExtractor20.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor20.Inputs("CoordinateSystem").Add oCSByCS18
    oFacePortExtractor20.Parameter("GeometrySelector") = 4
    oFacePortExtractor20.Parameter("FacesContext") = 2
    oFacePortExtractor20.Parameter("LookingAxis") = 3
    oFacePortExtractor20.Parameter("SurfaceType") = 1
    oFacePortExtractor20.Parameter("TrackFlag") = 1
    oFacePortExtractor20.Parameter("Offset") = 0#
    oFacePortExtractor20.Evaluate

    Dim oFacePortExtractor21 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor21 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "021-FacePortExtractor")
    oFacePortExtractor21.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor21.Inputs("CoordinateSystem").Add oCSByCS18
    oFacePortExtractor21.Parameter("GeometrySelector") = 4
    oFacePortExtractor21.Parameter("FacesContext") = 2
    oFacePortExtractor21.Parameter("LookingAxis") = 3
    oFacePortExtractor21.Parameter("SurfaceType") = 1
    oFacePortExtractor21.Parameter("TrackFlag") = 2
    oFacePortExtractor21.Parameter("Offset") = 0#
    oFacePortExtractor21.Evaluate
    
    'StandardMember
    Dim oFacePortExtractor22 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor22 = oGCFactory.CreateEntity("FacePortExtractor1", pPOM, "022-FacePortExtractor1")
    oFacePortExtractor22.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor22.Inputs("CoordinateSystem").Add oCSByCS18
    oFacePortExtractor22.Inputs("Face").Add oFacePortExtractor16
    oFacePortExtractor22.Parameter("GeometrySelector") = 4
    oFacePortExtractor22.Parameter("FacesContext") = 4
    oFacePortExtractor22.Parameter("LookingAxis") = 3
    oFacePortExtractor22.Parameter("SurfaceType") = 1
    oFacePortExtractor22.Parameter("TrackFlag") = 2
    oFacePortExtractor22.Parameter("Offset") = 0#
    oFacePortExtractor22.Evaluate
    pGeometricConstruction.ControlledInputs(FACEPORT_FLANGE).Clear
    pGeometricConstruction.ControlledInputs(FACEPORT_FLANGE).AddElements oFacePortExtractor22.ControlledInputs("Port")

    Dim oPointByCurveAndSurface23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface23 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "023-PointByCurveAndSurface")
    oPointByCurveAndSurface23.Inputs("Curve1").Add oLineFromCrossSection14
    oPointByCurveAndSurface23.Inputs("Surface2").Add oFacePortExtractor20
    oPointByCurveAndSurface23.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface23.Evaluate

    Dim oPointByCurveAndSurface24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface24 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "024-PointByCurveAndSurface")
    oPointByCurveAndSurface24.Inputs("Curve1").Add oLineFromCrossSection14
    oPointByCurveAndSurface24.Inputs("Surface2").Add oFacePortExtractor21
    oPointByCurveAndSurface24.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface24.Evaluate

    Dim oLineAtAngle25 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle25 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "025-LineAtAngle")
    oLineAtAngle25.Inputs("Line").Add oLineFromCrossSection15
    oLineAtAngle25.Inputs("Point").Add oPointByCurveAndSurface23
    oLineAtAngle25.Inputs("Plane").Add oFacePortExtractor20
    oLineAtAngle25.Parameter("Angle") = 0#
    oLineAtAngle25.Parameter("TrackFlag") = 1
    oLineAtAngle25.Evaluate

    Dim oLineAtAngle26 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle26 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "026-LineAtAngle")
    oLineAtAngle26.Inputs("Line").Add oLineAtAngle25
    oLineAtAngle26.Inputs("Point").Add oPointByCurveAndSurface24
    oLineAtAngle26.Inputs("Plane").Add oFacePortExtractor21
    oLineAtAngle26.Parameter("Angle") = 0#
    oLineAtAngle26.Parameter("TrackFlag") = 1
    oLineAtAngle26.Evaluate

    Dim oLineParallel27 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel27 = oGCFactory.CreateEntity("LineParallel", pPOM, "027-LineParallel")
    oLineParallel27.Inputs("Line").Add oLineAtAngle26
    oLineParallel27.Inputs("Plane").Add oFacePortExtractor21
    oLineParallel27.Inputs("TrackPoint").Add oPointAtCurveEnd11
    oLineParallel27.Parameter("Distance") = oParamDistConstant6
    oLineParallel27.Parameter("TrackFlag") = 1
    oLineParallel27.Evaluate

    Dim oLineParallel28 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel28 = oGCFactory.CreateEntity("LineParallel", pPOM, "028-LineParallel")
    oLineParallel28.Inputs("Line").Add oLineAtAngle25
    oLineParallel28.Inputs("Plane").Add oFacePortExtractor20
    oLineParallel28.Inputs("TrackPoint").Add oPointAtCurveEnd11
    oLineParallel28.Parameter("Distance") = oParamDistConstant6
    oLineParallel28.Parameter("TrackFlag") = 1
    oLineParallel28.Evaluate

    Dim oPointByCurveAndSurface29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface29 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "029-PointByCurveAndSurface")
    oPointByCurveAndSurface29.Inputs("Curve1").Add oLineParallel28
    oPointByCurveAndSurface29.Inputs("Surface2").Add oFaceFromCrossSection19
    oPointByCurveAndSurface29.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface29.Evaluate

    Dim oPointAtCurveMiddle30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle30 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "030-PointAtCurveMiddle")
    oPointAtCurveMiddle30.Inputs("Curve").Add oLineParallel28
    oPointAtCurveMiddle30.Evaluate

    Dim oParamDistMeasureLength31 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength31 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "031-ParamDistMeasureLength")
    oParamDistMeasureLength31.Inputs("Curve").Add oLineParallel27
    oParamDistMeasureLength31.Parameter("Value") = 0.2
    oParamDistMeasureLength31.Evaluate

    Dim oPointByCurveAndSurface32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface32 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "032-PointByCurveAndSurface")
    oPointByCurveAndSurface32.Inputs("Curve1").Add oLineParallel27
    oPointByCurveAndSurface32.Inputs("Surface2").Add oFaceFromCrossSection19
    oPointByCurveAndSurface32.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface32.Evaluate

    Dim oPointAtCurveMiddle33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle33 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "033-PointAtCurveMiddle")
    oPointAtCurveMiddle33.Inputs("Curve").Add oLineParallel27
    oPointAtCurveMiddle33.Evaluate

    Dim oPointAlongCurve34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve34 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "034-PointAlongCurve")
    oPointAlongCurve34.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve34.Inputs("Point").Add oPointByCurveAndSurface32
    oPointAlongCurve34.Inputs("TrackPoint").Add oPointAtCurveMiddle33
    oPointAlongCurve34.Parameter("Distance") = oParamDistConstant8
    oPointAlongCurve34.Parameter("TrackFlag") = 1
    oPointAlongCurve34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointAtCurveMiddle33
    oLineByPoints35.Inputs("EndPoint").Add oPointAtCurveMiddle30
    oLineByPoints35.Evaluate

    Dim oLineByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints36 = oGCFactory.CreateEntity("LineByPoints", pPOM, "036-LineByPoints")
    oLineByPoints36.Inputs("StartPoint").Add oPointByCurveAndSurface32
    oLineByPoints36.Inputs("EndPoint").Add oPointByCurveAndSurface29
    oLineByPoints36.Evaluate

    Dim oParamDistFunctionWith2Args37 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args37 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "037-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args37.Inputs("Arg1").Add oParamDistMeasureLength31
    oParamDistFunctionWith2Args37.Inputs("Arg2").Add oParamDouble2
    oParamDistFunctionWith2Args37.Parameter("Value") = 0.1
    oParamDistFunctionWith2Args37.Parameter("Operation") = 3
    oParamDistFunctionWith2Args37.Evaluate
    
    Dim oPointByCurveAndSurface38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface38 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "038-PointByCurveAndSurface")
    oPointByCurveAndSurface38.Inputs("Curve1").Add oLineByPoints36
    oPointByCurveAndSurface38.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface38.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface38.Evaluate
    
    Dim oLineByPoints39 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints39 = oGCFactory.CreateEntity("LineByPoints", pPOM, "039-LineByPoints")
    oLineByPoints39.Inputs("StartPoint").Add oPointByCurveAndSurface32
    oLineByPoints39.Inputs("EndPoint").Add oPointByCurveAndSurface38
    oLineByPoints39.Evaluate

    Dim oParamDistMeasureLength40 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength40 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "040-ParamDistMeasureLength")
    oParamDistMeasureLength40.Inputs("Curve").Add oLineByPoints39
    oParamDistMeasureLength40.Parameter("Value") = 0.015
    oParamDistMeasureLength40.Evaluate
    
    Dim oPointAlongCurve41 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve41 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "041-PointAlongCurve")
    oPointAlongCurve41.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve41.Inputs("Point").Add oPointAtCurveMiddle33
    oPointAlongCurve41.Inputs("TrackPoint").Add oPointAtCurveMiddle30
    oPointAlongCurve41.Parameter("Distance") = oParamDistMeasureLength40
    oPointAlongCurve41.Parameter("TrackFlag") = 1
    oPointAlongCurve41.Evaluate

    Dim oPointByCurveAndSurface42 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface42 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "042-PointByCurveAndSurface")
    oPointByCurveAndSurface42.Inputs("Curve1").Add oLineByPoints36
    oPointByCurveAndSurface42.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface42.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface42.Evaluate

    Dim oLineByPoints43 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints43 = oGCFactory.CreateEntity("LineByPoints", pPOM, "043-LineByPoints")
    oLineByPoints43.Inputs("StartPoint").Add oPointByCurveAndSurface42
    oLineByPoints43.Inputs("EndPoint").Add oPointAlongCurve41
    oLineByPoints43.Evaluate

    Dim oParamDistMeasureLength44 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength44 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "044-ParamDistMeasureLength")
    oParamDistMeasureLength44.Inputs("Curve").Add oLineByPoints43
    oParamDistMeasureLength44.Parameter("Value") = 0.015
    oParamDistMeasureLength44.Evaluate
    
    Dim oPointAlongCurve45 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve45 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "045-PointAlongCurve")
    oPointAlongCurve45.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve45.Inputs("Point").Add oPointAlongCurve41
    oPointAlongCurve45.Inputs("TrackPoint").Add oPointAtCurveMiddle30
    oPointAlongCurve45.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve45.Parameter("TrackFlag") = 1
    oPointAlongCurve45.Evaluate

    Dim oPointAlongCurve46 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve46 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "046-PointAlongCurve")
    oPointAlongCurve46.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve46.Inputs("Point").Add oPointAlongCurve45
    oPointAlongCurve46.Inputs("TrackPoint").Add oPointAtCurveMiddle30
    oPointAlongCurve46.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve46.Parameter("TrackFlag") = 1
    oPointAlongCurve46.Evaluate

    Dim oPointAlongCurve47 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve47 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "047-PointAlongCurve")
    oPointAlongCurve47.Inputs("Curve").Add oLineByPoints36
    oPointAlongCurve47.Inputs("Point").Add oPointByCurveAndSurface42
    oPointAlongCurve47.Inputs("TrackPoint").Add oPointByCurveAndSurface29
    oPointAlongCurve47.Parameter("Distance") = oParamDistConstant3
    oPointAlongCurve47.Parameter("TrackFlag") = 1
    oPointAlongCurve47.Evaluate

    Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "048-PointAlongCurve")
    oPointAlongCurve48.Inputs("Curve").Add oLineByPoints36
    oPointAlongCurve48.Inputs("Point").Add oPointAlongCurve47
    oPointAlongCurve48.Inputs("TrackPoint").Add oPointByCurveAndSurface29
    oPointAlongCurve48.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve48.Parameter("TrackFlag") = 1
    oPointAlongCurve48.Evaluate

    Dim oLineByPoints49 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints49 = oGCFactory.CreateEntity("LineByPoints", pPOM, "049-LineByPoints")
    oLineByPoints49.Inputs("StartPoint").Add oPointAlongCurve47
    oLineByPoints49.Inputs("EndPoint").Add oPointAlongCurve45
    oLineByPoints49.Evaluate

    Dim oFacePortExtractor50 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor50 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "050-FacePortExtractor")
    oFacePortExtractor50.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor50.Inputs("CoordinateSystem").Add oCSFromMember13
    oFacePortExtractor50.Parameter("GeometrySelector") = 4
    oFacePortExtractor50.Parameter("FacesContext") = 2
    oFacePortExtractor50.Parameter("LookingAxis") = 2
    oFacePortExtractor50.Parameter("SurfaceType") = 1
    oFacePortExtractor50.Parameter("TrackFlag") = 2
    oFacePortExtractor50.Parameter("Offset") = 0#
    oFacePortExtractor50.Evaluate

    Dim oLineByPoints51 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints51 = oGCFactory.CreateEntity("LineByPoints", pPOM, "051-LineByPoints")
    oLineByPoints51.Inputs("StartPoint").Add oPointAlongCurve46
    oLineByPoints51.Inputs("EndPoint").Add oPointAlongCurve48
    oLineByPoints51.Evaluate

    Dim oPointAlongCurve52 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve52 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "052-PointAlongCurve")
    oPointAlongCurve52.Inputs("Curve").Add oLineByPoints51
    oPointAlongCurve52.Inputs("Point").Add oPointAlongCurve46
    oPointAlongCurve52.Inputs("TrackPoint").Add oPointAlongCurve48
    oPointAlongCurve52.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve52.Parameter("TrackFlag") = 2
    oPointAlongCurve52.Evaluate

    Dim oPointAlongCurve53 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "053-PointAlongCurve")
    oPointAlongCurve53.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve53.Inputs("Point").Add oPointAlongCurve41
    oPointAlongCurve53.Inputs("TrackPoint").Add oPointAlongCurve45
    oPointAlongCurve53.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve53.Parameter("TrackFlag") = 1
    oPointAlongCurve53.Evaluate

    Dim oPointAlongCurve54 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve54 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "054-PointAlongCurve")
    oPointAlongCurve54.Inputs("Curve").Add oLineByPoints49
    oPointAlongCurve54.Inputs("Point").Add oPointAlongCurve47
    oPointAlongCurve54.Inputs("TrackPoint").Add oPointAlongCurve45
    oPointAlongCurve54.Parameter("Distance") = oParamDistConstant8
    oPointAlongCurve54.Parameter("TrackFlag") = 1
    oPointAlongCurve54.Evaluate

    Dim oPointAlongCurve55 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve55 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "055-PointAlongCurve")
    oPointAlongCurve55.Inputs("Curve").Add oLineByPoints49
    oPointAlongCurve55.Inputs("Point").Add oPointAlongCurve45
    oPointAlongCurve55.Inputs("TrackPoint").Add oPointAlongCurve47
    oPointAlongCurve55.Parameter("Distance") = oParamDistFunctionWith2Args37
    oPointAlongCurve55.Parameter("TrackFlag") = 2
    oPointAlongCurve55.Evaluate

    Dim oLineByPoints56 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints56 = oGCFactory.CreateEntity("LineByPoints", pPOM, "056-LineByPoints")
    oLineByPoints56.Inputs("StartPoint").Add oPointAlongCurve47
    oLineByPoints56.Inputs("EndPoint").Add oPointAlongCurve55
    oLineByPoints56.Evaluate

    Dim oPointByCurveAndSurface57 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface57 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "057-PointByCurveAndSurface")
    oPointByCurveAndSurface57.Inputs("Curve1").Add oLineByPoints56
    oPointByCurveAndSurface57.Inputs("Surface2").Add oFacePortExtractor16
    oPointByCurveAndSurface57.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface57.Evaluate

    Dim oPointByCurveAndSurface58 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface58 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "058-PointByCurveAndSurface")
    oPointByCurveAndSurface58.Inputs("Curve1").Add oLineByPoints56
    oPointByCurveAndSurface58.Inputs("Surface2").Add oFacePortExtractor17
    oPointByCurveAndSurface58.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface58.Evaluate

    Dim oLineByPoints59 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints59 = oGCFactory.CreateEntity("LineByPoints", pPOM, "059-LineByPoints")
    oLineByPoints59.Inputs("StartPoint").Add oPointAlongCurve34
    oLineByPoints59.Inputs("EndPoint").Add oPointAlongCurve54
    oLineByPoints59.Evaluate

    Dim oLineByPoints60 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints60 = oGCFactory.CreateEntity("LineByPoints", pPOM, "060-LineByPoints")
    oLineByPoints60.Inputs("StartPoint").Add oPointByCurveAndSurface58
    oLineByPoints60.Inputs("EndPoint").Add oPointByCurveAndSurface57
    oLineByPoints60.Evaluate

    Dim oParamDistMeasureLength61 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength61 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "061-ParamDistMeasureLength")
    oParamDistMeasureLength61.Inputs("Curve").Add oLineByPoints60
    oParamDistMeasureLength61.Parameter("Value") = 0.015
    oParamDistMeasureLength61.Evaluate

    Dim oParamDistFunctionWith2Args62 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args62 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "062-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args62.Inputs("Arg1").Add oParamDistMeasureLength61
    oParamDistFunctionWith2Args62.Inputs("Arg2").Add oParamDouble2
    oParamDistFunctionWith2Args62.Parameter("Value") = 0.0075
    oParamDistFunctionWith2Args62.Parameter("Operation") = 3
    oParamDistFunctionWith2Args62.Evaluate

    Dim oPointAlongCurve63 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve63 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "063-PointAlongCurve")
    oPointAlongCurve63.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve63.Inputs("Point").Add oPointAtCurveMiddle33
    oPointAlongCurve63.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve63.Parameter("Distance") = oParamDistFunctionWith2Args62
    oPointAlongCurve63.Parameter("TrackFlag") = 1
    oPointAlongCurve63.Evaluate

    Dim oPointAlongCurve64 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve64 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "064-PointAlongCurve")
    oPointAlongCurve64.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve64.Inputs("Point").Add oPointAlongCurve63
    oPointAlongCurve64.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve64.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve64.Parameter("TrackFlag") = 1
    oPointAlongCurve64.Evaluate

    Dim oPointByCurveAndSurface65 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface65 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "065-PointByCurveAndSurface")
    oPointByCurveAndSurface65.Inputs("Curve1").Add oLineByPoints59
    oPointByCurveAndSurface65.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface65.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface65.Evaluate

    Dim oPointAlongCurve66 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve66 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "066-PointAlongCurve")
    oPointAlongCurve66.Inputs("Curve").Add oLineByPoints59
    oPointAlongCurve66.Inputs("Point").Add oPointAlongCurve34
    oPointAlongCurve66.Inputs("TrackPoint").Add oPointAlongCurve54
    oPointAlongCurve66.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve66.Parameter("TrackFlag") = 2
    oPointAlongCurve66.Evaluate

    Dim oPointAlongCurve67 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve67 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "067-PointAlongCurve")
    oPointAlongCurve67.Inputs("Curve").Add oLineByPoints43
    oPointAlongCurve67.Inputs("Point").Add oPointAlongCurve41
    oPointAlongCurve67.Inputs("TrackPoint").Add oPointByCurveAndSurface42
    oPointAlongCurve67.Parameter("Distance") = oParamDistFunctionWith2Args62
    oPointAlongCurve67.Parameter("TrackFlag") = 1
    oPointAlongCurve67.Evaluate

    Dim oPointAlongCurve68 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve68 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "068-PointAlongCurve")
    oPointAlongCurve68.Inputs("Curve").Add oLineByPoints35
    oPointAlongCurve68.Inputs("Point").Add oPointAlongCurve53
    oPointAlongCurve68.Inputs("TrackPoint").Add oPointAlongCurve45
    oPointAlongCurve68.Parameter("Distance") = oParamDistFunctionWith2Args62
    oPointAlongCurve68.Parameter("TrackFlag") = 1
    oPointAlongCurve68.Evaluate

    Dim oPointAlongCurve69 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve69 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "069-PointAlongCurve")
    oPointAlongCurve69.Inputs("Curve").Add oLineByPoints43
    oPointAlongCurve69.Inputs("Point").Add oPointAlongCurve67
    oPointAlongCurve69.Inputs("TrackPoint").Add oPointByCurveAndSurface42
    oPointAlongCurve69.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve69.Parameter("TrackFlag") = 1
    oPointAlongCurve69.Evaluate

    Dim oLineByPoints70 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints70 = oGCFactory.CreateEntity("LineByPoints", pPOM, "070-LineByPoints")
    oLineByPoints70.Inputs("StartPoint").Add oPointAlongCurve66
    oLineByPoints70.Inputs("EndPoint").Add oPointAlongCurve54
    oLineByPoints70.Evaluate
    
    Dim oPointByCurveAndSurface71 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface71 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "071-PointByCurveAndSurface")
    oPointByCurveAndSurface71.Inputs("Curve1").Add oLineByPoints70
    oPointByCurveAndSurface71.Inputs("Surface2").Add oFacePortExtractor22
    oPointByCurveAndSurface71.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface71.Evaluate
    
    Dim oLineByPoints72 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints72 = oGCFactory.CreateEntity("LineByPoints", pPOM, "072-LineByPoints")
    oLineByPoints72.Inputs("StartPoint").Add oPointAlongCurve66
    oLineByPoints72.Inputs("EndPoint").Add oPointAlongCurve54
    oLineByPoints72.Evaluate

    Dim oPointAlongCurve73 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve73 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "073-PointAlongCurve")
    oPointAlongCurve73.Inputs("Curve").Add oLineByPoints51
    oPointAlongCurve73.Inputs("Point").Add oPointAlongCurve46
    oPointAlongCurve73.Inputs("TrackPoint").Add oPointAlongCurve48
    oPointAlongCurve73.Parameter("Distance") = oParamDistFunctionWith2Args62
    oPointAlongCurve73.Parameter("TrackFlag") = 1
    oPointAlongCurve73.Evaluate

    Dim oPointAlongCurve74 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve74 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "074-PointAlongCurve")
    oPointAlongCurve74.Inputs("Curve").Add oLineByPoints51
    oPointAlongCurve74.Inputs("Point").Add oPointAlongCurve73
    oPointAlongCurve74.Inputs("TrackPoint").Add oPointAlongCurve48
    oPointAlongCurve74.Parameter("Distance") = oParamDistConstant9
    oPointAlongCurve74.Parameter("TrackFlag") = 1
    oPointAlongCurve74.Evaluate

    Dim oLineByPoints75 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints75 = oGCFactory.CreateEntity("LineByPoints", pPOM, "075-LineByPoints")
    oLineByPoints75.Inputs("StartPoint").Add oPointAlongCurve74
    oLineByPoints75.Inputs("EndPoint").Add oPointAlongCurve46
    oLineByPoints75.Evaluate

    Dim oPlaneByPointNormal76 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal76 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "076-PlaneByPointNormal")
    oPlaneByPointNormal76.Inputs("Point").Add oPointAlongCurve63
    oPlaneByPointNormal76.Inputs("Line").Add oLineAxisPortExtractor1
    oPlaneByPointNormal76.Parameter("Range") = 0.999998
    oPlaneByPointNormal76.Evaluate

    Dim oPointAlongCurve77 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve77 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "077-PointAlongCurve")
    oPointAlongCurve77.Inputs("Curve").Add oLineByPoints59
    oPointAlongCurve77.Inputs("Point").Add oPointByCurveAndSurface65
    oPointAlongCurve77.Inputs("TrackPoint").Add oPointAlongCurve54
    oPointAlongCurve77.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve77.Parameter("TrackFlag") = 1
    oPointAlongCurve77.Evaluate

    Dim oLineByPoints78 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints78 = oGCFactory.CreateEntity("LineByPoints", pPOM, "078-LineByPoints")
    oLineByPoints78.Inputs("StartPoint").Add oPointByCurveAndSurface65
    oLineByPoints78.Inputs("EndPoint").Add oPointAlongCurve34
    oLineByPoints78.Evaluate

    Dim oLineByPoints79 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints79 = oGCFactory.CreateEntity("LineByPoints", pPOM, "079-LineByPoints")
    oLineByPoints79.Inputs("StartPoint").Add oPointAlongCurve74
    oLineByPoints79.Inputs("EndPoint").Add oPointAlongCurve54
    oLineByPoints79.Evaluate

    Dim oParamDistMeasureLength80 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength80 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "080-ParamDistMeasureLength")
    oParamDistMeasureLength80.Inputs("Curve").Add oLineByPoints78
    oParamDistMeasureLength80.Parameter("Value") = 0.025
    oParamDistMeasureLength80.Evaluate

    Dim oLineByPoints81 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints81 = oGCFactory.CreateEntity("LineByPoints", pPOM, "081-LineByPoints")
    oLineByPoints81.Inputs("StartPoint").Add oPointAlongCurve67
    oLineByPoints81.Inputs("EndPoint").Add oPointAlongCurve73
    oLineByPoints81.Evaluate

    Dim oPointAlongCurve82 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve82 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "082-PointAlongCurve")
    oPointAlongCurve82.Inputs("Curve").Add oLineByPoints81
    oPointAlongCurve82.Inputs("Point").Add oPointAlongCurve67
    oPointAlongCurve82.Inputs("TrackPoint").Add oPointByCurveAndSurface57
    oPointAlongCurve82.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve82.Parameter("TrackFlag") = 1
    oPointAlongCurve82.Evaluate

    Dim oLineByPoints83 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints83 = oGCFactory.CreateEntity("LineByPoints", pPOM, "083-LineByPoints")
    oLineByPoints83.Inputs("StartPoint").Add oPointAlongCurve82
    oLineByPoints83.Inputs("EndPoint").Add oPointAlongCurve68
    oLineByPoints83.Evaluate

    Dim oPointAlongCurve84 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve84 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "084-PointAlongCurve")
    oPointAlongCurve84.Inputs("Curve").Add oLineByPoints83
    oPointAlongCurve84.Inputs("Point").Add oPointAlongCurve68
    oPointAlongCurve84.Inputs("TrackPoint").Add oPointAlongCurve82
    oPointAlongCurve84.Parameter("Distance") = oParamDistFunctionWith2Args62
    oPointAlongCurve84.Parameter("TrackFlag") = 2
    oPointAlongCurve84.Evaluate

    Dim oLineByPoints85 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints85 = oGCFactory.CreateEntity("LineByPoints", pPOM, "085-LineByPoints")
    oLineByPoints85.Inputs("StartPoint").Add oPointAlongCurve82
    oLineByPoints85.Inputs("EndPoint").Add oPointAlongCurve84
    oLineByPoints85.Evaluate

    Dim oCpxStringByCurves86 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves86 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "086-CpxStringByCurves")
    oCpxStringByCurves86.Inputs("Curves").Add oLineByPoints72
    oCpxStringByCurves86.Inputs("Curves").Add oLineByPoints79
    oCpxStringByCurves86.Inputs("Curves").Add oLineByPoints75
    oCpxStringByCurves86.Evaluate

    Dim oSurfFromGType87 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType87 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "087-SurfFromGType")
    oSurfFromGType87.Inputs("Surface").Add oPlaneByPointNormal76
    oSurfFromGType87.Evaluate

    Dim oSurfByLinearExtrusion88 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion88 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "088-SurfByLinearExtrusion")
    oSurfByLinearExtrusion88.Inputs("PlanarCrossSection").Add oCpxStringByCurves86
    oSurfByLinearExtrusion88.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1
    oSurfByLinearExtrusion88.Evaluate

    Dim oPointAlongCurve89 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve89 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "089-PointAlongCurve")
    oPointAlongCurve89.Inputs("Curve").Add oLineParallel27
    oPointAlongCurve89.Inputs("Point").Add oPointAlongCurve64
    oPointAlongCurve89.Inputs("TrackPoint").Add oPointAlongCurve34
    oPointAlongCurve89.Parameter("Distance") = oParamDistMeasureLength80
    oPointAlongCurve89.Parameter("TrackFlag") = 1
    oPointAlongCurve89.Evaluate

    Dim oLineByPoints90 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints90 = oGCFactory.CreateEntity("LineByPoints", pPOM, "090-LineByPoints")
    oLineByPoints90.Inputs("StartPoint").Add oPointAlongCurve89
    oLineByPoints90.Inputs("EndPoint").Add oPointAlongCurve69
    oLineByPoints90.Evaluate

    Dim oPointAlongCurve91 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve91 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "091-PointAlongCurve")
    oPointAlongCurve91.Inputs("Curve").Add oLineByPoints90
    oPointAlongCurve91.Inputs("Point").Add oPointAlongCurve89
    oPointAlongCurve91.Inputs("TrackPoint").Add oPointAlongCurve69
    oPointAlongCurve91.Parameter("Distance") = oParamDistConstant5
    oPointAlongCurve91.Parameter("TrackFlag") = 2
    oPointAlongCurve91.Evaluate

    Dim oLineByPoints92 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints92 = oGCFactory.CreateEntity("LineByPoints", pPOM, "092-LineByPoints")
    oLineByPoints92.Inputs("StartPoint").Add oPointAlongCurve82
    oLineByPoints92.Inputs("EndPoint").Add oPointAlongCurve77
    oLineByPoints92.Evaluate

    Dim oPointAlongCurve93 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve93 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "093-PointAlongCurve")
    oPointAlongCurve93.Inputs("Curve").Add oLineByPoints92
    oPointAlongCurve93.Inputs("Point").Add oPointAlongCurve82
    oPointAlongCurve93.Inputs("TrackPoint").Add oPointAlongCurve77
    oPointAlongCurve93.Parameter("Distance") = oParamDistConstant7
    oPointAlongCurve93.Parameter("TrackFlag") = 1
    oPointAlongCurve93.Evaluate

    Dim oLineByPoints94 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints94 = oGCFactory.CreateEntity("LineByPoints", pPOM, "094-LineByPoints")
    oLineByPoints94.Inputs("StartPoint").Add oPointAlongCurve91
    oLineByPoints94.Inputs("EndPoint").Add oPointAlongCurve69
    oLineByPoints94.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    
    Dim oCpxStringByCurves96 As SP3DGeometricConstruction.GeometricConstruction
    Dim oSurfByLinearExtrusion97 As SP3DGeometricConstruction.GeometricConstruction

    If m_bCreateArcOnAPS = True Then
        Dim oArcByCenter95 As SP3DGeometricConstruction.GeometricConstruction
        Set oArcByCenter95 = oGCFactory.CreateEntity("ArcByCenter", pPOM, "095-ArcByCenter")
        oArcByCenter95.Inputs("Center").Add oPointAlongCurve93
        oArcByCenter95.Inputs("StartPoint").Add oPointAlongCurve69
        oArcByCenter95.Inputs("EndPoint").Add oPointAlongCurve82
        oArcByCenter95.Inputs("NormalLine").Add oLineAxisPortExtractor1
        oArcByCenter95.Inputs("TrackPoint").Add oPointAlongCurve67
        oArcByCenter95.Parameter("SweepAngle") = 1
        oArcByCenter95.Parameter("TrackFlag") = 1
        oArcByCenter95.Evaluate
    
        Set oCpxStringByCurves96 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "096-CpxStringByCurves")
        oCpxStringByCurves96.Inputs("Curves").Add oLineByPoints94
        oCpxStringByCurves96.Inputs("Curves").Add oArcByCenter95
        oCpxStringByCurves96.Inputs("Curves").Add oLineByPoints85
        oCpxStringByCurves96.Evaluate
    
        Set oSurfByLinearExtrusion97 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "097-SurfByLinearExtrusion")
        oSurfByLinearExtrusion97.Inputs("PlanarCrossSection").Add oCpxStringByCurves96, "1"
        oSurfByLinearExtrusion97.Inputs("ExtrusionLine").Add oLineAxisPortExtractor1, "1"
        oSurfByLinearExtrusion97.Evaluate
    
        oGCMacro.Output("Support", 1) = oSurfFromGType87.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion97.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion88.Output
        oGCMacro.Output("Boundary", FACEPORT_WEB) = oFacePortExtractor16.Output
    Else
        oGCMacro.Outputs("Boundary").Clear
        oGCMacro.Output("Support", 1) = oSurfFromGType87.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion88.Output
        oGCMacro.Output("Boundary", FACEPORT_WEB) = oFacePortExtractor16.Output
    End If
       
End Sub
