VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtractCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const PORT_NOT_FOUND = 2
Private Enum Errors
    ORTHOGONAL_MEMBER_NOT_TUBE = 1
End Enum
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart1", "Select a circular MemberPart", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("MemberPart2", "Select an orthogonal MemberPart", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2", "IJGeometry")
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    ' define errors
    Call pGeometricConstructionDefinition.AddErrorValue(ORTHOGONAL_MEMBER_NOT_TUBE, "OrthogonalMemberNotTube", "The orthogonal member does not have a circular cross-section")
    
    Call pGeometricConstructionDefinition.AddOutput(7, "Circle")
    Call pGeometricConstructionDefinition.AddOutput(GCSurfaceBody2, "Boundary")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Dim oMemberPart As Object: Set oMemberPart = pGeometricConstruction.Inputs("MemberPart2").Item(1)
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
        Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
        oLineAxisPortExtractor1.Inputs("MemberPart").Add oMemberPart
        oLineAxisPortExtractor1.Evaluate
    Else
        Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "001-AxisPortExtractor")
        oLineAxisPortExtractor1.Input("Connectable") = oMemberPart
        oLineAxisPortExtractor1.Parameter("GeometrySelector") = 4 ' GCStable
        oLineAxisPortExtractor1.Evaluate
    End If
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("AxisPort2"), oLineAxisPortExtractor1.ControlledInputs("Port"))

'    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
'    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
'    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart2").Item(1)
'    oLineAxisPortExtractor1.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("AxisPort2")
'    oLineAxisPortExtractor1.Public = True
'    oLineAxisPortExtractor1.Evaluate
'    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("AxisPort2"), oLineAxisPortExtractor1.ControlledInputs("Port"))

    Dim oLineAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor2 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "002-LineAxisPortExtractor")
    oLineAxisPortExtractor2.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart1").Item(1)
    oLineAxisPortExtractor2.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("AxisPort1")
    oLineAxisPortExtractor2.Public = True
    oLineAxisPortExtractor2.Evaluate
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("AxisPort1"), oLineAxisPortExtractor2.ControlledInputs("Port"))

    Dim oPointAtCurveMiddle3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle3 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "003-PointAtCurveMiddle")
    oPointAtCurveMiddle3.Inputs("Curve").Add oLineAxisPortExtractor2
    oPointAtCurveMiddle3.Evaluate
    
    Dim oCSByLines4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines4 = oGCFactory.CreateEntity("CSByLines", pPOM, "004-CSByLines")
    oCSByLines4.Inputs("AxisLine1").Add oLineAxisPortExtractor2
    oCSByLines4.Inputs("AxisLine2").Add oLineAxisPortExtractor1
    oCSByLines4.Parameter("AxesRoles") = 3
    oCSByLines4.Parameter("CSOrientation") = 1
    oCSByLines4.Public = True
    oCSByLines4.Evaluate
  
    Dim oPlaneFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS5 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "005-PlaneFromCS")
    oPlaneFromCS5.Inputs("CoordinateSystem").Add oCSByLines4
    oPlaneFromCS5.Parameter("LookingAxis") = 3
    oPlaneFromCS5.Parameter("Offset") = 0
    oPlaneFromCS5.Parameter("Range") = 5#
    oPlaneFromCS5.Public = True
    oPlaneFromCS5.Evaluate

    Dim oCSFromMember6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember6 = oGCFactory.CreateEntity("CSFromMember", pPOM, "006-CSFromMember")
    oCSFromMember6.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart1").Item(1)
    oCSFromMember6.Inputs("Point").Add oPointAtCurveMiddle3
    oCSFromMember6.Evaluate
    
    Dim oFacePortExtractor7 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor7 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "007-FacePortExtractor")
    oFacePortExtractor7.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart1").Item(1)
    oFacePortExtractor7.Inputs("CoordinateSystem").Add oCSFromMember6
    oFacePortExtractor7.ControlledInputs("Port").AddElements pGeometricConstruction.ControlledInputs("FacePort1")
    oFacePortExtractor7.Parameter("GeometrySelector") = 4 ' stable 2
    oFacePortExtractor7.Parameter("FacesContext") = 2
    oFacePortExtractor7.Parameter("LookingAxis") = 3 ' 1
    oFacePortExtractor7.Parameter("TrackFlag") = 1
    oFacePortExtractor7.Parameter("Offset") = 0 '-0.1
    oFacePortExtractor7.Public = True
    Call GCEvaluate(oFacePortExtractor7, pGeometricConstruction, PORT_NOT_FOUND, ORTHOGONAL_MEMBER_NOT_TUBE)
    Call Elements_ReplaceElements(pGeometricConstruction.ControlledInputs("FacePort1"), oFacePortExtractor7.ControlledInputs("Port"))

    Dim oCurveByIntersection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection8 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "008-CurveByIntersection")
    oCurveByIntersection8.Inputs("Surface1").Add oFacePortExtractor7
    oCurveByIntersection8.Inputs("Surface2").Add oPlaneFromCS5
    oCurveByIntersection8.Parameter("TrackFlag") = 1
    oCurveByIntersection8.Public = True
    oCurveByIntersection8.Evaluate
    
    If Not TypeOf oCurveByIntersection8.Output Is IJCircle Then
        Call GCProcessError(pGeometricConstruction, , ORTHOGONAL_MEMBER_NOT_TUBE)
    End If

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Circle", 1) = oCurveByIntersection8.Output
    oGCMacro.Output("Boundary", 1) = oFacePortExtractor7.Output
End Sub
Sub Elements_ReplaceElements(pElementsToBeReplaced As IJElements, pElementsReplacing As IJElements)
    If pElementsToBeReplaced.Count = 1 Then
        If Not pElementsToBeReplaced(1) Is pElementsReplacing(1) Then
            pElementsToBeReplaced.Clear
            pElementsToBeReplaced.AddElements pElementsReplacing
        End If
    Else
        pElementsToBeReplaced.AddElements pElementsReplacing
    End If
End Sub

