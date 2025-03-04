VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacePortNearPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a MemberPart", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("PlateSystem", "Select a PlateSystem", "IJPlateSystem AND [GCCMNSTRDefinitions.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1a")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1b")
    Call pGeometricConstructionDefinition.AddOutput(6, "FacePort")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort").Clear
    pGeometricConstruction.ControlledInputs("AxisPort").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")
    
    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0002-FacePortExtractor0")
    oFacePortExtractor02.Inputs("Connectable").Add pGeometricConstruction.Inputs("PlateSystem").Item(1), "1"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 2
    oFacePortExtractor02.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort2").Clear
    pGeometricConstruction.ControlledInputs("FacePort2").Add oFacePortExtractor02.ControlledInputs("Port").Item("1")
    
    Dim oCSByPlane3 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane3 = oGCFactory.CreateEntity("CSByPlane", pPOM, "0003-CSByPlane")
    oCSByPlane3.Inputs("Plane").Add pGeometricConstruction.Inputs("PlateSystem").Item(1), "1"
    oCSByPlane3.Evaluate

    Dim oLineFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS4 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0004-LineFromCS")
    oLineFromCS4.Inputs("CoordinateSystem").Add oCSByPlane3, "1"
    oLineFromCS4.Parameter("LookingAxis") = 3
    oLineFromCS4.Parameter("Length") = 2#
    oLineFromCS4.Parameter("CSOrientation") = 1
    oLineFromCS4.Parameter("LineJustification") = 1
    oLineFromCS4.Evaluate

    Dim oCSByLines5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines5 = oGCFactory.CreateEntity("CSByLines", pPOM, "0005-CSByLines")
    oCSByLines5.Inputs("AxisLine1").Add oLineAxisPortExtractor1, "1"
    oCSByLines5.Inputs("AxisLine2").Add oLineFromCS4, "1"
    oCSByLines5.Parameter("AxesRoles") = 6
    oCSByLines5.Parameter("CSOrientation") = 1
    oCSByLines5.Parameter("TrackFlag") = 1
    oCSByLines5.Evaluate

    Dim oFacePortExtractor6 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor6 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0006-FacePortExtractor")
    oFacePortExtractor6.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "2"
    oFacePortExtractor6.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oFacePortExtractor6.Parameter("GeometrySelector") = 4
    oFacePortExtractor6.Parameter("FacesContext") = 2
    oFacePortExtractor6.Parameter("LookingAxis") = 3
    oFacePortExtractor6.Parameter("IntersectingPlane") = 0
    oFacePortExtractor6.Parameter("SurfaceType") = 1
    oFacePortExtractor6.Parameter("TrackFlag") = 1
    oFacePortExtractor6.Parameter("Offset") = 0#
    oFacePortExtractor6.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort1a").Clear
    pGeometricConstruction.ControlledInputs("FacePort1a").Add oFacePortExtractor6.ControlledInputs("Port").Item("1")
    
    Dim oFacePortExtractor7 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor7 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0007-FacePortExtractor")
    oFacePortExtractor7.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1), "1"
    oFacePortExtractor7.Inputs("CoordinateSystem").Add oCSByLines5, "1"
    oFacePortExtractor7.Parameter("GeometrySelector") = 4
    oFacePortExtractor7.Parameter("FacesContext") = 2
    oFacePortExtractor7.Parameter("LookingAxis") = 3
    oFacePortExtractor7.Parameter("IntersectingPlane") = 0
    oFacePortExtractor7.Parameter("SurfaceType") = 1
    oFacePortExtractor7.Parameter("TrackFlag") = 2
    oFacePortExtractor7.Parameter("Offset") = 0#
    oFacePortExtractor7.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort1b").Clear
    pGeometricConstruction.ControlledInputs("FacePort1b").Add oFacePortExtractor7.ControlledInputs("Port").Item("1")
    
    Dim oParamDistMeasureParallel8 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel8 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0008-ParamDistMeasureParallel")
    oParamDistMeasureParallel8.Inputs("Graphics").Add oFacePortExtractor02, "1"
    oParamDistMeasureParallel8.Inputs("Graphics").Add oFacePortExtractor6, "2"
    oParamDistMeasureParallel8.Parameter("Value") = 1.3
    oParamDistMeasureParallel8.Evaluate

    Dim oParamDistMeasureParallel9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel9 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0009-ParamDistMeasureParallel")
    oParamDistMeasureParallel9.Inputs("Graphics").Add oFacePortExtractor02, "1"
    oParamDistMeasureParallel9.Inputs("Graphics").Add oFacePortExtractor7, "2"
    oParamDistMeasureParallel9.Parameter("Value") = 0.4
    oParamDistMeasureParallel9.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    If oParamDistMeasureParallel8.Parameter("Value") < oParamDistMeasureParallel9.Parameter("Value") Then
        oGCMacro.Output("FacePort", 1) = oFacePortExtractor6.Output
    Else
        oGCMacro.Output("FacePort", 1) = oFacePortExtractor7.Output
    End If
End Sub
