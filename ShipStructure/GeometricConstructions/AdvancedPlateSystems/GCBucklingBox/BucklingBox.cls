VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_oRegularAPSHelper As Object
Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCToDoDelegate

Dim m_oAdvancedDebug As New AdvancedDebug

Private Property Get Source() As String
    Let Source = "GCCustomPlate.CustomPlate"
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
'
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
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a Box Built-up", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort4")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_6_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddParameter("Distance", "Distance", 8, 1, 59, 0, 0, 1.5)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port_6_1").Clear
    pGeometricConstruction.ControlledInputs("Port_6_1").AddElements oLineAxisPortExtractor1.ControlledInputs("Port")

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Distance")
    oParamDistConstant2.Evaluate

    Dim oPointAtCurveStart3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart3 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "003-PointAtCurveStart")
    oPointAtCurveStart3.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveStart3.Evaluate

    Dim oPointAtCurveEnd4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd4 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "004-PointAtCurveEnd")
    oPointAtCurveEnd4.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAtCurveEnd4.Evaluate

    Dim oCSFromMember5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSFromMember5 = oGCFactory.CreateEntity("CSFromMember", pPOM, "005-CSFromMember")
    oCSFromMember5.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oCSFromMember5.Inputs("Point").Add oPointAtCurveStart3
    oCSFromMember5.Evaluate

    Dim oPointAlongCurve6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "006-PointAlongCurve")
    oPointAlongCurve6.Inputs("Curve").Add oLineAxisPortExtractor1
    oPointAlongCurve6.Inputs("Point").Add oPointAtCurveStart3
    oPointAlongCurve6.Inputs("TrackPoint").Add oPointAtCurveEnd4
    oPointAlongCurve6.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve6.Parameter("TrackFlag") = 1
    oPointAlongCurve6.Evaluate

    Dim oPlaneByPointNormal7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal7 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "007-PlaneByPointNormal")
    oPlaneByPointNormal7.Inputs("Point").Add oPointAlongCurve6
    oPlaneByPointNormal7.Inputs("Line").Add oLineAxisPortExtractor1
    oPlaneByPointNormal7.Parameter("Range") = 3#
    oPlaneByPointNormal7.Evaluate

    Dim oSurfFromGType8 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType8 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "008-SurfFromGType")
    oSurfFromGType8.Inputs("Surface").Add oPlaneByPointNormal7
    oSurfFromGType8.Evaluate

    Dim oFacePortExtractor9 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor9 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "009-FacePortExtractor")
    oFacePortExtractor9.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor9.Inputs("CoordinateSystem").Add oCSFromMember5
    oFacePortExtractor9.Parameter("GeometrySelector") = 4
    oFacePortExtractor9.Parameter("FacesContext") = 2
    oFacePortExtractor9.Parameter("LookingAxis") = 3
    oFacePortExtractor9.Parameter("SurfaceType") = 1
    oFacePortExtractor9.Parameter("TrackFlag") = 2
    oFacePortExtractor9.Parameter("Offset") = 0#
    oFacePortExtractor9.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort2").Clear
    pGeometricConstruction.ControlledInputs("FacePort2").AddElements oFacePortExtractor9.ControlledInputs("Port")

    Dim oFacePortExtractor10 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor10 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "010-FacePortExtractor")
    oFacePortExtractor10.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor10.Inputs("CoordinateSystem").Add oCSFromMember5
    oFacePortExtractor10.Parameter("GeometrySelector") = 4
    oFacePortExtractor10.Parameter("FacesContext") = 2
    oFacePortExtractor10.Parameter("LookingAxis") = 2
    oFacePortExtractor10.Parameter("SurfaceType") = 1
    oFacePortExtractor10.Parameter("TrackFlag") = 2
    oFacePortExtractor10.Parameter("Offset") = 0#
    oFacePortExtractor10.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort3").Clear
    pGeometricConstruction.ControlledInputs("FacePort3").AddElements oFacePortExtractor10.ControlledInputs("Port")

    Dim oFacePortExtractor11 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor11 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "011-FacePortExtractor")
    oFacePortExtractor11.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor11.Inputs("CoordinateSystem").Add oCSFromMember5
    oFacePortExtractor11.Parameter("GeometrySelector") = 4
    oFacePortExtractor11.Parameter("FacesContext") = 2
    oFacePortExtractor11.Parameter("LookingAxis") = 2
    oFacePortExtractor11.Parameter("SurfaceType") = 1
    oFacePortExtractor11.Parameter("TrackFlag") = 1
    oFacePortExtractor11.Parameter("Offset") = 0#
    oFacePortExtractor11.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort4").Clear
    pGeometricConstruction.ControlledInputs("FacePort4").AddElements oFacePortExtractor11.ControlledInputs("Port")

    Dim oFacePortExtractor12 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor12 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "012-FacePortExtractor")
    oFacePortExtractor12.Inputs("Connectable").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oFacePortExtractor12.Inputs("CoordinateSystem").Add oCSFromMember5
    oFacePortExtractor12.Parameter("GeometrySelector") = 4
    oFacePortExtractor12.Parameter("FacesContext") = 2
    oFacePortExtractor12.Parameter("LookingAxis") = 3
    oFacePortExtractor12.Parameter("SurfaceType") = 1
    oFacePortExtractor12.Parameter("TrackFlag") = 1
    oFacePortExtractor12.Parameter("Offset") = 0#
    oFacePortExtractor12.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort1").Clear
    pGeometricConstruction.ControlledInputs("FacePort1").AddElements oFacePortExtractor12.ControlledInputs("Port")

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oSurfFromGType8.Output
    
    Dim oMember As Object
    Set oMember = pGeometricConstruction.Inputs("MemberPart")(1)
    Dim oSD_Member As New StructDetailObjects.MemberPart
    Set oSD_Member.object = oMember
    
    If oSD_Member.IsCrossSectionABuiltUp Then
        oGCMacro.Output("Boundary", "FacePort1") = oFacePortExtractor12.Output
        oGCMacro.Output("Boundary", "FacePort2") = oFacePortExtractor9.Output
        oGCMacro.Output("Boundary", "FacePort3") = oFacePortExtractor10.Output
        oGCMacro.Output("Boundary", "FacePort4") = oFacePortExtractor11.Output
    Else
        oGCMacro.Output("Boundary", "FacePort1") = oFacePortExtractor12.Output
    End If
    
End Sub
