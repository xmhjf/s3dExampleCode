VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklgPltForTube"
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
    MISSING_MANDATORY_INPUT = 1
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
'
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select a Tube Built-up", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("Node", "Select the Node to measure the distance from", "IJPoint", 0, 1)
    Call pGeometricConstructionDefinition.AddControlledInput("Port_9_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1", "IJGeometry")

    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddParameter("Distance", "Distance", 8, 1, 59, 0, 0, 1#)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "001-LineAxisPortExtractor")
    oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oLineAxisPortExtractor1.Evaluate
    pGeometricConstruction.ControlledInputs("Port_1_1").Clear
    pGeometricConstruction.ControlledInputs("Port_1_1").AddElements oLineAxisPortExtractor1.ControlledInputs("Port")

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
    Dim bMeasureFromStartPoint As Boolean: Let bMeasureFromStartPoint = True
    If pGeometricConstruction.Inputs("Node").Count = 1 Then
        Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = Position_FromPoint(pGeometricConstruction.Input("Node"))
        Dim pPositionOfPointAtCurveStart As IJDPosition: Set pPositionOfPointAtCurveStart = Position_FromPoint(oPointAtCurveStart3)
        Dim pPositionOfPointAtCurveEnd As IJDPosition: Set pPositionOfPointAtCurveEnd = Position_FromPoint(oPointAtCurveEnd4)
        If pPositionOfNode.DistPt(pPositionOfPointAtCurveEnd) < pPositionOfNode.DistPt(pPositionOfPointAtCurveStart) Then Let bMeasureFromStartPoint = False
    End If
    If bMeasureFromStartPoint Then
        oPointAlongCurve6.Inputs("Point").Add oPointAtCurveStart3
        oPointAlongCurve6.Inputs("TrackPoint").Add oPointAtCurveEnd4
     Else
        oPointAlongCurve6.Inputs("Point").Add oPointAtCurveEnd4
        oPointAlongCurve6.Inputs("TrackPoint").Add oPointAtCurveStart3
    End If
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
    oFacePortExtractor9.Parameter("TrackFlag") = 1
    oFacePortExtractor9.Parameter("Offset") = 0#
    oFacePortExtractor9.Evaluate
    pGeometricConstruction.ControlledInputs("Port_9_1").Clear
    pGeometricConstruction.ControlledInputs("Port_9_1").AddElements oFacePortExtractor9.ControlledInputs("Port")

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Support", 1) = oSurfFromGType8.Output
    oGCMacro.Output("Boundary", "Port_9_1") = oFacePortExtractor9.Output
End Sub
Private Function Position_FromPoint(pPoint As IJPoint) As IJDPosition
    ' prepare result
    Dim pPosition As IJDPosition: Set pPosition = New DPosition
    
    Dim dX As Double, dY As Double, dZ As Double
    Call pPoint.GetPoint(dX, dY, dZ)
    Call pPosition.Set(dX, dY, dZ)
    
    ' return result
    Set Position_FromPoint = pPosition
End Function
