VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SidePlateAxial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   File:           SidePlateAxial.cls
'   ProgID:         GCSidePlateAxial.SidePlateAxial
'   Author:         Alligators Team(India)
'   Creation Date:  Tuesday, Oct 12 2010
'   Description:
'      This macro places an advanced plate on a built up member (supported) which is
'      connected to a standard member or built-up (supporting).
'
'   Change History:
'   dd.mmm.yyyy     who              change description
'   -----------     ---              ------------------
'
'   13.Oct.2010     svsmylav         CR-CP-181614 Modified code prepared by APN.
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oSidePlateHelper As Object
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
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

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
    
    Set m_oSidePlateHelper = CreateObject("GCAPSWithTrimBack2.SidePlateHelper")
End Sub
Private Sub Class_Terminate()
    Set m_oSidePlateHelper = Nothing
    
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub

Private Sub IJGCConnectedElementSC_PostDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PostDisconnectExternalRels"))
    
    Call m_oSidePlateHelper.PostDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PreDisconnectExternalRels"))
    
    Call m_oSidePlateHelper.PreDisconnectExternalRels(pGC, pInfo)

    Call m_oAdvancedDebug.ExitMethod

End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oSidePlateHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oSidePlateHelper.PostInitialize(pGCDefinition)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Call m_oSidePlateHelper.PreEvaluate(pGC, pPOM)
    Call Evaluate(pGC, pPOM)
    Call m_oSidePlateHelper.PostEvaluate(pGC, pPOM)

    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Call m_oSidePlateHelper.ProcessError(pGC, Err.Number)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oSidePlateHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oSidePlateHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Call m_oSidePlateHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))

    Set IJGCToDoDelegate_ToDoDelegate = m_oSidePlateHelper.ToDoDelegate(pGC)

    Call m_oAdvancedDebug.ExitMethod
End Property
'
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("SupportingMemberPart", "Select Supporting Member (a RolledMember or a DesignedMember)", "ISPSWallPart OR ISPSSlabEntity or ISPSMemberPartPrismatic or IJStiffenerPart or IJPlatePart or ISPSDesignedMember or IJPlateSystem", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("MemberPart", "Select Supported Member (a RolledMember or a DesignedMember)", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("EdgePort")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort_6_1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_6_1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort_7_1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_7_1")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("ChLen", "Chamfer Length", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("ChWidth", "Chamfer Width", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("SetBack", "SetBack", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Bottom", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Top", 2)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = pGeometricConstruction.Parameter("ChLen")
    oParamDistConstant1.Evaluate

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("ChWidth")
    oParamDistConstant2.Evaluate

    Dim oPointAtCurveMiddle3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle3 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "003-PointAtCurveMiddle")
    oPointAtCurveMiddle3.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oPointAtCurveMiddle3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Length")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("SetBack")
    oParamDistConstant5.Evaluate

    Dim oOrderAxes6 As SP3DGeometricConstruction.GeometricConstruction
    Set oOrderAxes6 = oGCFactory.CreateEntity("OrderAxes", pPOM, "006-OrderAxes")
    oOrderAxes6.Inputs("MemberParts").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oOrderAxes6.Inputs("SupportingMemberPart").Add pGeometricConstruction.Inputs("SupportingMemberPart").Item(1)
    oOrderAxes6.Parameter("Support") = CLng(pGeometricConstruction.Parameter("Location"))
    oOrderAxes6.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort_7_1").Clear
    pGeometricConstruction.ControlledInputs("AxisPort_7_1").AddElements oOrderAxes6.ControlledInputs("Port")

    Dim oLineFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7 = oGCFactory.CreateEntity("LineFromCS", pPOM, "007-LineFromCS")
    oLineFromCS7.Inputs("CoordinateSystem").Add oOrderAxes6.Output("CoordinateSystem", 1)
    oLineFromCS7.Parameter("LookingAxis") = 3
    oLineFromCS7.Parameter("Length") = 3#
    oLineFromCS7.Parameter("CSOrientation") = 3 - CLng(pGeometricConstruction.Parameter("Location"))
    oLineFromCS7.Parameter("LineJustification") = 1
    oLineFromCS7.Evaluate

    Dim oLineFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS8 = oGCFactory.CreateEntity("LineFromCS", pPOM, "008-LineFromCS")
    oLineFromCS8.Inputs("CoordinateSystem").Add oOrderAxes6.Output("CoordinateSystem", 1)
    oLineFromCS8.Parameter("LookingAxis") = 3
    oLineFromCS8.Parameter("Length") = 3#
    oLineFromCS8.Parameter("CSOrientation") = 1
    oLineFromCS8.Parameter("LineJustification") = 1
    oLineFromCS8.Evaluate

    Dim oExtractPorts9 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts9 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "009-ExtractPorts")
    oExtractPorts9.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SupportingMemberPart").Item(1)
    oExtractPorts9.Inputs("ZAxis").Add oLineFromCS8
    oExtractPorts9.Parameter("Support") = CLng(pGeometricConstruction.Parameter("Location"))
    oExtractPorts9.Parameter("Debug") = 0
    oExtractPorts9.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort_6_1").Clear
    pGeometricConstruction.ControlledInputs("AxisPort_6_1").AddElements oExtractPorts9.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort_6_1").Clear
    pGeometricConstruction.ControlledInputs("FacePort_6_1").AddElements oExtractPorts9.ControlledInputs("FacePort")

    Dim oExtractPorts10 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts10 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "010-ExtractPorts")
    oExtractPorts10.Inputs("MemberPart").Add pGeometricConstruction.Inputs("MemberPart").Item(1)
    oExtractPorts10.Inputs("ZAxis").Add oLineFromCS8
    oExtractPorts10.Parameter("Support") = CLng(pGeometricConstruction.Parameter("Location"))
    oExtractPorts10.Parameter("Debug") = 0
    oExtractPorts10.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort_7_1").Clear
    pGeometricConstruction.ControlledInputs("AxisPort_7_1").AddElements oExtractPorts10.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort_7_1").Clear
    pGeometricConstruction.ControlledInputs("FacePort_7_1").AddElements oExtractPorts10.ControlledInputs("FacePort")

    Dim oFacePortExtractor11 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor11 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "011-FacePortExtractor")
    oFacePortExtractor11.Inputs("Connectable").Add pGeometricConstruction.Inputs("SupportingMemberPart").Item(1)
    oFacePortExtractor11.Inputs("CoordinateSystem").Add oOrderAxes6.Output("CoordinateSystem", 1)
    oFacePortExtractor11.Parameter("GeometrySelector") = 2
    oFacePortExtractor11.Parameter("LookingAxis") = 3
    oFacePortExtractor11.Parameter("SurfaceType") = 2
    oFacePortExtractor11.Parameter("TrackFlag") = CLng(pGeometricConstruction.Parameter("Location"))
    oFacePortExtractor11.Parameter("Offset") = 0#
    oFacePortExtractor11.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort").Clear
    pGeometricConstruction.ControlledInputs("FacePort").AddElements oFacePortExtractor11.ControlledInputs("Port")

    Dim oEdgePortExtractor12 As SP3DGeometricConstruction.GeometricConstruction
    Set oEdgePortExtractor12 = oGCFactory.CreateEntity("EdgePortExtractor", pPOM, "012-EdgePortExtractor")
    oEdgePortExtractor12.Inputs("Connectable").Add pGeometricConstruction.Inputs("SupportingMemberPart").Item(1)
    oEdgePortExtractor12.Inputs("FacePort").Add oFacePortExtractor11
    oEdgePortExtractor12.Inputs("CoordinateSystem").Add oOrderAxes6.Output("CoordinateSystem", 1)
    oEdgePortExtractor12.Parameter("GeometrySelector") = 2
    oEdgePortExtractor12.Parameter("LookingAxis") = 2
    oEdgePortExtractor12.Parameter("TrackFlag") = 2
    oEdgePortExtractor12.Evaluate
    pGeometricConstruction.ControlledInputs("EdgePort").Clear
    pGeometricConstruction.ControlledInputs("EdgePort").AddElements oEdgePortExtractor12.ControlledInputs("Port")
        
    Dim oLineExtend13 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend13 = oGCFactory.CreateEntity("LineExtend", pPOM, "013-LineExtend")
    oLineExtend13.Inputs("Line").Add oExtractPorts10.Output("MemberAxis", 1)
    oLineExtend13.Inputs("Curve1").Add oExtractPorts9.Output("MemberAxis", 1)
    oLineExtend13.Parameter("Context1") = 1
    oLineExtend13.Parameter("Context2") = 3
    oLineExtend13.Evaluate
    
    Dim oPointByCurves14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves14 = oGCFactory.CreateEntity("PointByCurves", pPOM, "014-PointByCurves")
    oPointByCurves14.Inputs("Curve1").Add oEdgePortExtractor12
    oPointByCurves14.Inputs("Curve2").Add oLineExtend13 'oExtractPorts10.Output("MemberAxis", 1)
    oPointByCurves14.Parameter("TrackFlag") = 1
    oPointByCurves14.Evaluate
    
    Dim oCSByCS15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS15 = oGCFactory.CreateEntity("CSByCS", pPOM, "015-CSByCS")
    oCSByCS15.Inputs("CoordinateSystem").Add oOrderAxes6.Output("CoordinateSystem", 1)
    oCSByCS15.Inputs("Origin").Add oPointByCurves14
    oCSByCS15.Parameter("AxesRoles") = 1
    oCSByCS15.Parameter("CSOrientation") = 3 - CLng(pGeometricConstruction.Parameter("Location"))
    oCSByCS15.Parameter("TrackFlag") = 1
    oCSByCS15.Evaluate
    
    Dim oLineFromCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS16 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0016-LineFromCS")
    oLineFromCS16.Inputs("CoordinateSystem").Add oCSByCS15
    oLineFromCS16.Parameter("LookingAxis") = 3
    oLineFromCS16.Parameter("Length") = -5
    oLineFromCS16.Parameter("CSOrientation") = 3 - CLng(pGeometricConstruction.Parameter("Location"))
    oLineFromCS16.Parameter("LineJustification") = 1
    oLineFromCS16.Evaluate

    Dim oSurfByLinearExtrusion17 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion17 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "017-SurfByLinearExtrusion")
    oSurfByLinearExtrusion17.Inputs("PlanarCrossSection").Add oEdgePortExtractor12
    oSurfByLinearExtrusion17.Inputs("ExtrusionLine").Add oLineFromCS16 'oLineFromCS7 '
    oSurfByLinearExtrusion17.Evaluate

    Dim oPointAlongCurve18 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve18 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "018-PointAlongCurve")
    oPointAlongCurve18.Inputs("Curve").Add oLineExtend13 'oExtractPorts10.Output("MemberAxis", 1)
    oPointAlongCurve18.Inputs("Point").Add oPointByCurves14
    oPointAlongCurve18.Inputs("TrackPoint").Add oPointAtCurveMiddle3
    oPointAlongCurve18.Parameter("Distance") = oParamDistConstant4
    oPointAlongCurve18.Parameter("TrackFlag") = 1
    oPointAlongCurve18.Evaluate

    Dim oLineParallel19 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel19 = oGCFactory.CreateEntity("LineParallel", pPOM, "019-LineParallel")
    oLineParallel19.Inputs("Line").Add oExtractPorts10.Output("CurveRight", 1)
    oLineParallel19.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineParallel19.Inputs("TrackPoint").Add oPointByCurves14
    oLineParallel19.Parameter("Distance") = oParamDistConstant2
    oLineParallel19.Parameter("TrackFlag") = 2
    oLineParallel19.Evaluate

    Dim oLineParallel20 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel20 = oGCFactory.CreateEntity("LineParallel", pPOM, "020-LineParallel")
    oLineParallel20.Inputs("Line").Add oExtractPorts10.Output("CurveLeft", 1)
    oLineParallel20.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineParallel20.Inputs("TrackPoint").Add oPointByCurves14
    oLineParallel20.Parameter("Distance") = oParamDistConstant2
    oLineParallel20.Parameter("TrackFlag") = 2
    oLineParallel20.Evaluate

    Dim oLineAtAngle21 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle21 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "021-LineAtAngle")
    oLineAtAngle21.Inputs("Line").Add oLineExtend13 'oExtractPorts10.Output("MemberAxis", 1)
    oLineAtAngle21.Inputs("Point").Add oPointAlongCurve18
    oLineAtAngle21.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineAtAngle21.Parameter("Angle") = 1.570796
    oLineAtAngle21.Parameter("TrackFlag") = 1
    oLineAtAngle21.Evaluate

    Dim oPlaneByPointNormal22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal22 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "022-PlaneByPointNormal")
    oPlaneByPointNormal22.Inputs("Point").Add oPointAlongCurve18
    oPlaneByPointNormal22.Inputs("Line").Add oLineExtend13 'oExtractPorts10.Output("MemberAxis", 1)
    oPlaneByPointNormal22.Parameter("Range") = 5#
    oPlaneByPointNormal22.Evaluate

    Dim oSurfFromGType23 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType23 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "023-SurfFromGType")
    oSurfFromGType23.Inputs("Surface").Add oPlaneByPointNormal22
    oSurfFromGType23.Evaluate

    Dim oLineParallel24 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel24 = oGCFactory.CreateEntity("LineParallel", pPOM, "024-LineParallel")
    oLineParallel24.Inputs("Line").Add oLineAtAngle21
    oLineParallel24.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineParallel24.Inputs("TrackPoint").Add oPointByCurves14
    oLineParallel24.Parameter("Distance") = oParamDistConstant5
    oLineParallel24.Parameter("TrackFlag") = 1
    oLineParallel24.Evaluate

    Dim oLineParallel25 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel25 = oGCFactory.CreateEntity("LineParallel", pPOM, "025-LineParallel")
    oLineParallel25.Inputs("Line").Add oLineAtAngle21
    oLineParallel25.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineParallel25.Inputs("TrackPoint").Add oPointByCurves14
    oLineParallel25.Parameter("Distance") = 0.003
    oLineParallel25.Parameter("TrackFlag") = 2
    oLineParallel25.Evaluate

    Dim oPointByCurves26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves26 = oGCFactory.CreateEntity("PointByCurves", pPOM, "026-PointByCurves")
    oPointByCurves26.Inputs("Curve1").Add oExtractPorts10.Output("CurveRight", 1)
    oPointByCurves26.Inputs("Curve2").Add oLineParallel24
    oPointByCurves26.Parameter("TrackFlag") = 1
    oPointByCurves26.Evaluate

    Dim oPointByCurves27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves27 = oGCFactory.CreateEntity("PointByCurves", pPOM, "027-PointByCurves")
    oPointByCurves27.Inputs("Curve1").Add oExtractPorts10.Output("CurveLeft", 1)
    oPointByCurves27.Inputs("Curve2").Add oLineParallel24
    oPointByCurves27.Parameter("TrackFlag") = 1
    oPointByCurves27.Evaluate

    Dim oPointByCurves28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves28 = oGCFactory.CreateEntity("PointByCurves", pPOM, "028-PointByCurves")
    oPointByCurves28.Inputs("Curve1").Add oLineParallel25
    oPointByCurves28.Inputs("Curve2").Add oExtractPorts10.Output("CurveRight", 1)
    oPointByCurves28.Parameter("TrackFlag") = 1
    oPointByCurves28.Evaluate

    Dim oPointByCurves29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves29 = oGCFactory.CreateEntity("PointByCurves", pPOM, "029-PointByCurves")
    oPointByCurves29.Inputs("Curve1").Add oLineParallel25
    oPointByCurves29.Inputs("Curve2").Add oExtractPorts10.Output("CurveLeft", 1)
    oPointByCurves29.Parameter("TrackFlag") = 1
    oPointByCurves29.Evaluate

    Dim oLineByPoints30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints30 = oGCFactory.CreateEntity("LineByPoints", pPOM, "030-LineByPoints")
    oLineByPoints30.Inputs("StartPoint").Add oPointByCurves26
    oLineByPoints30.Inputs("EndPoint").Add oPointByCurves28
    oLineByPoints30.Evaluate

    Dim oLineByPoints31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints31 = oGCFactory.CreateEntity("LineByPoints", pPOM, "031-LineByPoints")
    oLineByPoints31.Inputs("StartPoint").Add oPointByCurves27
    oLineByPoints31.Inputs("EndPoint").Add oPointByCurves29
    oLineByPoints31.Evaluate

    Dim oLineParallel32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineParallel32 = oGCFactory.CreateEntity("LineParallel", pPOM, "032-LineParallel")
    oLineParallel32.Inputs("Line").Add oLineParallel24
    oLineParallel32.Inputs("Plane").Add oExtractPorts10.Output("Support", 1)
    oLineParallel32.Inputs("TrackPoint").Add oPointByCurves14
    oLineParallel32.Parameter("Distance") = oParamDistConstant1
    oLineParallel32.Parameter("TrackFlag") = 1
    oLineParallel32.Evaluate

    Dim oPointByCurves33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves33 = oGCFactory.CreateEntity("PointByCurves", pPOM, "033-PointByCurves")
    oPointByCurves33.Inputs("Curve1").Add oLineParallel32
    oPointByCurves33.Inputs("Curve2").Add oLineParallel20
    oPointByCurves33.Parameter("TrackFlag") = 1
    oPointByCurves33.Evaluate

    Dim oPointByCurves34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves34 = oGCFactory.CreateEntity("PointByCurves", pPOM, "034-PointByCurves")
    oPointByCurves34.Inputs("Curve1").Add oLineParallel19
    oPointByCurves34.Inputs("Curve2").Add oLineParallel32
    oPointByCurves34.Parameter("TrackFlag") = 1
    oPointByCurves34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointByCurves33
    oLineByPoints35.Inputs("EndPoint").Add oPointByCurves27
    oLineByPoints35.Evaluate

    Dim oLineByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints36 = oGCFactory.CreateEntity("LineByPoints", pPOM, "036-LineByPoints")
    oLineByPoints36.Inputs("StartPoint").Add oPointByCurves34
    oLineByPoints36.Inputs("EndPoint").Add oPointByCurves26
    oLineByPoints36.Evaluate

    Dim oLineExtend37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend37 = oGCFactory.CreateEntity("LineExtend", pPOM, "037-LineExtend")
    oLineExtend37.Inputs("Line").Add oLineParallel19
    oLineExtend37.Inputs("Curve1").Add oLineByPoints36
    oLineExtend37.Inputs("Curve2").Add oExtractPorts9.Output("MemberAxis", 1)
    oLineExtend37.Parameter("Context1") = 1
    oLineExtend37.Parameter("Context2") = 3
    oLineExtend37.Evaluate

    Dim oLineExtend38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend38 = oGCFactory.CreateEntity("LineExtend", pPOM, "038-LineExtend")
    oLineExtend38.Inputs("Line").Add oLineParallel20
    oLineExtend38.Inputs("Curve1").Add oLineByPoints35
    oLineExtend38.Inputs("Curve2").Add oExtractPorts9.Output("MemberAxis", 1)
    oLineExtend38.Parameter("Context1") = 1
    oLineExtend38.Parameter("Context2") = 3
    oLineExtend38.Evaluate

    Dim oCpxStringByCurves39 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves39 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "039-CpxStringByCurves")
    oCpxStringByCurves39.Inputs("Curves").Add oLineExtend38
    oCpxStringByCurves39.Inputs("Curves").Add oLineByPoints35
    oCpxStringByCurves39.Inputs("Curves").Add oLineByPoints31
    oCpxStringByCurves39.Evaluate

    Dim oCpxStringByCurves40 As SP3DGeometricConstruction.GeometricConstruction
    Set oCpxStringByCurves40 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "040-CpxStringByCurves")
    oCpxStringByCurves40.Inputs("Curves").Add oLineExtend37
    oCpxStringByCurves40.Inputs("Curves").Add oLineByPoints36
    oCpxStringByCurves40.Inputs("Curves").Add oLineByPoints30
    oCpxStringByCurves40.Evaluate

    Dim oSurfByLinearExtrusion41 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion41 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "041-SurfByLinearExtrusion")
    oSurfByLinearExtrusion41.Inputs("PlanarCrossSection").Add oCpxStringByCurves39
    oSurfByLinearExtrusion41.Inputs("ExtrusionLine").Add oLineFromCS16
    oSurfByLinearExtrusion41.Evaluate

    Dim oSurfByLinearExtrusion42 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion42 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "042-SurfByLinearExtrusion")
    oSurfByLinearExtrusion42.Inputs("PlanarCrossSection").Add oCpxStringByCurves40
    oSurfByLinearExtrusion42.Inputs("ExtrusionLine").Add oLineFromCS16
    oSurfByLinearExtrusion42.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion42.Output
    oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion41.Output
    oGCMacro.Output("Boundary", 3) = oSurfFromGType23.Output
    oGCMacro.Output("Boundary", "EdgePort") = oSurfByLinearExtrusion17.Output
    oGCMacro.Output("Support", 1) = oExtractPorts10.Output("Support", 1)
End Sub
