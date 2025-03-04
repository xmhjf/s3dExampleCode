VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertPlate"
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

Private dT As Double
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
Private Property Get METHOD(sMethod As String) As String

    Let METHOD = Source + "::" + sMethod

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
    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCConnectedElementSC_PostDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PostDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCConnectedElementSC_PreDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PreDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)

    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGeometricConstructionDefinitionService_Initialize"))

    Call m_oGenericPlateHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oGenericPlateHelper.PostInitialize(pGCDefinition)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)

    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGeometricConstructionDefinitionService_Evaluate"))

    Dim bEvaluateGeometry As Boolean
    Call m_oGenericPlateHelper.PreEvaluate(pGC, pPOM, bEvaluateGeometry)
    Call m_oAdvancedDebug.EnterMethod(METHOD("EvaluateGeometry"))
    If bEvaluateGeometry Then
        Call Evaluate(pGC, pPOM)
    Else
        Call m_oAdvancedDebug.ShowMsg("Skip")
    End If
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.PostEvaluate(pGC, pPOM)

    Dim oMemberPart As Object: Set oMemberPart = pGC.Inputs("PrimaryProfile1").Item(1)
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
    
        Dim dRealThickness As Double
        Dim dFlangeThickness As Double

        If pGC.ControlledInputs("AdvancedPlateSystem").Count > 0 Then
    
            dFlangeThickness = dT

            Dim oPlate As Object
            Set oPlate = pGC.ControlledInputs("AdvancedPlateSystem")(1)
            Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = oPlate

            Dim oPlateDimensions As IJDPlateDimensions
            Set oPlateDimensions = GetPlateDimension1(oPlate, dFlangeThickness, dRealThickness)

            Dim oPlateSys As IJPlate
            Dim bResetThickness As Double
            Set oPlateSys = oPlate

            bResetThickness = pGC.Parameter("MPlateThickness")
        
             If bResetThickness = 0.05 Then

                oPlateSys.Dimensions = oPlateDimensions
    
                pGC.Parameter("MPlateThickness") = 0.06
             End If

         Else
    
             dRealThickness = pGC.Parameter("MPlateThickness")
             
         End If

    Else
    
        Dim dRest As Double
        dRest = pGC.Parameter("MPlateThickness")

        If pGC.ControlledInputs("AdvancedPlateSystem").Count > 0 Then

            Dim oPlateER As Object
            Set oPlateER = pGC.ControlledInputs("AdvancedPlateSystem")(1)

            Dim oProfileSystem As Object

            Set oProfileSystem = pGC.Inputs("PrimaryProfile1")(1)
            Dim dWebThick As Double

            Dim oSDProfileSys As StructDetailObjects.ProfileSystem
            Set oSDProfileSys = New StructDetailObjects.ProfileSystem
            Set oSDProfileSys.object = oProfileSystem

            dWebThick = oSDProfileSys.WebThickness

'             Dim dRealThickness As Double
             Dim oPlateDimensionsER As IJDPlateDimensions
             Set oPlateDimensionsER = GetPlateDimension(oPlateER, dWebThick, dRealThickness)

             Dim oPlateERSys As IJPlate
             Set oPlateERSys = oPlateER

             oPlateERSys.Dimensions = oPlateDimensionsER

             pGC.Parameter("ResetThickness") = 0.02

         End If
         
    End If

'''    Dim oAPS As Object:
'''    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
'''        Set oAPS = pGC.ControlledInputs("AdvancedPlateSystem")(1)
'''        Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = oAPS
'''        pPlateMoldedConventions.plateThicknessDirection = WithNormalDir
'''    End If
    
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

    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCSemanticConnection_PostConnectionAdded"))

    Call m_oGenericPlateHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)

    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCSemanticConnection_PreConnectionRemoved"))

    Call m_oGenericPlateHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)

    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCMigrate_Migrate"))

    Call m_oGenericPlateHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMirror interface
'
Private Sub IJGCMirror_Adapt(ByVal pGCOfOriginalGC As IJGeometricConstruction, ByVal pGCOfMirroredGC As IJGeometricConstruction, _
                             ByVal pPlaneOfMirrorPlane As IJPlane, ByVal pT4x4OfMirrorTransformation As IJDT4x4, _
                             ByVal pElementsOfGCsFromInitalCopySet As IJElements)

    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCMirror_Adapt"))

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
    If sName <> "MPlateThickness" Then
        Call m_oGenericPlateHelper.IsOnRibbonBar(sName, pGC, vInfo, bIsOnRibbonBar)
    End If
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(METHOD("IJGCToDoDelegate_ToDoDelegate"))

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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select a Continuous Beam Member or ER", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersSplitByPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("SecondaryProfile1", "Select a Column or Brace", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfSecondaryProfile1", "BehaviorOfSecondaryProfile1", 1, 0, 0, 0, 0, "MembersBoundByPlate,MembersBoundPlate,MembersPenetrateByPlate", 0)
    'Added optional input3
    Call pGeometricConstructionDefinition.AddInput("SecondaryProfile2", "Select N Column or Brace", "ISPSMemberPartPrismatic or ISPSDesignedMember", 0, 100, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfSecondaryProfile2", "BehaviorOfSecondaryProfile2", 1, 0, 0, 0, 0, "MembersBoundByPlate,MembersBoundPlate,MembersPenetrateByPlate", 0)
    
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_4_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_11_1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort_2_1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort_15_1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_2_1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_88_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_89_1")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem_19_1", "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem_19_1", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddParameter("Lengthoffset", "Lengthoffset", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Widthoffset", "Widthoffset", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("offset", "offset", 8, 1, 59, 0, 0, 0.0001)
    Call pGeometricConstructionDefinition.AddParameter("slope", "slope", 8, 58, 0, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", 4, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Negative", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Positive", 2)
    Call pGeometricConstructionDefinition.AddParameter("MPlateThickness", "MPlateThickness", 8, 1, 59, 0, 0, 0.05, False)
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory

    Dim oMemberPart As Object: Set oMemberPart = pGeometricConstruction.Inputs("PrimaryProfile1").Item(1)
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then

        Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0001-LineAxisPortExtractor")
        oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
        oLineAxisPortExtractor1.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port_1_1").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")
    
        Dim oPointAtCurveMiddle85 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle85 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0085-PointAtCurveMiddle")
        oPointAtCurveMiddle85.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oPointAtCurveMiddle85.Evaluate
        
        Dim oCSFromMember86 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSFromMember86 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0086-CSFromMember")
        oCSFromMember86.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
        oCSFromMember86.Inputs("Point").Add oPointAtCurveMiddle85, "1"
        oCSFromMember86.Evaluate
        
        Dim oFacePortExtractor87 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor87 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0087-FacePortExtractor")
        oFacePortExtractor87.Inputs("Connectable").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "2"
        oFacePortExtractor87.Inputs("CoordinateSystem").Add oCSFromMember86, "1"
        oFacePortExtractor87.Parameter("GeometrySelector") = 4
        oFacePortExtractor87.Parameter("FacesContext") = 2
        oFacePortExtractor87.Parameter("LookingAxis") = 2
        oFacePortExtractor87.Parameter("IntersectingPlane") = 0
        oFacePortExtractor87.Parameter("SurfaceType") = 1
        oFacePortExtractor87.Parameter("TrackFlag") = 1
        oFacePortExtractor87.Parameter("Offset") = 0#
        oFacePortExtractor87.Evaluate
    
        pGeometricConstruction.ControlledInputs("FacePort3").Clear
        pGeometricConstruction.ControlledInputs("FacePort3").Add oFacePortExtractor87.ControlledInputs("Port").Item("1")
        Dim oPlaneFromCS88 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneFromCS88 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "0088-PlaneFromCS")
        oPlaneFromCS88.Inputs("CoordinateSystem").Add oCSFromMember86, "1"
        oPlaneFromCS88.Parameter("LookingAxis") = 1
        oPlaneFromCS88.Parameter("Offset") = 0.5
        oPlaneFromCS88.Parameter("Range") = 2#
        oPlaneFromCS88.Evaluate
        
        Dim oCurveByIntersection89 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection89 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0089-CurveByIntersection")
        oCurveByIntersection89.Inputs("Surface1").Add oPlaneFromCS88, "1"
        oCurveByIntersection89.Inputs("Surface2").Add oFacePortExtractor87, "1"
        oCurveByIntersection89.Parameter("TrackFlag") = 2
        oCurveByIntersection89.Parameter("ConstructionSurface") = 0
        oCurveByIntersection89.Evaluate
        
        Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0003-ParamDistConstant")
        oParamDistConstant3.Parameter("Value") = 0.1
        oParamDistConstant3.Evaluate
    
        Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0004-ParamDistConstant")
        oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Lengthoffset")
        oParamDistConstant4.Evaluate
    
        Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0005-ParamDistConstant")
        oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("Widthoffset")
        oParamDistConstant5.Evaluate
    
        Dim oParamDistConstant6 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant6 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0006-ParamDistConstant")
        oParamDistConstant6.Parameter("Value") = pGeometricConstruction.Parameter("offset")
        oParamDistConstant6.Evaluate
    
        Dim oParamDouble7 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDouble7 = oGCFactory.CreateEntity("ParamDouble", pPOM, "0007-ParamDouble")
        oParamDouble7.Parameter("Value") = pGeometricConstruction.Parameter("slope")
        oParamDouble7.Evaluate
    
        Dim oLineAxisPortExtractor8 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAxisPortExtractor8 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0008-LineAxisPortExtractor")
        oLineAxisPortExtractor8.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oLineAxisPortExtractor8.Evaluate
    
        pGeometricConstruction.ControlledInputs("AxisPort_2_1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort_2_1").Add oLineAxisPortExtractor8.ControlledInputs("Port").Item("1")
        Dim oParamDistMeasureLength9 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureLength9 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "0009-ParamDistMeasureLength")
        oParamDistMeasureLength9.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oParamDistMeasureLength9.Parameter("Value") = 7.34614622802165
        oParamDistMeasureLength9.Evaluate
    
        Dim oParamDistFunctionWith2Args10 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args10 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0010-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args10.Inputs("Arg1").Add oParamDistConstant3, "1"
        oParamDistFunctionWith2Args10.Inputs("Arg2").Add oParamDouble7, "1"
        oParamDistFunctionWith2Args10.Parameter("Value") = 0.1
        oParamDistFunctionWith2Args10.Parameter("Operation") = 3
        oParamDistFunctionWith2Args10.Evaluate
    
        Dim oPointAtCurveStart11 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart11 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0011-PointAtCurveStart")
        oPointAtCurveStart11.Inputs("Curve").Add oLineAxisPortExtractor8, "1"
        oPointAtCurveStart11.Evaluate
    
        Dim oLineFromCrossSection12 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection12 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0012-LineFromCrossSection")
        oLineFromCrossSection12.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oLineFromCrossSection12.Inputs("Point").Add oPointAtCurveStart11, "1"
        oLineFromCrossSection12.Parameter("EdgeSelector") = 514
        oLineFromCrossSection12.Evaluate
    
        Dim oLineFromCrossSection13 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection13 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0013-LineFromCrossSection")
        oLineFromCrossSection13.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oLineFromCrossSection13.Inputs("Point").Add oPointAtCurveStart11, "1"
        oLineFromCrossSection13.Parameter("EdgeSelector") = 258
        oLineFromCrossSection13.Evaluate
    
        Dim oPointByCurves14 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves14 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0014-PointByCurves")
        oPointByCurves14.Inputs("Curve1").Add oLineFromCrossSection12, "1"
        oPointByCurves14.Inputs("Curve2").Add oLineFromCrossSection13, "1"
        oPointByCurves14.Parameter("TrackFlag") = 1
        oPointByCurves14.Evaluate
    
        Dim oPointAtCurveMiddle15 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle15 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0015-PointAtCurveMiddle")
        oPointAtCurveMiddle15.Inputs("Curve").Add oLineFromCrossSection13, "1"
        oPointAtCurveMiddle15.Evaluate
    
        Dim oPointAlongCurve16 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve16 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0016-PointAlongCurve")
        oPointAlongCurve16.Inputs("Curve").Add oLineFromCrossSection13, "1"
        oPointAlongCurve16.Inputs("Point").Add oPointByCurves14, "1"
        oPointAlongCurve16.Inputs("TrackPoint").Add oPointAtCurveMiddle15, "1"
        oPointAlongCurve16.Parameter("Distance") = 1#
        oPointAlongCurve16.Parameter("TrackFlag") = 2
        oPointAlongCurve16.Evaluate
    
        Dim oLineByPoints17 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints17 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0017-LineByPoints")
        oLineByPoints17.Inputs("StartPoint").Add oPointByCurves14, "1"
        oLineByPoints17.Inputs("EndPoint").Add oPointAlongCurve16, "1"
        oLineByPoints17.Evaluate
    
        Dim oExtractPorts18 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts18 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0018-ExtractPorts")
        oExtractPorts18.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oExtractPorts18.Inputs("ZAxis").Add oLineByPoints17, "2"
        oExtractPorts18.Parameter("Support") = pGeometricConstruction.Parameter("Location")
        oExtractPorts18.Parameter("Debug") = 0
        oExtractPorts18.Evaluate
    
        pGeometricConstruction.ControlledInputs("AxisPort_2_1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort_2_1").Add oExtractPorts18.ControlledInputs("AxisPort").Item("1")
        pGeometricConstruction.ControlledInputs("FacePort_2_1").Clear
        pGeometricConstruction.ControlledInputs("FacePort_2_1").Add oExtractPorts18.ControlledInputs("FacePort").Item("1")
        Dim oParamDistMeasureParallel19 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel19 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0019-ParamDistMeasureParallel")
        oParamDistMeasureParallel19.Inputs("Graphics").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oParamDistMeasureParallel19.Inputs("Graphics").Add oExtractPorts18.Output("CurveRight", 1), "2"
        oParamDistMeasureParallel19.Parameter("Value") = 0.203199999999999
        oParamDistMeasureParallel19.Evaluate
    
        Dim oPointAtCurveStart20 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart20 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0020-PointAtCurveStart")
        oPointAtCurveStart20.Inputs("Curve").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointAtCurveStart20.Evaluate
    
        Dim oParamDistFunctionWith2Args21 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args21 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0021-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args21.Inputs("Arg1").Add oParamDistMeasureParallel19, "1"
        oParamDistFunctionWith2Args21.Inputs("Arg2").Add oParamDistConstant6, "1"
        oParamDistFunctionWith2Args21.Parameter("Value") = 0.203299999999999
        oParamDistFunctionWith2Args21.Parameter("Operation") = 1
        oParamDistFunctionWith2Args21.Evaluate
        
        Dim oPointAtCurveExtremity90 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity90 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0090-PointAtCurveExtremity")
        oPointAtCurveExtremity90.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oPointAtCurveExtremity90.Inputs("TrackPoint").Add oPointByCurves14, "1"
        oPointAtCurveExtremity90.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity90.Evaluate
    
        Dim oPointAtCurveExtremity91 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity91 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0091-PointAtCurveExtremity")
        oPointAtCurveExtremity91.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oPointAtCurveExtremity91.Inputs("TrackPoint").Add oPointByCurves14, "1"
        oPointAtCurveExtremity91.Parameter("TrackFlag") = 2
        oPointAtCurveExtremity91.Evaluate
            
        Dim oVectorTangentToCurve92 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve92 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0092-VectorTangentToCurve")
        oVectorTangentToCurve92.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oVectorTangentToCurve92.Inputs("Point").Add oPointAtCurveExtremity90, "1"
        oVectorTangentToCurve92.Inputs("TrackPoint").Add oPointAtCurveExtremity91, "1"
        oVectorTangentToCurve92.Parameter("Range") = 5#
        oVectorTangentToCurve92.Parameter("Orientation") = 1
        oVectorTangentToCurve92.Parameter("TrackFlag") = 2
        oVectorTangentToCurve92.Evaluate
    
        Dim oPointAtCurveExtremity93 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity93 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0093-PointAtCurveExtremity")
        oPointAtCurveExtremity93.Inputs("Curve").Add oVectorTangentToCurve92, "1"
        oPointAtCurveExtremity93.Inputs("TrackPoint").Add oPointAtCurveExtremity90, "1"
        oPointAtCurveExtremity93.Parameter("TrackFlag") = 2
        oPointAtCurveExtremity93.Evaluate
    
        Dim oLineByPoints94 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints94 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0094-LineByPoints")
        oLineByPoints94.Inputs("StartPoint").Add oPointAtCurveExtremity93, "1"
        oLineByPoints94.Inputs("EndPoint").Add oPointAtCurveExtremity91, "1"
        oLineByPoints94.Evaluate
    
        Dim oPointByCurves22 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves22 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0022-PointByCurves")
        oPointByCurves22.Inputs("Curve1").Add oLineByPoints94, "1"
        oPointByCurves22.Inputs("Curve2").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointByCurves22.Parameter("TrackFlag") = 1
        oPointByCurves22.Evaluate
    
        Dim oCSFromMember23 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSFromMember23 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0023-CSFromMember")
        oCSFromMember23.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oCSFromMember23.Inputs("Point").Add oPointByCurves22, "1"
        oCSFromMember23.Evaluate
    
        Dim oLineFromCrossSection24 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCrossSection24 = oGCFactory.CreateEntity("LineFromCrossSection", pPOM, "0024-LineFromCrossSection")
        oLineFromCrossSection24.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oLineFromCrossSection24.Inputs("Point").Add oPointByCurves22, "1"
        oLineFromCrossSection24.Parameter("EdgeSelector") = 514
        oLineFromCrossSection24.Evaluate
    
        Dim oPlaneByPointNormal25 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal25 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0025-PlaneByPointNormal")
        oPlaneByPointNormal25.Inputs("Point").Add oPointByCurves22, "1"
        oPlaneByPointNormal25.Inputs("Line").Add oLineFromCrossSection24, "1"
        oPlaneByPointNormal25.Parameter("Range") = 10# 'oParamDistMeasureLength9
        oPlaneByPointNormal25.Evaluate
    
        Dim oPointAtCurveStart26 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart26 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0026-PointAtCurveStart")
        oPointAtCurveStart26.Inputs("Curve").Add oLineFromCrossSection24, "1"
        oPointAtCurveStart26.Evaluate
    
        Dim oLineFromCS27 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS27 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0027-LineFromCS")
        oLineFromCS27.Inputs("CoordinateSystem").Add oCSFromMember23, "1"
        oLineFromCS27.Parameter("LookingAxis") = 3
        oLineFromCS27.Parameter("Length") = -5#
        oLineFromCS27.Parameter("CSOrientation") = 1
        oLineFromCS27.Parameter("LineJustification") = 1
        oLineFromCS27.Evaluate
    
        Dim oFacePortExtractor28 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor28 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0028-FacePortExtractor")
        oFacePortExtractor28.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oFacePortExtractor28.Inputs("CoordinateSystem").Add oCSFromMember23, "1"
        oFacePortExtractor28.Parameter("GeometrySelector") = 2
        oFacePortExtractor28.Parameter("FacesContext") = 2
        oFacePortExtractor28.Parameter("LookingAxis") = 3
        oFacePortExtractor28.Parameter("IntersectingPlane") = 0
        oFacePortExtractor28.Parameter("SurfaceType") = 1
        oFacePortExtractor28.Parameter("TrackFlag") = 1
        oFacePortExtractor28.Parameter("Offset") = 0#
        oFacePortExtractor28.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_88_1").Clear
        pGeometricConstruction.ControlledInputs("Port_88_1").Add oFacePortExtractor28.ControlledInputs("Port").Item("1")
        Dim oFacePortExtractor29 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor29 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0029-FacePortExtractor")
        oFacePortExtractor29.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oFacePortExtractor29.Inputs("CoordinateSystem").Add oCSFromMember23, "1"
        oFacePortExtractor29.Parameter("GeometrySelector") = 2
        oFacePortExtractor29.Parameter("FacesContext") = 4
        oFacePortExtractor29.Parameter("LookingAxis") = 3
        oFacePortExtractor29.Parameter("IntersectingPlane") = 0
        oFacePortExtractor29.Parameter("SurfaceType") = 1
        oFacePortExtractor29.Parameter("TrackFlag") = 1
        oFacePortExtractor29.Parameter("Offset") = 0#
        oFacePortExtractor29.Evaluate
        pGeometricConstruction.ControlledInputs("Port_89_1").Clear
        pGeometricConstruction.ControlledInputs("Port_89_1").Add oFacePortExtractor29.ControlledInputs("Port").Item("1")
        
        Dim oMember As Object
        Set oMember = pGeometricConstruction.Inputs("PrimaryProfile1")(1)
        Dim oSD_Member As New StructDetailObjects.MemberPart
        Set oSD_Member.object = oMember
        If oSD_Member.SectionType = "S" Then
        
            Dim oFacePortExtractor95 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortExtractor95 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0095-FacePortExtractor")
            oFacePortExtractor95.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
            oFacePortExtractor95.Inputs("CoordinateSystem").Add oCSFromMember23, "1"
            oFacePortExtractor95.Parameter("GeometrySelector") = 4
            oFacePortExtractor95.Parameter("FacesContext") = 2
            oFacePortExtractor95.Parameter("LookingAxis") = 2
            oFacePortExtractor95.Parameter("IntersectingPlane") = 0
            oFacePortExtractor95.Parameter("SurfaceType") = 1
            oFacePortExtractor95.Parameter("TrackFlag") = 1
            oFacePortExtractor95.Parameter("Offset") = 0#
            oFacePortExtractor95.Evaluate
            
            Dim oCurveByIntersection96 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection96 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0096-CurveByIntersection")
            oCurveByIntersection96.Inputs("Surface1").Add oFacePortExtractor29, "2"
            oCurveByIntersection96.Inputs("Surface2").Add oFacePortExtractor95, "1"
            oCurveByIntersection96.Parameter("TrackFlag") = 1
            oCurveByIntersection96.Parameter("ConstructionSurface") = 0
            oCurveByIntersection96.Evaluate
        
            Dim oCurveByProjection97 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByProjection97 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0097-CurveByProjection")
            oCurveByProjection97.Inputs("Curve").Add oCurveByIntersection96, "1"
            oCurveByProjection97.Inputs("Surface").Add oFacePortExtractor28, "1"
            oCurveByProjection97.Inputs("Line").Add oLineFromCrossSection13, "1"
            oCurveByProjection97.Parameter("TrackFlag") = 1
            oCurveByProjection97.Parameter("ConstructionSurface") = 0
            oCurveByProjection97.Evaluate
        
            Dim oParamDistMeasureParallel30 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel30 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0030-ParamDistMeasureParallel")
            oParamDistMeasureParallel30.Inputs("Graphics").Add oCurveByIntersection96, "1"
            oParamDistMeasureParallel30.Inputs("Graphics").Add oCurveByProjection97, "2"
            oParamDistMeasureParallel30.Parameter("Value") = 3.53569989408697E-02
            oParamDistMeasureParallel30.Evaluate
    
        Else
        
            Set oParamDistMeasureParallel30 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0030-ParamDistMeasureParallel")
            oParamDistMeasureParallel30.Inputs("Graphics").Add oFacePortExtractor28, "1"
            oParamDistMeasureParallel30.Inputs("Graphics").Add oFacePortExtractor29, "2"
            oParamDistMeasureParallel30.Parameter("Value") = 4.49580000000012E-02
            oParamDistMeasureParallel30.Evaluate
        
        End If
    
        Dim oPlaneByPointNormal31 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal31 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0031-PlaneByPointNormal")
        oPlaneByPointNormal31.Inputs("Point").Add oPointByCurves22, "2"
        oPlaneByPointNormal31.Inputs("Line").Add oLineFromCrossSection13, "1" 'oLineAxisPortExtractor1, "2"
        oPlaneByPointNormal31.Parameter("Range") = 10#
        oPlaneByPointNormal31.Evaluate
    
        Dim oLineParallel32 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel32 = oGCFactory.CreateEntity("LineParallel", pPOM, "0032-LineParallel")
        oLineParallel32.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oLineParallel32.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineParallel32.Inputs("TrackPoint").Add oPointAtCurveStart26, "3"
        oLineParallel32.Parameter("Distance") = oParamDistFunctionWith2Args21
        oLineParallel32.Parameter("TrackFlag") = 1
        oLineParallel32.Evaluate
    
        Dim oLineParallel33 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel33 = oGCFactory.CreateEntity("LineParallel", pPOM, "0033-LineParallel")
        oLineParallel33.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oLineParallel33.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineParallel33.Inputs("TrackPoint").Add oPointAtCurveStart26, "1"
        oLineParallel33.Parameter("Distance") = oParamDistFunctionWith2Args21
        oLineParallel33.Parameter("TrackFlag") = 2
        oLineParallel33.Evaluate
        
        Dim Value1
        Value1 = pGeometricConstruction.Inputs("SecondaryProfile2").Count
        If Value1 > 0 Then  'Input3 has at least one input
        
            Dim oLineAxisPortExtractor98 As SP3DGeometricConstruction.GeometricConstruction
            Set oLineAxisPortExtractor98 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0098-LineAxisPortExtractor")
            oLineAxisPortExtractor98.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "1"
            oLineAxisPortExtractor98.Evaluate

            Dim oPointAtCurveMiddle99 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointAtCurveMiddle99 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0099-PointAtCurveMiddle")
            oPointAtCurveMiddle99.Inputs("Curve").Add oLineAxisPortExtractor98, "1"
            oPointAtCurveMiddle99.Evaluate
            
            Dim oCSFromMember100 As SP3DGeometricConstruction.GeometricConstruction
            Set oCSFromMember100 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0100-CSFromMember")
            oCSFromMember100.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "1"
            oCSFromMember100.Inputs("Point").Add oPointAtCurveMiddle99, "1"
            oCSFromMember100.Evaluate
            
            Dim oFacePortExtractor101 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortExtractor101 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0101-FacePortExtractor")
            oFacePortExtractor101.Inputs("Connectable").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "2"
            oFacePortExtractor101.Inputs("CoordinateSystem").Add oCSFromMember100, "1"
            oFacePortExtractor101.Parameter("GeometrySelector") = 4
            oFacePortExtractor101.Parameter("FacesContext") = 2
            oFacePortExtractor101.Parameter("LookingAxis") = 2
            oFacePortExtractor101.Parameter("IntersectingPlane") = 0
            oFacePortExtractor101.Parameter("SurfaceType") = 1
            oFacePortExtractor101.Parameter("TrackFlag") = 1
            oFacePortExtractor101.Parameter("Offset") = 0#
            oFacePortExtractor101.Evaluate
            
            Dim oLineByPoints103 As SP3DGeometricConstruction.GeometricConstruction
            Set oLineByPoints103 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0103-LineByPoints")
            oLineByPoints103.Inputs("StartPoint").Add oPointAtCurveMiddle85, "1"
            oLineByPoints103.Inputs("EndPoint").Add oPointAtCurveMiddle99, "1"
            oLineByPoints103.Evaluate
            
            Dim oPointAtCurveMiddle104 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointAtCurveMiddle104 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0104-PointAtCurveMiddle")
            oPointAtCurveMiddle104.Inputs("Curve").Add oLineByPoints103, "1"
            oPointAtCurveMiddle104.Evaluate
            
            Dim oPointAtCurveExtremity102 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointAtCurveExtremity102 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0102-PointAtCurveExtremity")
            oPointAtCurveExtremity102.Inputs("Curve").Add oLineAxisPortExtractor98, "1"
            oPointAtCurveExtremity102.Inputs("TrackPoint").Add oPointByCurves22, "1"
            oPointAtCurveExtremity102.Parameter("TrackFlag") = 2
            oPointAtCurveExtremity102.Evaluate
        
            Dim oCurveByIntersection34 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection34 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0034-CurveByIntersection")
            oCurveByIntersection34.Inputs("Surface1").Add oPlaneByPointNormal25, "1"
            oCurveByIntersection34.Inputs("Surface2").Add oFacePortExtractor87, "2"
            oCurveByIntersection34.Inputs("TrackPoint").Add oPointAtCurveMiddle104, "1"
            oCurveByIntersection34.Parameter("TrackFlag") = 2
            oCurveByIntersection34.Parameter("ConstructionSurface") = 0
            oCurveByIntersection34.Evaluate
            
            Dim oCurveByIntersection35 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection35 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0035-CurveByIntersection")
            oCurveByIntersection35.Inputs("Surface1").Add oPlaneByPointNormal25, "1"
            oCurveByIntersection35.Inputs("Surface2").Add oFacePortExtractor101, "2"
            oCurveByIntersection35.Inputs("TrackPoint").Add oPointAtCurveMiddle104, "1"
            oCurveByIntersection35.Parameter("TrackFlag") = 2
            oCurveByIntersection35.Parameter("ConstructionSurface") = 0
            oCurveByIntersection35.Evaluate
                
        Else 'Input3 no input.
    
            Set oCurveByIntersection34 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0034-CurveByIntersection")
            oCurveByIntersection34.Inputs("Surface1").Add oPlaneByPointNormal25, "1"
            oCurveByIntersection34.Inputs("Surface2").Add oFacePortExtractor87, "2"
            oCurveByIntersection34.Inputs("TrackPoint").Add oPointAtCurveStart20, "1"
            oCurveByIntersection34.Parameter("TrackFlag") = 1
            oCurveByIntersection34.Parameter("ConstructionSurface") = 0
            oCurveByIntersection34.Evaluate

            Set oCurveByIntersection35 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0035-CurveByIntersection")
            oCurveByIntersection35.Inputs("Surface1").Add oPlaneByPointNormal25, "1"
            oCurveByIntersection35.Inputs("Surface2").Add oFacePortExtractor87, "2"
            oCurveByIntersection35.Inputs("TrackPoint").Add oPointAtCurveStart20, "1"
            oCurveByIntersection35.Parameter("TrackFlag") = 2
            oCurveByIntersection35.Parameter("ConstructionSurface") = 0
            oCurveByIntersection35.Evaluate
        
        End If
    
        Dim oPointByCurves36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves36 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0036-PointByCurves")
        oPointByCurves36.Inputs("Curve1").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointByCurves36.Inputs("Curve2").Add oCurveByIntersection34, "1"
        oPointByCurves36.Parameter("TrackFlag") = 1
        oPointByCurves36.Evaluate
        
        Dim oPointByCurves37 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves37 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0037-PointByCurves")
        oPointByCurves37.Inputs("Curve1").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointByCurves37.Inputs("Curve2").Add oCurveByIntersection35, "1"
        oPointByCurves37.Parameter("TrackFlag") = 1
        oPointByCurves37.Evaluate
    
        Dim oCurveByIntersection38 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection38 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0038-CurveByIntersection")
        oCurveByIntersection38.Inputs("Surface1").Add oPlaneByPointNormal25, "1"
        oCurveByIntersection38.Inputs("Surface2").Add oFacePortExtractor87, "2"
        oCurveByIntersection38.Inputs("TrackPoint").Add oPointAtCurveStart20, "1"
        oCurveByIntersection38.Parameter("TrackFlag") = 2
        oCurveByIntersection38.Parameter("ConstructionSurface") = 0
        oCurveByIntersection38.Evaluate
    
        Dim oParamDistMeasureParallel39 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel39 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0039-ParamDistMeasureParallel")
        oParamDistMeasureParallel39.Inputs("Graphics").Add oLineAxisPortExtractor1, "1"
        oParamDistMeasureParallel39.Inputs("Graphics").Add oCurveByIntersection38, "2"
        oParamDistMeasureParallel39.Parameter("Value") = 0.5
        oParamDistMeasureParallel39.Evaluate
    
        Dim oPointAlongCurve40 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve40 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0040-PointAlongCurve")
        oPointAlongCurve40.Inputs("Curve").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointAlongCurve40.Inputs("Point").Add oPointByCurves37, "1"
        oPointAlongCurve40.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve40.Parameter("Distance") = oParamDistConstant4
        oPointAlongCurve40.Parameter("TrackFlag") = 2
        oPointAlongCurve40.Evaluate
    
        Dim oPointAlongCurve41 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve41 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0041-PointAlongCurve")
        oPointAlongCurve41.Inputs("Curve").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPointAlongCurve41.Inputs("Point").Add oPointByCurves36, "2"
        oPointAlongCurve41.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve41.Parameter("Distance") = oParamDistConstant4
        oPointAlongCurve41.Parameter("TrackFlag") = 2
        oPointAlongCurve41.Evaluate
    
        Dim oPlaneByPointNormal42 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal42 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0042-PlaneByPointNormal")
        oPlaneByPointNormal42.Inputs("Point").Add oPointAlongCurve40, "1"
        oPlaneByPointNormal42.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal42.Parameter("Range") = 5#
        oPlaneByPointNormal42.Evaluate
    
        Dim oPlaneByPointNormal43 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal43 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0043-PlaneByPointNormal")
        oPlaneByPointNormal43.Inputs("Point").Add oPointAlongCurve41, "2"
        oPlaneByPointNormal43.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal43.Parameter("Range") = 5#
        oPlaneByPointNormal43.Evaluate
    
        Dim oCurveByIntersection44 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection44 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0044-CurveByIntersection")
        oCurveByIntersection44.Inputs("Surface1").Add oPlaneByPointNormal42, "1"
        oCurveByIntersection44.Inputs("Surface2").Add oPlaneByPointNormal31, "2"
        oCurveByIntersection44.Parameter("TrackFlag") = 2
        oCurveByIntersection44.Parameter("ConstructionSurface") = 0
        oCurveByIntersection44.Evaluate
    
        Dim oCurveByIntersection45 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection45 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0045-CurveByIntersection")
        oCurveByIntersection45.Inputs("Surface1").Add oPlaneByPointNormal43, "2"
        oCurveByIntersection45.Inputs("Surface2").Add oPlaneByPointNormal31, "1"
        oCurveByIntersection45.Parameter("TrackFlag") = 2
        oCurveByIntersection45.Parameter("ConstructionSurface") = 0
        oCurveByIntersection45.Evaluate
    
        Dim oSurfFromGType46 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType46 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0046-SurfFromGType")
        oSurfFromGType46.Inputs("Surface").Add oPlaneByPointNormal42, "1"
        oSurfFromGType46.Evaluate
    
        Dim oSurfFromGType47 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType47 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0047-SurfFromGType")
        oSurfFromGType47.Inputs("Surface").Add oPlaneByPointNormal43, "2"
        oSurfFromGType47.Evaluate
    
        Dim oParamDistFunctionWith2Args48 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args48 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0048-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args48.Inputs("Arg1").Add oParamDistMeasureParallel39, "1"
        oParamDistFunctionWith2Args48.Inputs("Arg2").Add oParamDistConstant5, "1"
        oParamDistFunctionWith2Args48.Parameter("Value") = 0.6
        oParamDistFunctionWith2Args48.Parameter("Operation") = 1
        oParamDistFunctionWith2Args48.Evaluate
    
        Dim oLineParallel49 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel49 = oGCFactory.CreateEntity("LineParallel", pPOM, "0049-LineParallel")
        oLineParallel49.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oLineParallel49.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineParallel49.Inputs("TrackPoint").Add oPointAtCurveStart26, "1"
        oLineParallel49.Parameter("Distance") = oParamDistFunctionWith2Args48
        oLineParallel49.Parameter("TrackFlag") = 1
        oLineParallel49.Evaluate
    
        Dim oLineParallel50 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel50 = oGCFactory.CreateEntity("LineParallel", pPOM, "0050-LineParallel")
        oLineParallel50.Inputs("Line").Add oExtractPorts18.Output("MemberAxis", 1), "1"
        oLineParallel50.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineParallel50.Inputs("TrackPoint").Add oPointAtCurveStart26, "1"
        oLineParallel50.Parameter("Distance") = oParamDistFunctionWith2Args48
        oLineParallel50.Parameter("TrackFlag") = 2
        oLineParallel50.Evaluate
    
        Dim oPointByCurves51 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves51 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0051-PointByCurves")
        oPointByCurves51.Inputs("Curve1").Add oLineParallel32, "1"
        oPointByCurves51.Inputs("Curve2").Add oCurveByIntersection44, "1"
        oPointByCurves51.Parameter("TrackFlag") = 1
        oPointByCurves51.Evaluate
    
        Dim oPointByCurves52 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves52 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0052-PointByCurves")
        oPointByCurves52.Inputs("Curve1").Add oLineParallel32, "1"
        oPointByCurves52.Inputs("Curve2").Add oCurveByIntersection45, "1"
        oPointByCurves52.Parameter("TrackFlag") = 1
        oPointByCurves52.Evaluate
    
        Dim oPointByCurves53 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves53 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0053-PointByCurves")
        oPointByCurves53.Inputs("Curve1").Add oLineParallel33, "2"
        oPointByCurves53.Inputs("Curve2").Add oCurveByIntersection44, "1"
        oPointByCurves53.Parameter("TrackFlag") = 1
        oPointByCurves53.Evaluate
    
        Dim oPointByCurves54 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves54 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0054-PointByCurves")
        oPointByCurves54.Inputs("Curve1").Add oLineParallel33, "1"
        oPointByCurves54.Inputs("Curve2").Add oCurveByIntersection45, "1"
        oPointByCurves54.Parameter("TrackFlag") = 1
        oPointByCurves54.Evaluate
    
        Dim oPointAlongCurve55 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve55 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0055-PointAlongCurve")
        oPointAlongCurve55.Inputs("Curve").Add oLineParallel32, "1"
        oPointAlongCurve55.Inputs("Point").Add oPointByCurves51, "1"
        oPointAlongCurve55.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve55.Parameter("Distance") = oParamDistFunctionWith2Args10
        oPointAlongCurve55.Parameter("TrackFlag") = 1
        oPointAlongCurve55.Evaluate
    
        Dim oPointAlongCurve56 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve56 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0056-PointAlongCurve")
        oPointAlongCurve56.Inputs("Curve").Add oLineParallel32, "1"
        oPointAlongCurve56.Inputs("Point").Add oPointByCurves52, "2"
        oPointAlongCurve56.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve56.Parameter("Distance") = oParamDistFunctionWith2Args10
        oPointAlongCurve56.Parameter("TrackFlag") = 1
        oPointAlongCurve56.Evaluate
    
        Dim oPointAlongCurve57 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve57 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0057-PointAlongCurve")
        oPointAlongCurve57.Inputs("Curve").Add oLineParallel33, "2"
        oPointAlongCurve57.Inputs("Point").Add oPointByCurves53, "2"
        oPointAlongCurve57.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve57.Parameter("Distance") = oParamDistFunctionWith2Args10
        oPointAlongCurve57.Parameter("TrackFlag") = 1
        oPointAlongCurve57.Evaluate
    
        Dim oPointAlongCurve58 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve58 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0058-PointAlongCurve")
        oPointAlongCurve58.Inputs("Curve").Add oLineParallel33, "1"
        oPointAlongCurve58.Inputs("Point").Add oPointByCurves54, "2"
        oPointAlongCurve58.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve58.Parameter("Distance") = oParamDistFunctionWith2Args10
        oPointAlongCurve58.Parameter("TrackFlag") = 1
        oPointAlongCurve58.Evaluate
    
        Dim oLineAtAngle59 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle59 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0059-LineAtAngle")
        oLineAtAngle59.Inputs("Line").Add oLineParallel32, "1"
        oLineAtAngle59.Inputs("Point").Add oPointAlongCurve55, "1"
        oLineAtAngle59.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineAtAngle59.Parameter("Angle") = 1.5707963267949
        oLineAtAngle59.Parameter("TrackFlag") = 1
        oLineAtAngle59.Evaluate
    
        Dim oLineAtAngle60 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle60 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0060-LineAtAngle")
        oLineAtAngle60.Inputs("Line").Add oLineParallel32, "1"
        oLineAtAngle60.Inputs("Point").Add oPointAlongCurve56, "1"
        oLineAtAngle60.Inputs("Plane").Add oExtractPorts18.Output("Support", 1), "1"
        oLineAtAngle60.Parameter("Angle") = 1.5707963267949
        oLineAtAngle60.Parameter("TrackFlag") = 1
        oLineAtAngle60.Evaluate
    
        Dim oPointAlongCurve61 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve61 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0061-PointAlongCurve")
        oPointAlongCurve61.Inputs("Curve").Add oLineAtAngle59, "1"
        oPointAlongCurve61.Inputs("Point").Add oPointAlongCurve55, "1"
        oPointAlongCurve61.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve61.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve61.Parameter("TrackFlag") = 2
        oPointAlongCurve61.Evaluate
    
        Dim oPointAlongCurve62 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve62 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0062-PointAlongCurve")
        oPointAlongCurve62.Inputs("Curve").Add oLineAtAngle59, "1"
        oPointAlongCurve62.Inputs("Point").Add oPointAlongCurve57, "2"
        oPointAlongCurve62.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve62.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve62.Parameter("TrackFlag") = 2
        oPointAlongCurve62.Evaluate
    
        Dim oPointAlongCurve63 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve63 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0063-PointAlongCurve")
        oPointAlongCurve63.Inputs("Curve").Add oLineAtAngle60, "2"
        oPointAlongCurve63.Inputs("Point").Add oPointAlongCurve56, "2"
        oPointAlongCurve63.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve63.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve63.Parameter("TrackFlag") = 2
        oPointAlongCurve63.Evaluate
    
        Dim oPointAlongCurve64 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve64 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0064-PointAlongCurve")
        oPointAlongCurve64.Inputs("Curve").Add oLineAtAngle60, "1"
        oPointAlongCurve64.Inputs("Point").Add oPointAlongCurve58, "2"
        oPointAlongCurve64.Inputs("TrackPoint").Add oPointByCurves22, "1"
        oPointAlongCurve64.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve64.Parameter("TrackFlag") = 2
        oPointAlongCurve64.Evaluate
    
        Dim oLineByPoints65 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints65 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0065-LineByPoints")
        oLineByPoints65.Inputs("StartPoint").Add oPointAlongCurve61, "2"
        oLineByPoints65.Inputs("EndPoint").Add oPointByCurves51, "2"
        oLineByPoints65.Evaluate
    
        Dim oLineByPoints66 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints66 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0066-LineByPoints")
        oLineByPoints66.Inputs("StartPoint").Add oPointAlongCurve62, "1"
        oLineByPoints66.Inputs("EndPoint").Add oPointByCurves53, "1"
        oLineByPoints66.Evaluate
    
        Dim oLineByPoints67 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints67 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0067-LineByPoints")
        oLineByPoints67.Inputs("StartPoint").Add oPointAlongCurve63, "3"
        oLineByPoints67.Inputs("EndPoint").Add oPointByCurves52, "1"
        oLineByPoints67.Evaluate
    
        Dim oLineByPoints68 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints68 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0068-LineByPoints")
        oLineByPoints68.Inputs("StartPoint").Add oPointAlongCurve64, "2"
        oLineByPoints68.Inputs("EndPoint").Add oPointByCurves54, "1"
        oLineByPoints68.Evaluate
    
        Dim oLineExtend69 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend69 = oGCFactory.CreateEntity("LineExtend", pPOM, "0069-LineExtend")
        oLineExtend69.Inputs("Line").Add oLineByPoints66, "1"
        oLineExtend69.Inputs("Curve1").Add oLineParallel50, "1"
        oLineExtend69.Parameter("Context1") = 1
        oLineExtend69.Parameter("Context2") = 1
        oLineExtend69.Evaluate
    
        Dim oLineExtend70 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend70 = oGCFactory.CreateEntity("LineExtend", pPOM, "0070-LineExtend")
        oLineExtend70.Inputs("Line").Add oLineByPoints68, "1"
        oLineExtend70.Inputs("Curve1").Add oLineParallel50, "1"
        oLineExtend70.Parameter("Context1") = 1
        oLineExtend70.Parameter("Context2") = 1
        oLineExtend70.Evaluate
    
        Dim oLineExtend71 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend71 = oGCFactory.CreateEntity("LineExtend", pPOM, "0071-LineExtend")
        oLineExtend71.Inputs("Line").Add oLineByPoints65, "1"
        oLineExtend71.Inputs("Curve1").Add oLineParallel49, "1"
        oLineExtend71.Parameter("Context1") = 1
        oLineExtend71.Parameter("Context2") = 1
        oLineExtend71.Evaluate
    
        Dim oLineExtend72 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend72 = oGCFactory.CreateEntity("LineExtend", pPOM, "0072-LineExtend")
        oLineExtend72.Inputs("Line").Add oLineByPoints67, "1"
        oLineExtend72.Inputs("Curve1").Add oLineParallel49, "1"
        oLineExtend72.Parameter("Context1") = 1
        oLineExtend72.Parameter("Context2") = 1
        oLineExtend72.Evaluate
    
        Dim oPointByCurves73 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves73 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0073-PointByCurves")
        oPointByCurves73.Inputs("Curve1").Add oLineParallel49, "1"
        oPointByCurves73.Inputs("Curve2").Add oLineExtend71, "1"
        oPointByCurves73.Parameter("TrackFlag") = 1
        oPointByCurves73.Evaluate
    
        Dim oPointByCurves74 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves74 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0074-PointByCurves")
        oPointByCurves74.Inputs("Curve1").Add oLineParallel49, "1"
        oPointByCurves74.Inputs("Curve2").Add oLineExtend72, "1"
        oPointByCurves74.Parameter("TrackFlag") = 1
        oPointByCurves74.Evaluate
    
        Dim oPointAlongCurve75 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve75 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0075-PointAlongCurve")
        oPointAlongCurve75.Inputs("Curve").Add oLineByPoints65, "1"
        oPointAlongCurve75.Inputs("Point").Add oPointByCurves51, "1"
        oPointAlongCurve75.Inputs("TrackPoint").Add oPointAlongCurve61, "2"
        oPointAlongCurve75.Parameter("Distance") = 0.05
        oPointAlongCurve75.Parameter("TrackFlag") = 2
        oPointAlongCurve75.Evaluate
    
        Dim oPointAlongCurve76 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve76 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0076-PointAlongCurve")
        oPointAlongCurve76.Inputs("Curve").Add oLineByPoints66, "2"
        oPointAlongCurve76.Inputs("Point").Add oPointByCurves53, "2"
        oPointAlongCurve76.Inputs("TrackPoint").Add oPointAlongCurve62, "2"
        oPointAlongCurve76.Parameter("Distance") = 0.05
        oPointAlongCurve76.Parameter("TrackFlag") = 2
        oPointAlongCurve76.Evaluate
    
        Dim oPointAlongCurve77 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve77 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0077-PointAlongCurve")
        oPointAlongCurve77.Inputs("Curve").Add oLineByPoints67, "1"
        oPointAlongCurve77.Inputs("Point").Add oPointByCurves52, "1"
        oPointAlongCurve77.Inputs("TrackPoint").Add oPointAlongCurve63, "1"
        oPointAlongCurve77.Parameter("Distance") = 0.05
        oPointAlongCurve77.Parameter("TrackFlag") = 2
        oPointAlongCurve77.Evaluate
    
        Dim oPointAlongCurve78 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve78 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0078-PointAlongCurve")
        oPointAlongCurve78.Inputs("Curve").Add oLineByPoints68, "2"
        oPointAlongCurve78.Inputs("Point").Add oPointByCurves54, "2"
        oPointAlongCurve78.Inputs("TrackPoint").Add oPointAlongCurve64, "2"
        oPointAlongCurve78.Parameter("Distance") = 0.05
        oPointAlongCurve78.Parameter("TrackFlag") = 2
        oPointAlongCurve78.Evaluate
    
        Dim oPointByCurves79 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves79 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0079-PointByCurves")
        oPointByCurves79.Inputs("Curve1").Add oLineParallel50, "2"
        oPointByCurves79.Inputs("Curve2").Add oLineExtend69, "1"
        oPointByCurves79.Parameter("TrackFlag") = 1
        oPointByCurves79.Evaluate
    
        Dim oPointByCurves80 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves80 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0080-PointByCurves")
        oPointByCurves80.Inputs("Curve1").Add oLineParallel50, "1"
        oPointByCurves80.Inputs("Curve2").Add oLineExtend70, "1"
        oPointByCurves80.Parameter("TrackFlag") = 1
        oPointByCurves80.Evaluate
    
        Dim oLineStringByPoints81 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineStringByPoints81 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "0081-LineStringByPoints")
        oLineStringByPoints81.Inputs("Points").Add oPointAlongCurve75, "1"
        oLineStringByPoints81.Inputs("Points").Add oPointByCurves73, "2"
        oLineStringByPoints81.Inputs("Points").Add oPointByCurves74, "3"
        oLineStringByPoints81.Inputs("Points").Add oPointAlongCurve77, "4"
        oLineStringByPoints81.Parameter("ClosedFlag") = 0
        oLineStringByPoints81.Evaluate
    
        Dim oSurfByLinearExtrusion82 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion82 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0082-SurfByLinearExtrusion")
        oSurfByLinearExtrusion82.Inputs("PlanarCrossSection").Add oLineStringByPoints81, "1"
        oSurfByLinearExtrusion82.Inputs("ExtrusionLine").Add oLineFromCS27, "1"
        oSurfByLinearExtrusion82.Evaluate
    
        Dim oLineStringByPoints83 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineStringByPoints83 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "0083-LineStringByPoints")
        oLineStringByPoints83.Inputs("Points").Add oPointAlongCurve76, "1"
        oLineStringByPoints83.Inputs("Points").Add oPointByCurves79, "2"
        oLineStringByPoints83.Inputs("Points").Add oPointByCurves80, "3"
        oLineStringByPoints83.Inputs("Points").Add oPointAlongCurve78, "4"
        oLineStringByPoints83.Parameter("ClosedFlag") = 0
        oLineStringByPoints83.Evaluate
    
        Dim oSurfByLinearExtrusion84 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion84 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0084-SurfByLinearExtrusion")
        oSurfByLinearExtrusion84.Inputs("PlanarCrossSection").Add oLineStringByPoints83, "1"
        oSurfByLinearExtrusion84.Inputs("ExtrusionLine").Add oLineFromCS27, "1"
        oSurfByLinearExtrusion84.Evaluate
        
        dT = oParamDistMeasureParallel30.Parameter("Value")
    
        Dim oGCMacro As IJGeometricConstructionMacro
        Set oGCMacro = pGeometricConstruction
    
        oGCMacro.Output("Support", 1) = oExtractPorts18.Output("Support", 1)
        oGCMacro.Output("Boundary", 1) = oSurfFromGType46.Output
        oGCMacro.Output("Boundary", 2) = oSurfFromGType47.Output
        oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion82.Output
        oGCMacro.Output("Boundary", 4) = oSurfByLinearExtrusion84.Output
    '    oGCMacro.Output("Boundary", "FacePort3") = oFacePortExtractor87.Output'(Function:Create LC Between Tube and APS)
        oGCMacro.Output("CoordinateSystem", 1) = oCSFromMember23.Output
        oGCMacro.Output("Node", 1) = oPointByCurves22.Output
    
    Else

'        Dim oLineAxisPortExtractor1 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAxisPortExtractor1 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0001-LineAxisPortExtractor")
        oLineAxisPortExtractor1.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
        oLineAxisPortExtractor1.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port_1_1").Add oLineAxisPortExtractor1.ControlledInputs("Port").Item("1")
        Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0002-ParamDistConstant")
        oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Lengthoffset")
        oParamDistConstant2.Evaluate
    
'        Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0003-ParamDistConstant")
        oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Widthoffset")
        oParamDistConstant3.Evaluate
    
'        Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0004-ParamDistConstant")
        oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("offset")
        oParamDistConstant4.Evaluate
    
        Dim oParamDouble5 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDouble5 = oGCFactory.CreateEntity("ParamDouble", pPOM, "0005-ParamDouble")
        oParamDouble5.Parameter("Value") = pGeometricConstruction.Parameter("slope")
        oParamDouble5.Evaluate
    
        Dim oAxisPortExtractor6 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor6 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "0006-AxisPortExtractor")
        oAxisPortExtractor6.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "2"
        oAxisPortExtractor6.Parameter("GeometrySelector") = 2
        oAxisPortExtractor6.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_11_1").Clear
        pGeometricConstruction.ControlledInputs("Port_11_1").Add oAxisPortExtractor6.ControlledInputs("Port").Item("1")
        Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0007-ParamDistConstant")
        oParamDistConstant7.Parameter("Value") = 3
        oParamDistConstant7.Evaluate
    
        Dim oCSByStiffener8 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByStiffener8 = oGCFactory.CreateEntity("CSByStiffener", pPOM, "0008-CSByStiffener")
        oCSByStiffener8.Inputs("Stiffener").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oCSByStiffener8.Parameter("Ortho") = -1
        oCSByStiffener8.Evaluate
    
        pGeometricConstruction.ControlledInputs("RootPlateSystem_19_1").Clear
        pGeometricConstruction.ControlledInputs("RootPlateSystem_19_1").Add oCSByStiffener8.ControlledInputs("RootPlateSystem").Item("1")
        pGeometricConstruction.ControlledInputs("LeafPlateSystem_19_1").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem_19_1").Add oCSByStiffener8.ControlledInputs("LeafPlateSystem").Item(1)
        Dim oPointAtCurveMiddle9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle9 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0009-PointAtCurveMiddle")
        oPointAtCurveMiddle9.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oPointAtCurveMiddle9.Evaluate
    
        Dim oExtractPorts10 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts10 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0010-ExtractPorts")
        oExtractPorts10.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oExtractPorts10.Inputs("ZAxis").Add oLineAxisPortExtractor1, "1"
        oExtractPorts10.Parameter("Support") = 2
        oExtractPorts10.Parameter("Debug") = 0
        oExtractPorts10.Evaluate
    
        pGeometricConstruction.ControlledInputs("AxisPort_15_1").Clear
        pGeometricConstruction.ControlledInputs("AxisPort_15_1").Add oExtractPorts10.ControlledInputs("AxisPort").Item("1")
    '    Dim oPointByCurves11 As SP3DGeometricConstruction.GeometricConstruction
    '    Set oPointByCurves11 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0011-PointByCurves")
    '    oPointByCurves11.Inputs("Curve1").Add oAxisPortExtractor6, "1"
    '    oPointByCurves11.Inputs("Curve2").Add oLineAxisPortExtractor1, "1"
    '    oPointByCurves11.Parameter("TrackFlag") = 1
    '    oPointByCurves11.Evaluate
    
        Dim oCSByLines80 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines80 = oGCFactory.CreateEntity("CSByLines", pPOM, "0080-CSByLines")
        oCSByLines80.Inputs("AxisLine1").Add oLineAxisPortExtractor1, "1"
        oCSByLines80.Inputs("AxisLine2").Add oAxisPortExtractor6, "1"
        oCSByLines80.Parameter("AxesRoles") = 1
        oCSByLines80.Parameter("CSOrientation") = 1
        oCSByLines80.Parameter("TrackFlag") = 1
        oCSByLines80.Evaluate
        
        Dim oPointFromCS81 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS81 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0081-PointFromCS")
        oPointFromCS81.Inputs("CoordinateSystem").Add oCSByLines80, "1"
        oPointFromCS81.Parameter("X") = 0#
        oPointFromCS81.Parameter("Y") = 0#
        oPointFromCS81.Parameter("Z") = 0#
        oPointFromCS81.Evaluate
        
        Dim oPointAtCurveExtremity82 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity82 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0082-PointAtCurveExtremity")
        oPointAtCurveExtremity82.Inputs("Curve").Add oLineAxisPortExtractor1, "1"
        oPointAtCurveExtremity82.Inputs("TrackPoint").Add oPointFromCS81, "1"
        oPointAtCurveExtremity82.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity82.Evaluate
        
        Dim oPointByCurveAndSurface88 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurveAndSurface88 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0088-PointByCurveAndSurface")
        oPointByCurveAndSurface88.Inputs("Curve1").Add oLineAxisPortExtractor1, "1"
        oPointByCurveAndSurface88.Inputs("Surface2").Add oExtractPorts10.Output("Support", 1), "1"
        oPointByCurveAndSurface88.Evaluate
    
        Dim oCSFromMember12 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSFromMember12 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0012-CSFromMember")
        oCSFromMember12.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
        oCSFromMember12.Inputs("Point").Add oPointAtCurveMiddle9, "1"
        oCSFromMember12.Evaluate
    
        Dim oFacePortExtractor13 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor13 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0013-FacePortExtractor")
        oFacePortExtractor13.Inputs("Connectable").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
        oFacePortExtractor13.Inputs("CoordinateSystem").Add oCSFromMember12, "1"
        oFacePortExtractor13.Parameter("GeometrySelector") = 4
        oFacePortExtractor13.Parameter("FacesContext") = 2
        oFacePortExtractor13.Parameter("LookingAxis") = 2
        oFacePortExtractor13.Parameter("IntersectingPlane") = 0
        oFacePortExtractor13.Parameter("SurfaceType") = 1
        oFacePortExtractor13.Parameter("TrackFlag") = 1
        oFacePortExtractor13.Parameter("Offset") = 0#
        oFacePortExtractor13.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_4_1").Clear
        pGeometricConstruction.ControlledInputs("Port_4_1").Add oFacePortExtractor13.ControlledInputs("Port").Item("1")
        Dim oPlaneFromCS14 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneFromCS14 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "0014-PlaneFromCS")
        oPlaneFromCS14.Inputs("CoordinateSystem").Add oCSByStiffener8, "1"
        oPlaneFromCS14.Parameter("LookingAxis") = 2
        oPlaneFromCS14.Parameter("Offset") = 0.5
        oPlaneFromCS14.Parameter("Range") = 2#
        oPlaneFromCS14.Evaluate
    
'        Dim oCurveByIntersection15 As SP3DGeometricConstruction.GeometricConstruction
'        Set oCurveByIntersection15 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0015-CurveByIntersection")
'        oCurveByIntersection15.Inputs("Surface1").Add oPlaneFromCS14, "1"
'        oCurveByIntersection15.Inputs("Surface2").Add oFacePortExtractor13, "1"
'        oCurveByIntersection15.Parameter("TrackFlag") = 1
'        oCurveByIntersection15.Parameter("ConstructionSurface") = 0
'        oCurveByIntersection15.Evaluate
    
        Dim oLineAtAngle16 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle16 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0016-LineAtAngle")
        oLineAtAngle16.Inputs("Line").Add oAxisPortExtractor6, "1"
        oLineAtAngle16.Inputs("Point").Add oPointByCurveAndSurface88, "1"
        oLineAtAngle16.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle16.Parameter("Angle") = 1.5707963267949
        oLineAtAngle16.Parameter("TrackFlag") = 1
        oLineAtAngle16.Evaluate
        
        Dim oLineAtAngle83 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle83 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0083-LineAtAngle")
        oLineAtAngle83.Inputs("Line").Add oAxisPortExtractor6, "1"
        oLineAtAngle83.Inputs("Point").Add oPointByCurveAndSurface88, "1"
        oLineAtAngle83.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle83.Parameter("Angle") = 0
        oLineAtAngle83.Parameter("TrackFlag") = 1
        oLineAtAngle83.Evaluate
    
        Dim oParamDistFunctionWith2Args17 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args17 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0017-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args17.Inputs("Arg1").Add oParamDistConstant7, "1"
        oParamDistFunctionWith2Args17.Inputs("Arg2").Add oParamDouble5, "1"
        oParamDistFunctionWith2Args17.Parameter("Value") = 0.1
        oParamDistFunctionWith2Args17.Parameter("Operation") = 3
        oParamDistFunctionWith2Args17.Evaluate
    
        Dim oPointAtCurveStart18 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart18 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0018-PointAtCurveStart")
        oPointAtCurveStart18.Inputs("Curve").Add oAxisPortExtractor6, "1"
        oPointAtCurveStart18.Evaluate
    
'        Dim oParamDistMeasureParallel19 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel19 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0019-ParamDistMeasureParallel")
        oParamDistMeasureParallel19.Inputs("Graphics").Add oExtractPorts10.Output("MemberAxis", 1), "1"
        oParamDistMeasureParallel19.Inputs("Graphics").Add oExtractPorts10.Output("CurveRight", 1), "2"
        oParamDistMeasureParallel19.Parameter("Value") = 7.50000000000028E-02
        oParamDistMeasureParallel19.Evaluate
    
        Dim oPointFromCS20 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS20 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0020-PointFromCS")
        oPointFromCS20.Inputs("CoordinateSystem").Add oCSByStiffener8, "1"
        oPointFromCS20.Parameter("X") = 0#
        oPointFromCS20.Parameter("Y") = 0#
        oPointFromCS20.Parameter("Z") = 1#
        oPointFromCS20.Evaluate
    
        Dim oLineFromCS21 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS21 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0021-LineFromCS")
        oLineFromCS21.Inputs("CoordinateSystem").Add oCSByStiffener8, "1"
        oLineFromCS21.Parameter("LookingAxis") = 2
        oLineFromCS21.Parameter("Length") = -0.5
        oLineFromCS21.Parameter("CSOrientation") = 1
        oLineFromCS21.Parameter("LineJustification") = 1
        oLineFromCS21.Evaluate
    
        Dim oPointFromCS22 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS22 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0022-PointFromCS")
        oPointFromCS22.Inputs("CoordinateSystem").Add oCSByStiffener8, "1"
        oPointFromCS22.Parameter("X") = 0#
        oPointFromCS22.Parameter("Y") = 0#
        oPointFromCS22.Parameter("Z") = 0#
        oPointFromCS22.Evaluate
        
        Value1 = pGeometricConstruction.Inputs("SecondaryProfile2").Count
        If Value1 > 0 Then

            Dim oLineAxisPortExtractor84 As SP3DGeometricConstruction.GeometricConstruction
            Set oLineAxisPortExtractor84 = oGCFactory.CreateEntity("LineAxisPortExtractor", pPOM, "0084-LineAxisPortExtractor")
            oLineAxisPortExtractor84.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "1"
            oLineAxisPortExtractor84.Evaluate
            
'            Dim oPointAtCurveMiddle85 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointAtCurveMiddle85 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0085-PointAtCurveMiddle")
            oPointAtCurveMiddle85.Inputs("Curve").Add oLineAxisPortExtractor84, "1"
            oPointAtCurveMiddle85.Evaluate
            
'            Dim oCSFromMember86 As SP3DGeometricConstruction.GeometricConstruction
            Set oCSFromMember86 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0086-CSFromMember")
            oCSFromMember86.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "1"
            oCSFromMember86.Inputs("Point").Add oPointAtCurveMiddle85, "1"
            oCSFromMember86.Evaluate
        
'            Dim oFacePortExtractor87 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortExtractor87 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0087-FacePortExtractor")
            oFacePortExtractor87.Inputs("Connectable").Add pGeometricConstruction.Inputs("SecondaryProfile2").Item(1), "1"
            oFacePortExtractor87.Inputs("CoordinateSystem").Add oCSFromMember86, "1"
            oFacePortExtractor87.Parameter("GeometrySelector") = 4
            oFacePortExtractor87.Parameter("FacesContext") = 2
            oFacePortExtractor87.Parameter("LookingAxis") = 2
            oFacePortExtractor87.Parameter("IntersectingPlane") = 0
            oFacePortExtractor87.Parameter("SurfaceType") = 1
            oFacePortExtractor87.Parameter("TrackFlag") = 1
            oFacePortExtractor87.Parameter("Offset") = 0#
            oFacePortExtractor87.Evaluate
    
            Dim oPointByCurveAndSurface23 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointByCurveAndSurface23 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0023-PointByCurveAndSurface")
            oPointByCurveAndSurface23.Inputs("Curve1").Add oLineAtAngle83, "1"
            oPointByCurveAndSurface23.Inputs("Surface2").Add oFacePortExtractor13, "1"
            oPointByCurveAndSurface23.Inputs("TrackPoint").Add oPointAtCurveMiddle85, "1"
            oPointByCurveAndSurface23.Parameter("TrackFlag") = 2
            oPointByCurveAndSurface23.Evaluate
        
            Dim oPointByCurveAndSurface24 As SP3DGeometricConstruction.GeometricConstruction
            Set oPointByCurveAndSurface24 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0024-PointByCurveAndSurface")
            oPointByCurveAndSurface24.Inputs("Curve1").Add oLineAtAngle83, "1"
            oPointByCurveAndSurface24.Inputs("Surface2").Add oFacePortExtractor87, "1"
            oPointByCurveAndSurface24.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
            oPointByCurveAndSurface24.Parameter("TrackFlag") = 2
            oPointByCurveAndSurface24.Evaluate
        
        
        Else

            Set oPointByCurveAndSurface23 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0023-PointByCurveAndSurface")
            oPointByCurveAndSurface23.Inputs("Curve1").Add oLineAtAngle83, "1"
            oPointByCurveAndSurface23.Inputs("Surface2").Add oFacePortExtractor13, "1"
            oPointByCurveAndSurface23.Inputs("TrackPoint").Add oPointAtCurveStart18, "1"
            oPointByCurveAndSurface23.Parameter("TrackFlag") = 1
            oPointByCurveAndSurface23.Evaluate

            Set oPointByCurveAndSurface24 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0024-PointByCurveAndSurface")
            oPointByCurveAndSurface24.Inputs("Curve1").Add oLineAtAngle83, "1"
            oPointByCurveAndSurface24.Inputs("Surface2").Add oFacePortExtractor13, "1"
            oPointByCurveAndSurface24.Inputs("TrackPoint").Add oPointAtCurveStart18, "1"
            oPointByCurveAndSurface24.Parameter("TrackFlag") = 2
            oPointByCurveAndSurface24.Evaluate
        
        End If
    
        Dim oPointByCurveAndSurface25 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurveAndSurface25 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0025-PointByCurveAndSurface")
        oPointByCurveAndSurface25.Inputs("Curve1").Add oLineAtAngle16, "1"
        oPointByCurveAndSurface25.Inputs("Surface2").Add oFacePortExtractor13, "1"
        oPointByCurveAndSurface25.Inputs("TrackPoint").Add oPointFromCS20, "1"
        oPointByCurveAndSurface25.Parameter("TrackFlag") = 2
        oPointByCurveAndSurface25.Evaluate
    
        Dim oPointByCurveAndSurface26 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurveAndSurface26 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "0026-PointByCurveAndSurface")
        oPointByCurveAndSurface26.Inputs("Curve1").Add oLineAtAngle16, "1"
        oPointByCurveAndSurface26.Inputs("Surface2").Add oFacePortExtractor13, "1"
        oPointByCurveAndSurface26.Inputs("TrackPoint").Add oPointFromCS20, "1"
        oPointByCurveAndSurface26.Parameter("TrackFlag") = 1
        oPointByCurveAndSurface26.Evaluate
    
        Dim oPointAlongCurve27 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve27 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0027-PointAlongCurve")
        oPointAlongCurve27.Inputs("Curve").Add oLineAtAngle83, "1"
        oPointAlongCurve27.Inputs("Point").Add oPointByCurveAndSurface23, "1"
        oPointAlongCurve27.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve27.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve27.Parameter("TrackFlag") = 2
        oPointAlongCurve27.Evaluate
    
        Dim oPointAlongCurve28 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve28 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0028-PointAlongCurve")
        oPointAlongCurve28.Inputs("Curve").Add oLineAtAngle83, "1"
        oPointAlongCurve28.Inputs("Point").Add oPointByCurveAndSurface24, "2"
        oPointAlongCurve28.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve28.Parameter("Distance") = oParamDistConstant2
        oPointAlongCurve28.Parameter("TrackFlag") = 2
        oPointAlongCurve28.Evaluate
    
        Dim oPointAlongCurve29 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve29 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0029-PointAlongCurve")
        oPointAlongCurve29.Inputs("Curve").Add oLineAtAngle16, "1"
        oPointAlongCurve29.Inputs("Point").Add oPointByCurveAndSurface25, "1"
        oPointAlongCurve29.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve29.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve29.Parameter("TrackFlag") = 2
        oPointAlongCurve29.Evaluate
    
        Dim oPointAlongCurve30 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve30 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0030-PointAlongCurve")
        oPointAlongCurve30.Inputs("Curve").Add oLineAtAngle16, "1"
        oPointAlongCurve30.Inputs("Point").Add oPointByCurveAndSurface26, "2"
        oPointAlongCurve30.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve30.Parameter("Distance") = oParamDistConstant3
        oPointAlongCurve30.Parameter("TrackFlag") = 2
        oPointAlongCurve30.Evaluate
    
        Dim oParamDistFunctionWith2Args31 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args31 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0031-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args31.Inputs("Arg1").Add oParamDistMeasureParallel19, "1"
        oParamDistFunctionWith2Args31.Inputs("Arg2").Add oParamDistConstant4, "1"
        oParamDistFunctionWith2Args31.Parameter("Value") = 7.51000000000028E-02
        oParamDistFunctionWith2Args31.Parameter("Operation") = 1
        oParamDistFunctionWith2Args31.Evaluate
    
        Dim oPlaneByPointNormal32 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal32 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0032-PlaneByPointNormal")
        oPlaneByPointNormal32.Inputs("Point").Add oPointAlongCurve27, "1"
        oPlaneByPointNormal32.Inputs("Line").Add oAxisPortExtractor6, "1"
        oPlaneByPointNormal32.Parameter("Range") = 0.5
        oPlaneByPointNormal32.Evaluate
    
        Dim oPlaneByPointNormal33 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal33 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0033-PlaneByPointNormal")
        oPlaneByPointNormal33.Inputs("Point").Add oPointAlongCurve28, "2"
        oPlaneByPointNormal33.Inputs("Line").Add oAxisPortExtractor6, "1"
        oPlaneByPointNormal33.Parameter("Range") = 0.5
        oPlaneByPointNormal33.Evaluate
    
        Dim oLineParallel34 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel34 = oGCFactory.CreateEntity("LineParallel", pPOM, "0034-LineParallel")
        oLineParallel34.Inputs("Line").Add oExtractPorts10.Output("MemberAxis", 1), "1"
        oLineParallel34.Inputs("Plane").Add oExtractPorts10.Output("Support", 1), "1"
        oLineParallel34.Inputs("TrackPoint").Add oPointFromCS20, "1"
        oLineParallel34.Parameter("Distance") = oParamDistFunctionWith2Args31
        oLineParallel34.Parameter("TrackFlag") = 1
        oLineParallel34.Evaluate
    
        Dim oLineParallel35 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel35 = oGCFactory.CreateEntity("LineParallel", pPOM, "0035-LineParallel")
        oLineParallel35.Inputs("Line").Add oExtractPorts10.Output("MemberAxis", 1), "1"
        oLineParallel35.Inputs("Plane").Add oExtractPorts10.Output("Support", 1), "1"
        oLineParallel35.Inputs("TrackPoint").Add oPointFromCS20, "1"
        oLineParallel35.Parameter("Distance") = oParamDistFunctionWith2Args31
        oLineParallel35.Parameter("TrackFlag") = 2
        oLineParallel35.Evaluate
    
        Dim oPointAlongCurve36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve36 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0036-PointAlongCurve")
        oPointAlongCurve36.Inputs("Curve").Add oLineFromCS21, "1"
        oPointAlongCurve36.Inputs("Point").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve36.Inputs("TrackPoint").Add oPointFromCS22, "2"
        oPointAlongCurve36.Parameter("Distance") = 0.3
        oPointAlongCurve36.Parameter("TrackFlag") = 1
        oPointAlongCurve36.Evaluate
    
        Dim oSurfFromGType37 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType37 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0037-SurfFromGType")
        oSurfFromGType37.Inputs("Surface").Add oPlaneByPointNormal32, "1"
        oSurfFromGType37.Evaluate
    
        Dim oSurfFromGType38 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType38 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0038-SurfFromGType")
        oSurfFromGType38.Inputs("Surface").Add oPlaneByPointNormal33, "2"
        oSurfFromGType38.Evaluate
    
        Dim oLineByPoints39 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints39 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0039-LineByPoints")
        oLineByPoints39.Inputs("StartPoint").Add oPointFromCS22, "1"
        oLineByPoints39.Inputs("EndPoint").Add oPointAlongCurve36, "1"
        oLineByPoints39.Evaluate
    
        Dim oLineAtAngle40 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle40 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0040-LineAtAngle")
        oLineAtAngle40.Inputs("Line").Add oLineAtAngle16, "1"
        oLineAtAngle40.Inputs("Point").Add oPointAlongCurve30, "1"
        oLineAtAngle40.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle40.Parameter("Angle") = 1.5707963267949
        oLineAtAngle40.Parameter("TrackFlag") = 1
        oLineAtAngle40.Evaluate
    
        Dim oLineAtAngle41 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle41 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0041-LineAtAngle")
        oLineAtAngle41.Inputs("Line").Add oLineAtAngle16, "1"
        oLineAtAngle41.Inputs("Point").Add oPointAlongCurve29, "1"
        oLineAtAngle41.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle41.Parameter("Angle") = 1.5707963267949
        oLineAtAngle41.Parameter("TrackFlag") = 1
        oLineAtAngle41.Evaluate
    
        Dim oLineAtAngle42 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle42 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0042-LineAtAngle")
        oLineAtAngle42.Inputs("Line").Add oLineAtAngle83, "2"
        oLineAtAngle42.Inputs("Point").Add oPointAlongCurve27, "1"
        oLineAtAngle42.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle42.Parameter("Angle") = 1.5707963267949
        oLineAtAngle42.Parameter("TrackFlag") = 1
        oLineAtAngle42.Evaluate
    
        Dim oLineAtAngle43 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineAtAngle43 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "0043-LineAtAngle")
        oLineAtAngle43.Inputs("Line").Add oLineAtAngle83, "1"
        oLineAtAngle43.Inputs("Point").Add oPointAlongCurve28, "1"
        oLineAtAngle43.Inputs("Plane").Add oPlaneFromCS14, "1"
        oLineAtAngle43.Parameter("Angle") = 1.5707963267949
        oLineAtAngle43.Parameter("TrackFlag") = 1
        oLineAtAngle43.Evaluate
    
        Dim oPlaneByPointNormal44 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal44 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0044-PlaneByPointNormal")
        oPlaneByPointNormal44.Inputs("Point").Add oPointByCurveAndSurface88, "1"
        oPlaneByPointNormal44.Inputs("Line").Add oLineByPoints39, "1"
        oPlaneByPointNormal44.Parameter("Range") = 1#
        oPlaneByPointNormal44.Evaluate
    
        Dim oPointByCurves45 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves45 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0045-PointByCurves")
        oPointByCurves45.Inputs("Curve1").Add oLineAtAngle42, "1"
        oPointByCurves45.Inputs("Curve2").Add oLineParallel34, "1"
        oPointByCurves45.Parameter("TrackFlag") = 1
        oPointByCurves45.Evaluate
    
        Dim oPointByCurves46 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves46 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0046-PointByCurves")
        oPointByCurves46.Inputs("Curve1").Add oLineAtAngle43, "2"
        oPointByCurves46.Inputs("Curve2").Add oLineParallel34, "1"
        oPointByCurves46.Parameter("TrackFlag") = 1
        oPointByCurves46.Evaluate
    
        Dim oPointByCurves47 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves47 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0047-PointByCurves")
        oPointByCurves47.Inputs("Curve1").Add oLineAtAngle42, "1"
        oPointByCurves47.Inputs("Curve2").Add oLineParallel35, "1"
        oPointByCurves47.Parameter("TrackFlag") = 1
        oPointByCurves47.Evaluate
    
        Dim oPointByCurves48 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves48 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0048-PointByCurves")
        oPointByCurves48.Inputs("Curve1").Add oLineAtAngle43, "1"
        oPointByCurves48.Inputs("Curve2").Add oLineParallel35, "1"
        oPointByCurves48.Parameter("TrackFlag") = 1
        oPointByCurves48.Evaluate
    
        Dim oSurfFromGType49 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType49 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0049-SurfFromGType")
        oSurfFromGType49.Inputs("Surface").Add oPlaneByPointNormal44, "1"
        oSurfFromGType49.Evaluate
    
        Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0050-PointAlongCurve")
        oPointAlongCurve50.Inputs("Curve").Add oLineParallel34, "1"
        oPointAlongCurve50.Inputs("Point").Add oPointByCurves46, "1"
        oPointAlongCurve50.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve50.Parameter("Distance") = oParamDistFunctionWith2Args17
        oPointAlongCurve50.Parameter("TrackFlag") = 1
        oPointAlongCurve50.Evaluate
    
        Dim oPointAlongCurve51 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve51 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0051-PointAlongCurve")
        oPointAlongCurve51.Inputs("Curve").Add oLineParallel34, "1"
        oPointAlongCurve51.Inputs("Point").Add oPointByCurves45, "2"
        oPointAlongCurve51.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve51.Parameter("Distance") = oParamDistFunctionWith2Args17
        oPointAlongCurve51.Parameter("TrackFlag") = 1
        oPointAlongCurve51.Evaluate
    
        Dim oPointAlongCurve52 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve52 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0052-PointAlongCurve")
        oPointAlongCurve52.Inputs("Curve").Add oLineParallel35, "2"
        oPointAlongCurve52.Inputs("Point").Add oPointByCurves47, "2"
        oPointAlongCurve52.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve52.Parameter("Distance") = oParamDistFunctionWith2Args17
        oPointAlongCurve52.Parameter("TrackFlag") = 1
        oPointAlongCurve52.Evaluate
    
        Dim oPointAlongCurve53 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0053-PointAlongCurve")
        oPointAlongCurve53.Inputs("Curve").Add oLineParallel35, "1"
        oPointAlongCurve53.Inputs("Point").Add oPointByCurves48, "2"
        oPointAlongCurve53.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve53.Parameter("Distance") = oParamDistFunctionWith2Args17
        oPointAlongCurve53.Parameter("TrackFlag") = 1
        oPointAlongCurve53.Evaluate
    
        Dim oLineByPoints54 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints54 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0054-LineByPoints")
        oLineByPoints54.Inputs("StartPoint").Add oPointAlongCurve50, "1"
        oLineByPoints54.Inputs("EndPoint").Add oPointAlongCurve53, "1"
        oLineByPoints54.Evaluate
    
        Dim oLineByPoints55 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints55 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0055-LineByPoints")
        oLineByPoints55.Inputs("StartPoint").Add oPointAlongCurve51, "2"
        oLineByPoints55.Inputs("EndPoint").Add oPointAlongCurve52, "1"
        oLineByPoints55.Evaluate
    
'        Dim oPointAlongCurve56 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve56 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0056-PointAlongCurve")
        oPointAlongCurve56.Inputs("Curve").Add oLineByPoints54, "1"
        oPointAlongCurve56.Inputs("Point").Add oPointAlongCurve50, "1"
        oPointAlongCurve56.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve56.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve56.Parameter("TrackFlag") = 2
        oPointAlongCurve56.Evaluate
    
'        Dim oPointAlongCurve57 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve57 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0057-PointAlongCurve")
        oPointAlongCurve57.Inputs("Curve").Add oLineByPoints54, "1"
        oPointAlongCurve57.Inputs("Point").Add oPointAlongCurve53, "2"
        oPointAlongCurve57.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve57.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve57.Parameter("TrackFlag") = 2
        oPointAlongCurve57.Evaluate
    
'        Dim oPointAlongCurve58 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve58 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0058-PointAlongCurve")
        oPointAlongCurve58.Inputs("Curve").Add oLineByPoints55, "2"
        oPointAlongCurve58.Inputs("Point").Add oPointAlongCurve52, "2"
        oPointAlongCurve58.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve58.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve58.Parameter("TrackFlag") = 2
        oPointAlongCurve58.Evaluate
    
        Dim oPointAlongCurve59 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve59 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0059-PointAlongCurve")
        oPointAlongCurve59.Inputs("Curve").Add oLineByPoints55, "1"
        oPointAlongCurve59.Inputs("Point").Add oPointAlongCurve51, "2"
        oPointAlongCurve59.Inputs("TrackPoint").Add oPointByCurveAndSurface88, "1"
        oPointAlongCurve59.Parameter("Distance") = oParamDistConstant7
        oPointAlongCurve59.Parameter("TrackFlag") = 2
        oPointAlongCurve59.Evaluate
    
        Dim oLineByPoints60 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints60 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0060-LineByPoints")
        oLineByPoints60.Inputs("StartPoint").Add oPointByCurves45, "2"
        oLineByPoints60.Inputs("EndPoint").Add oPointAlongCurve59, "1"
        oLineByPoints60.Evaluate
    
'        Dim oPointAlongCurve61 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve61 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0061-PointAlongCurve")
        oPointAlongCurve61.Inputs("Curve").Add oLineByPoints60, "1"
        oPointAlongCurve61.Inputs("Point").Add oPointByCurves45, "1"
        oPointAlongCurve61.Inputs("TrackPoint").Add oPointAlongCurve59, "1"
        oPointAlongCurve61.Parameter("Distance") = 0.05
        oPointAlongCurve61.Parameter("TrackFlag") = 2
        oPointAlongCurve61.Evaluate
    
        Dim oLineByPoints62 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints62 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0062-LineByPoints")
        oLineByPoints62.Inputs("StartPoint").Add oPointAlongCurve58, "2"
        oLineByPoints62.Inputs("EndPoint").Add oPointByCurves47, "1"
        oLineByPoints62.Evaluate
    
'        Dim oPointAlongCurve63 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve63 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0063-PointAlongCurve")
        oPointAlongCurve63.Inputs("Curve").Add oLineByPoints62, "2"
        oPointAlongCurve63.Inputs("Point").Add oPointByCurves47, "2"
        oPointAlongCurve63.Inputs("TrackPoint").Add oPointAlongCurve58, "2"
        oPointAlongCurve63.Parameter("Distance") = 0.05
        oPointAlongCurve63.Parameter("TrackFlag") = 2
        oPointAlongCurve63.Evaluate
    
        Dim oLineByPoints64 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints64 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0064-LineByPoints")
        oLineByPoints64.Inputs("StartPoint").Add oPointAlongCurve56, "1"
        oLineByPoints64.Inputs("EndPoint").Add oPointByCurves46, "1"
        oLineByPoints64.Evaluate
    
        Dim oPointAlongCurve65 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve65 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0065-PointAlongCurve")
        oPointAlongCurve65.Inputs("Curve").Add oLineByPoints64, "2"
        oPointAlongCurve65.Inputs("Point").Add oPointByCurves46, "2"
        oPointAlongCurve65.Inputs("TrackPoint").Add oPointAlongCurve56, "2"
        oPointAlongCurve65.Parameter("Distance") = 0.05
        oPointAlongCurve65.Parameter("TrackFlag") = 2
        oPointAlongCurve65.Evaluate
    
'        Dim oLineByPoints66 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints66 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0066-LineByPoints")
        oLineByPoints66.Inputs("StartPoint").Add oPointByCurves48, "2"
        oLineByPoints66.Inputs("EndPoint").Add oPointAlongCurve57, "1"
        oLineByPoints66.Evaluate
    
        Dim oPointAlongCurve67 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve67 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0067-PointAlongCurve")
        oPointAlongCurve67.Inputs("Curve").Add oLineByPoints66, "2"
        oPointAlongCurve67.Inputs("Point").Add oPointByCurves48, "2"
        oPointAlongCurve67.Inputs("TrackPoint").Add oPointAlongCurve57, "2"
        oPointAlongCurve67.Parameter("Distance") = 0.05
        oPointAlongCurve67.Parameter("TrackFlag") = 2
        oPointAlongCurve67.Evaluate
    
'        Dim oLineExtend68 As SP3DGeometricConstruction.GeometricConstruction
'        Set oLineExtend68 = oGCFactory.CreateEntity("LineExtend", pPOM, "0068-LineExtend")
'        oLineExtend68.Inputs("Line").Add oLineByPoints60, "1"
'        oLineExtend68.Inputs("Curve1").Add oLineAtAngle40, "1"
'        oLineExtend68.Parameter("Context1") = 3
'        oLineExtend68.Parameter("Context2") = 1
'        oLineExtend68.Evaluate
'
'        Dim oLineExtend69 As SP3DGeometricConstruction.GeometricConstruction
'        Set oLineExtend69 = oGCFactory.CreateEntity("LineExtend", pPOM, "0069-LineExtend")
'        oLineExtend69.Inputs("Line").Add oLineByPoints64, "1"
'        oLineExtend69.Inputs("Curve1").Add oLineAtAngle40, "1"
'        oLineExtend69.Parameter("Context1") = 1
'        oLineExtend69.Parameter("Context2") = 1
'        oLineExtend69.Evaluate
    
        Dim oPointByCurves70 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves70 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0070-PointByCurves")
        oPointByCurves70.Inputs("Curve1").Add oLineAtAngle40, "2"
        oPointByCurves70.Inputs("Curve2").Add oLineByPoints60, "1"
        oPointByCurves70.Parameter("TrackFlag") = 1
        oPointByCurves70.Evaluate
    
        Dim oPointByCurves71 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves71 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0071-PointByCurves")
        oPointByCurves71.Inputs("Curve1").Add oLineAtAngle40, "1"
        oPointByCurves71.Inputs("Curve2").Add oLineByPoints64, "1"
        oPointByCurves71.Parameter("TrackFlag") = 1
        oPointByCurves71.Evaluate
    
'        Dim oLineExtend72 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend72 = oGCFactory.CreateEntity("LineExtend", pPOM, "0072-LineExtend")
        oLineExtend72.Inputs("Line").Add oLineAtAngle41, "1"
        oLineExtend72.Inputs("Curve1").Add oLineByPoints62, "1"
        oLineExtend72.Parameter("Context1") = 1
        oLineExtend72.Parameter("Context2") = 3
        oLineExtend72.Evaluate

        Dim oLineExtend73 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineExtend73 = oGCFactory.CreateEntity("LineExtend", pPOM, "0073-LineExtend")
        oLineExtend73.Inputs("Line").Add oLineAtAngle41, "1"
        oLineExtend73.Inputs("Curve1").Add oLineByPoints66, "1"
        oLineExtend73.Parameter("Context1") = 1
        oLineExtend73.Parameter("Context2") = 3
        oLineExtend73.Evaluate
    
'        Dim oPointByCurves74 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves74 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0074-PointByCurves")
        oPointByCurves74.Inputs("Curve1").Add oLineExtend72, "1" 'oLineAtAngle41, "2"
        oPointByCurves74.Inputs("Curve2").Add oLineByPoints62, "1"
        oPointByCurves74.Parameter("TrackFlag") = 1
        oPointByCurves74.Evaluate
    
        Dim oPointByCurves75 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves75 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0075-PointByCurves")
        oPointByCurves75.Inputs("Curve1").Add oLineExtend73, "1" 'oLineAtAngle41, "1"
        oPointByCurves75.Inputs("Curve2").Add oLineByPoints66, "1"
        oPointByCurves75.Parameter("TrackFlag") = 1
        oPointByCurves75.Evaluate
    
        Dim oLineStringByPoints76 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineStringByPoints76 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "0076-LineStringByPoints")
        oLineStringByPoints76.Inputs("Points").Add oPointAlongCurve61, "1"
        oLineStringByPoints76.Inputs("Points").Add oPointByCurves70, "2"
        oLineStringByPoints76.Inputs("Points").Add oPointByCurves71, "3"
        oLineStringByPoints76.Inputs("Points").Add oPointAlongCurve65, "4"
        oLineStringByPoints76.Parameter("ClosedFlag") = 0
        oLineStringByPoints76.Evaluate
    
        Dim oLineStringByPoints77 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineStringByPoints77 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "0077-LineStringByPoints")
        oLineStringByPoints77.Inputs("Points").Add oPointAlongCurve63, "1"
        oLineStringByPoints77.Inputs("Points").Add oPointByCurves74, "2"
        oLineStringByPoints77.Inputs("Points").Add oPointByCurves75, "3"
        oLineStringByPoints77.Inputs("Points").Add oPointAlongCurve67, "4"
        oLineStringByPoints77.Parameter("ClosedFlag") = 0
        oLineStringByPoints77.Evaluate
    
        Dim oSurfByLinearExtrusion78 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion78 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0078-SurfByLinearExtrusion")
        oSurfByLinearExtrusion78.Inputs("PlanarCrossSection").Add oLineStringByPoints76, "1"
        oSurfByLinearExtrusion78.Inputs("ExtrusionLine").Add oLineFromCS21, "1"
        oSurfByLinearExtrusion78.Evaluate
    
        Dim oSurfByLinearExtrusion79 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion79 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0079-SurfByLinearExtrusion")
        oSurfByLinearExtrusion79.Inputs("PlanarCrossSection").Add oLineStringByPoints77, "1"
        oSurfByLinearExtrusion79.Inputs("ExtrusionLine").Add oLineFromCS21, "1"
        oSurfByLinearExtrusion79.Evaluate
    
'        Dim oGCMacro As IJGeometricConstructionMacro
        Set oGCMacro = pGeometricConstruction
    
        oGCMacro.Output("Support", 1) = oSurfFromGType49.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion78.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion79.Output
        oGCMacro.Output("Boundary", 3) = oSurfFromGType37.Output
        oGCMacro.Output("Boundary", 4) = oSurfFromGType38.Output
        oGCMacro.Output("CoordinateSystem", 1) = oCSByStiffener8.Output
        oGCMacro.Output("Node", 1) = oPointByCurveAndSurface88.Output

    End If
    
End Sub
