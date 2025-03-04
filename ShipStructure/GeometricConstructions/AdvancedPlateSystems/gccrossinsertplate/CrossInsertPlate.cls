VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrossInsertPlate"
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
    Let Source = "GCAPSWithTrimBack3.CrossInsertPlate"
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
    Call pGC.RaiseError(Err.Number)
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
    ' until TR-209576 (An input defined with ComputeIIDs="NoInterface" does not get split notification) is fixed, we have to specify an interface, which does not trigger any semantic attached to a Member
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select a supporting Member (in 1 or 2 pieces)", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 2, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersCheckIntersectionPlate;MembersSplitByPlate;MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile2", "Select a first supported Member ", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile2", "BehaviorOfPrimaryProfile2", 1, 0, 0, 0, 0, "MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile3", "Select a second aligned supported Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile3", "BehaviorOfPrimaryProfile3", 1, 0, 0, 0, 0, "MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_2", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_3", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_4")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Direction", "Direction", 4, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Direction", "Negative", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Direction", "Positive", 2)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 0.1)
    
    Call GeometricConstructionDefinition_CopyErrors(pGeometricConstructionDefinition, GeometricConstructionDefinition_GetFromTypeName("OrderAxes2"), 100)

    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
  
    On Error Resume Next
    Dim oOrderAxes21 As SP3DGeometricConstruction.GeometricConstruction
    Set oOrderAxes21 = oGCFactory.CreateEntity("OrderAxes2", pPOM, "001-OrderAxes2")
    oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
    If pGeometricConstruction.Inputs("PrimaryProfile1").Count = 2 Then
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(2), "2"
    End If
    oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "3"
    oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile3").Item(1), "4"
    oOrderAxes21.Parameter("Support") = pGeometricConstruction.Parameter("Direction")
    oOrderAxes21.Evaluate
    Dim lError As Long: Let lError = Err.Number
    On Error GoTo 0
    If lError <> 0 Then
        Err.Raise m_oGenericPlateHelper.GetErrorNumber(oOrderAxes21) + 100
    End If
    
    pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_1").Add oOrderAxes21.ControlledInputs("Port1").Item("1")
    If pGeometricConstruction.Inputs("PrimaryProfile1").Count = 2 Then
        pGeometricConstruction.ControlledInputs("Port1_1_1").Add oOrderAxes21.ControlledInputs("Port1").Item("2")
    End If
    pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_2").Add oOrderAxes21.ControlledInputs("Port1").Item("3")
    pGeometricConstruction.ControlledInputs("Port1_1_3").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_3").Add oOrderAxes21.ControlledInputs("Port1").Item("4")

    'Filed DI-CP-233119 on Geometric Constructions to createOrder Axes with Stiffener Part Axis
    'once the DI is done then below fix can be removed
    'For Stiffeners Update the OrderAxes with Stiffener Parts Axis
    If TypeOf pGeometricConstruction.Inputs("PrimaryProfile1").Item(1) Is IJStiffener Then
        UpdateOrderAxeswithStiffenerParts oOrderAxes21, pGeometricConstruction
    End If

    Dim oParamDouble2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble2 = oGCFactory.CreateEntity("ParamDouble", pPOM, "002-ParamDouble")
    oParamDouble2.Parameter("Value") = 2#
    oParamDouble2.Evaluate

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = pGeometricConstruction.Parameter("Length")
    oParamDistConstant3.Evaluate

    Dim oPlaneFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS4 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "004-PlaneFromCS")
    oPlaneFromCS4.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS4.Parameter("LookingAxis") = 3
    oPlaneFromCS4.Parameter("Offset") = 0#
    oPlaneFromCS4.Parameter("Range") = 3#
    oPlaneFromCS4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = pGeometricConstruction.Parameter("Radius")
    oParamDistConstant5.Evaluate

    Dim oParamDistFunctionWith2Args6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args6 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "006-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args6.Inputs("Arg1").Add oParamDistConstant3
    oParamDistFunctionWith2Args6.Inputs("Arg2").Add oParamDouble2
    oParamDistFunctionWith2Args6.Parameter("Value") = 0.5
    oParamDistFunctionWith2Args6.Parameter("Operation") = 4
    oParamDistFunctionWith2Args6.Evaluate

    Dim oPlaneFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS7 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "007-PlaneFromCS")
    oPlaneFromCS7.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS7.Parameter("LookingAxis") = 1
    oPlaneFromCS7.Parameter("Offset") = oParamDistConstant3
    oPlaneFromCS7.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS7.Evaluate

    Dim oPlaneFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS8 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "008-PlaneFromCS")
    oPlaneFromCS8.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS8.Parameter("LookingAxis") = 2
    oPlaneFromCS8.Parameter("Offset") = oParamDistConstant3
    oPlaneFromCS8.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS8.Evaluate

    Dim oParamDistFunctionWith1Arg9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg9 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "009-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg9.Inputs("Arg").Add oParamDistConstant3
    oParamDistFunctionWith1Arg9.Parameter("Value") = -1#
    oParamDistFunctionWith1Arg9.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg9.Evaluate

    Dim oPlaneFromCS10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS10 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "010-PlaneFromCS")
    oPlaneFromCS10.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS10.Parameter("LookingAxis") = 2
    oPlaneFromCS10.Parameter("Offset") = oParamDistFunctionWith1Arg9
    oPlaneFromCS10.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS10.Evaluate

    Dim oPlaneFromCS11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS11 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "011-PlaneFromCS")
    oPlaneFromCS11.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS11.Parameter("LookingAxis") = 1
    oPlaneFromCS11.Parameter("Offset") = oParamDistFunctionWith1Arg9
    oPlaneFromCS11.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS11.Evaluate

    Dim oParamDistFunctionWith1Arg12 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg12 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "012-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg12.Inputs("Arg").Add oParamDistFunctionWith2Args6
    oParamDistFunctionWith1Arg12.Parameter("Value") = -0.5
    oParamDistFunctionWith1Arg12.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg12.Evaluate

    Dim oPointFromCS13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS13 = oGCFactory.CreateEntity("PointFromCS", pPOM, "013-PointFromCS")
    oPointFromCS13.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS13.Parameter("X") = oParamDistFunctionWith2Args6
    oPointFromCS13.Parameter("Y") = 0#
    oPointFromCS13.Parameter("Z") = 0#
    oPointFromCS13.Evaluate

    Dim oPointFromCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS14 = oGCFactory.CreateEntity("PointFromCS", pPOM, "014-PointFromCS")
    oPointFromCS14.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS14.Parameter("X") = 0#
    oPointFromCS14.Parameter("Y") = oParamDistFunctionWith2Args6
    oPointFromCS14.Parameter("Z") = 0#
    oPointFromCS14.Evaluate

    Dim oPointFromCS15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS15 = oGCFactory.CreateEntity("PointFromCS", pPOM, "015-PointFromCS")
    oPointFromCS15.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS15.Parameter("X") = 0#
    oPointFromCS15.Parameter("Y") = 0#
    oPointFromCS15.Parameter("Z") = 2
    oPointFromCS15.Evaluate

    Dim oCurveByProjection16 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection16 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "016-CurveByProjection")
    oCurveByProjection16.Inputs("Curve").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection16.Inputs("Surface").Add oPlaneFromCS7
    oCurveByProjection16.Inputs("Line").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection16.Parameter("TrackFlag") = 1
    oCurveByProjection16.Parameter("ConstructionSurface") = 0
    oCurveByProjection16.Evaluate

    Dim oCurveByProjection17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection17 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "017-CurveByProjection")
    oCurveByProjection17.Inputs("Curve").Add oOrderAxes21.Output("Axis", 2)
    oCurveByProjection17.Inputs("Surface").Add oPlaneFromCS7
    oCurveByProjection17.Inputs("Line").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection17.Parameter("TrackFlag") = 1
    oCurveByProjection17.Parameter("ConstructionSurface") = 0
    oCurveByProjection17.Evaluate

    Dim oSurfFromGType18 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType18 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "018-SurfFromGType")
    oSurfFromGType18.Inputs("Surface").Add oPlaneFromCS7
    oSurfFromGType18.Evaluate

    Dim oPointByProjectOnSurf19 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf19 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "019-PointByProjectOnSurf")
    oPointByProjectOnSurf19.Inputs("Point").Add oPointFromCS14
    oPointByProjectOnSurf19.Inputs("Surface").Add oPlaneFromCS7
    oPointByProjectOnSurf19.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf19.Evaluate

    Dim oCurveByProjection20 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection20 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "020-CurveByProjection")
    oCurveByProjection20.Inputs("Curve").Add oOrderAxes21.Output("Axis", 3)
    oCurveByProjection20.Inputs("Surface").Add oPlaneFromCS8
    oCurveByProjection20.Inputs("Line").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection20.Parameter("TrackFlag") = 1
    oCurveByProjection20.Parameter("ConstructionSurface") = 0
    oCurveByProjection20.Evaluate

    Dim oSurfFromGType21 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType21 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "021-SurfFromGType")
    oSurfFromGType21.Inputs("Surface").Add oPlaneFromCS8
    oSurfFromGType21.Evaluate

    Dim oCurveByProjection22 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection22 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "022-CurveByProjection")
    oCurveByProjection22.Inputs("Curve").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection22.Inputs("Surface").Add oPlaneFromCS8
    oCurveByProjection22.Inputs("Line").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection22.Parameter("TrackFlag") = 1
    oCurveByProjection22.Parameter("ConstructionSurface") = 0
    oCurveByProjection22.Evaluate

    Dim oPointByProjectOnSurf23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf23 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "023-PointByProjectOnSurf")
    oPointByProjectOnSurf23.Inputs("Point").Add oPointFromCS13
    oPointByProjectOnSurf23.Inputs("Surface").Add oPlaneFromCS8
    oPointByProjectOnSurf23.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf23.Evaluate

    Dim oCurveByProjection24 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection24 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "024-CurveByProjection")
    oCurveByProjection24.Inputs("Curve").Add oOrderAxes21.Output("Axis", 3)
    oCurveByProjection24.Inputs("Surface").Add oPlaneFromCS10
    oCurveByProjection24.Inputs("Line").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection24.Parameter("TrackFlag") = 1
    oCurveByProjection24.Parameter("ConstructionSurface") = 0
    oCurveByProjection24.Evaluate

    Dim oSurfFromGType25 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType25 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "025-SurfFromGType")
    oSurfFromGType25.Inputs("Surface").Add oPlaneFromCS10
    oSurfFromGType25.Evaluate

    Dim oCurveByProjection26 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection26 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "026-CurveByProjection")
    oCurveByProjection26.Inputs("Curve").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection26.Inputs("Surface").Add oPlaneFromCS10
    oCurveByProjection26.Inputs("Line").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection26.Parameter("TrackFlag") = 1
    oCurveByProjection26.Parameter("ConstructionSurface") = 0
    oCurveByProjection26.Evaluate

    Dim oPointByProjectOnSurf27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf27 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "027-PointByProjectOnSurf")
    oPointByProjectOnSurf27.Inputs("Point").Add oPointFromCS13
    oPointByProjectOnSurf27.Inputs("Surface").Add oPlaneFromCS10
    oPointByProjectOnSurf27.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf27.Evaluate

    Dim oCurveByProjection28 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection28 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "028-CurveByProjection")
    oCurveByProjection28.Inputs("Curve").Add oOrderAxes21.Output("Axis", 4)
    oCurveByProjection28.Inputs("Surface").Add oPlaneFromCS11
    oCurveByProjection28.Inputs("Line").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection28.Parameter("TrackFlag") = 1
    oCurveByProjection28.Parameter("ConstructionSurface") = 0
    oCurveByProjection28.Evaluate

    Dim oCurveByProjection29 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByProjection29 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "029-CurveByProjection")
    oCurveByProjection29.Inputs("Curve").Add oOrderAxes21.Output("Axis", 2)
    oCurveByProjection29.Inputs("Surface").Add oPlaneFromCS11
    oCurveByProjection29.Inputs("Line").Add oOrderAxes21.Output("Axis", 1)
    oCurveByProjection29.Parameter("TrackFlag") = 1
    oCurveByProjection29.Parameter("ConstructionSurface") = 0
    oCurveByProjection29.Evaluate

    Dim oSurfFromGType30 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType30 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "030-SurfFromGType")
    oSurfFromGType30.Inputs("Surface").Add oPlaneFromCS11
    oSurfFromGType30.Evaluate

    Dim oPointByProjectOnSurf31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf31 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "031-PointByProjectOnSurf")
    oPointByProjectOnSurf31.Inputs("Point").Add oPointFromCS14
    oPointByProjectOnSurf31.Inputs("Surface").Add oPlaneFromCS11
    oPointByProjectOnSurf31.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf31.Evaluate

    Dim oPointFromCS32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS32 = oGCFactory.CreateEntity("PointFromCS", pPOM, "032-PointFromCS")
    oPointFromCS32.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS32.Parameter("X") = oParamDistFunctionWith1Arg12
    oPointFromCS32.Parameter("Y") = 0#
    oPointFromCS32.Parameter("Z") = 0#
    oPointFromCS32.Evaluate

    Dim oPointFromCS33 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS33 = oGCFactory.CreateEntity("PointFromCS", pPOM, "033-PointFromCS")
    oPointFromCS33.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS33.Parameter("X") = 0#
    oPointFromCS33.Parameter("Y") = oParamDistFunctionWith1Arg12
    oPointFromCS33.Parameter("Z") = 0#
    oPointFromCS33.Evaluate

    Dim oPointFromCS34 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS34 = oGCFactory.CreateEntity("PointFromCS", pPOM, "034-PointFromCS")
    oPointFromCS34.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS34.Parameter("X") = 0#
    oPointFromCS34.Parameter("Y") = 0#
    oPointFromCS34.Parameter("Z") = -2
    oPointFromCS34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointFromCS13
    oLineByPoints35.Inputs("EndPoint").Add oPointByProjectOnSurf23
    oLineByPoints35.Evaluate

    Dim oLineByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints36 = oGCFactory.CreateEntity("LineByPoints", pPOM, "036-LineByPoints")
    oLineByPoints36.Inputs("StartPoint").Add oPointFromCS13
    oLineByPoints36.Inputs("EndPoint").Add oPointByProjectOnSurf27
    oLineByPoints36.Evaluate

    Dim oLineByPoints37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints37 = oGCFactory.CreateEntity("LineByPoints", pPOM, "037-LineByPoints")
    oLineByPoints37.Inputs("StartPoint").Add oPointByProjectOnSurf31
    oLineByPoints37.Inputs("EndPoint").Add oPointFromCS14
    oLineByPoints37.Evaluate

    Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "038-LineByPoints")
    oLineByPoints38.Inputs("StartPoint").Add oPointFromCS14
    oLineByPoints38.Inputs("EndPoint").Add oPointByProjectOnSurf19
    oLineByPoints38.Evaluate

    Dim oLineExtend39 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend39 = oGCFactory.CreateEntity("LineExtend", pPOM, "039-LineExtend")
    oLineExtend39.Inputs("Line").Add oLineByPoints35
    oLineExtend39.Inputs("Curve1").Add oLineByPoints38
    oLineExtend39.Parameter("Context1") = 1
    oLineExtend39.Parameter("Context2") = 3
    oLineExtend39.Evaluate

    Dim oLineExtend40 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend40 = oGCFactory.CreateEntity("LineExtend", pPOM, "040-LineExtend")
    oLineExtend40.Inputs("Line").Add oLineByPoints38
    oLineExtend40.Inputs("Curve1").Add oLineExtend39
    oLineExtend40.Parameter("Context1") = 1
    oLineExtend40.Parameter("Context2") = 3
    oLineExtend40.Evaluate

    Dim oArcByFillet41 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet41 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "041-ArcByFillet")
    oArcByFillet41.Inputs("Curve1").Add oLineExtend40
    oArcByFillet41.Inputs("Curve2").Add oLineExtend39
    oArcByFillet41.Parameter("Radius") = oParamDistConstant5
    oArcByFillet41.Evaluate

    Dim oLineExtend42 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend42 = oGCFactory.CreateEntity("LineExtend", pPOM, "042-LineExtend")
    oLineExtend42.Inputs("Line").Add oLineExtend39
    oLineExtend42.Inputs("Curve1").Add oArcByFillet41
    oLineExtend42.Parameter("Context1") = 1
    oLineExtend42.Parameter("Context2") = 3
    oLineExtend42.Evaluate

    Dim oLineExtend43 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend43 = oGCFactory.CreateEntity("LineExtend", pPOM, "043-LineExtend")
    oLineExtend43.Inputs("Line").Add oLineExtend40
    oLineExtend43.Inputs("Curve1").Add oArcByFillet41
    oLineExtend43.Parameter("Context1") = 1
    oLineExtend43.Parameter("Context2") = 3
    oLineExtend43.Evaluate

    Dim oCurveByCurves44 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves44 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "044-CurveByCurves")
    oCurveByCurves44.Inputs("Curves").Add oLineExtend43
    oCurveByCurves44.Inputs("Curves").Add oArcByFillet41
    oCurveByCurves44.Inputs("Curves").Add oLineExtend42
    oCurveByCurves44.Parameter("ConstructionSurface") = 0
    oCurveByCurves44.Evaluate

    Dim oLineByPoints45 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints45 = oGCFactory.CreateEntity("LineByPoints", pPOM, "045-LineByPoints")
    oLineByPoints45.Inputs("StartPoint").Add oPointFromCS34
    oLineByPoints45.Inputs("EndPoint").Add oPointFromCS15
    oLineByPoints45.Evaluate

  
    Dim oPointByProjectOnSurf47 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf47 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "047-PointByProjectOnSurf")
    oPointByProjectOnSurf47.Inputs("Point").Add oPointFromCS33
    oPointByProjectOnSurf47.Inputs("Surface").Add oPlaneFromCS7
    oPointByProjectOnSurf47.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf47.Evaluate

    Dim oPointByProjectOnSurf48 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf48 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "048-PointByProjectOnSurf")
    oPointByProjectOnSurf48.Inputs("Point").Add oPointFromCS32
    oPointByProjectOnSurf48.Inputs("Surface").Add oPlaneFromCS8
    oPointByProjectOnSurf48.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf48.Evaluate

    Dim oPointByProjectOnSurf49 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf49 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "049-PointByProjectOnSurf")
    oPointByProjectOnSurf49.Inputs("Point").Add oPointFromCS32
    oPointByProjectOnSurf49.Inputs("Surface").Add oPlaneFromCS10
    oPointByProjectOnSurf49.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf49.Evaluate

    Dim oPointByProjectOnSurf50 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf50 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "050-PointByProjectOnSurf")
    oPointByProjectOnSurf50.Inputs("Point").Add oPointFromCS33
    oPointByProjectOnSurf50.Inputs("Surface").Add oPlaneFromCS11
    oPointByProjectOnSurf50.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf50.Evaluate

    Dim oLineByPoints51 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints51 = oGCFactory.CreateEntity("LineByPoints", pPOM, "051-LineByPoints")
    oLineByPoints51.Inputs("StartPoint").Add oPointByProjectOnSurf48
    oLineByPoints51.Inputs("EndPoint").Add oPointFromCS32
    oLineByPoints51.Evaluate

    Dim oLineByPoints52 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints52 = oGCFactory.CreateEntity("LineByPoints", pPOM, "052-LineByPoints")
    oLineByPoints52.Inputs("StartPoint").Add oPointByProjectOnSurf49
    oLineByPoints52.Inputs("EndPoint").Add oPointFromCS32
    oLineByPoints52.Evaluate

    Dim oLineByPoints53 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints53 = oGCFactory.CreateEntity("LineByPoints", pPOM, "053-LineByPoints")
    oLineByPoints53.Inputs("StartPoint").Add oPointFromCS33
    oLineByPoints53.Inputs("EndPoint").Add oPointByProjectOnSurf50
    oLineByPoints53.Evaluate

    Dim oLineByPoints54 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints54 = oGCFactory.CreateEntity("LineByPoints", pPOM, "054-LineByPoints")
    oLineByPoints54.Inputs("StartPoint").Add oPointByProjectOnSurf47
    oLineByPoints54.Inputs("EndPoint").Add oPointFromCS33
    oLineByPoints54.Evaluate

    Dim oSurfByLinearExtrusion55 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion55 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "055-SurfByLinearExtrusion")
    oSurfByLinearExtrusion55.Inputs("PlanarCrossSection").Add oCurveByCurves44
    oSurfByLinearExtrusion55.Inputs("ExtrusionLine").Add oLineByPoints45
    oSurfByLinearExtrusion55.Evaluate

    Dim oLineExtend56 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend56 = oGCFactory.CreateEntity("LineExtend", pPOM, "056-LineExtend")
    oLineExtend56.Inputs("Line").Add oLineByPoints54
    oLineExtend56.Inputs("Curve1").Add oLineByPoints36
    oLineExtend56.Inputs("Curve2").Add oCurveByProjection16
    oLineExtend56.Parameter("Context1") = 1
    oLineExtend56.Parameter("Context2") = 3
    oLineExtend56.Evaluate

    Dim oLineExtend57 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend57 = oGCFactory.CreateEntity("LineExtend", pPOM, "057-LineExtend")
    oLineExtend57.Inputs("Line").Add oLineByPoints51
    oLineExtend57.Inputs("Curve1").Add oLineByPoints37
    oLineExtend57.Inputs("Curve2").Add oCurveByProjection20
    oLineExtend57.Parameter("Context1") = 1
    oLineExtend57.Parameter("Context2") = 3
    oLineExtend57.Evaluate

    Dim oLineExtend58 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend58 = oGCFactory.CreateEntity("LineExtend", pPOM, "058-LineExtend")
    oLineExtend58.Inputs("Line").Add oLineByPoints52
    oLineExtend58.Inputs("Curve1").Add oLineByPoints53
    oLineExtend58.Inputs("Curve2").Add oCurveByProjection24
    oLineExtend58.Parameter("Context1") = 1
    oLineExtend58.Parameter("Context2") = 3
    oLineExtend58.Evaluate

    Dim oLineExtend59 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend59 = oGCFactory.CreateEntity("LineExtend", pPOM, "059-LineExtend")
    oLineExtend59.Inputs("Line").Add oLineByPoints36
    oLineExtend59.Inputs("Curve1").Add oLineExtend56
    oLineExtend59.Parameter("Context1") = 1
    oLineExtend59.Parameter("Context2") = 3
    oLineExtend59.Evaluate

    Dim oLineExtend60 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend60 = oGCFactory.CreateEntity("LineExtend", pPOM, "060-LineExtend")
    oLineExtend60.Inputs("Line").Add oLineByPoints53
    oLineExtend60.Inputs("Curve1").Add oLineExtend58
    oLineExtend60.Parameter("Context1") = 1
    oLineExtend60.Parameter("Context2") = 3
    oLineExtend60.Evaluate

    Dim oArcByFillet61 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet61 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "061-ArcByFillet")
    oArcByFillet61.Inputs("Curve1").Add oLineExtend60
    oArcByFillet61.Inputs("Curve2").Add oLineExtend58
    oArcByFillet61.Parameter("Radius") = oParamDistConstant5
    oArcByFillet61.Evaluate

    Dim oLineExtend62 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend62 = oGCFactory.CreateEntity("LineExtend", pPOM, "062-LineExtend")
    oLineExtend62.Inputs("Line").Add oLineExtend58
    oLineExtend62.Inputs("Curve1").Add oArcByFillet61
    oLineExtend62.Parameter("Context1") = 1
    oLineExtend62.Parameter("Context2") = 3
    oLineExtend62.Evaluate

    Dim oLineExtend63 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend63 = oGCFactory.CreateEntity("LineExtend", pPOM, "063-LineExtend")
    oLineExtend63.Inputs("Line").Add oLineExtend60
    oLineExtend63.Inputs("Curve1").Add oArcByFillet61
    oLineExtend63.Parameter("Context1") = 1
    oLineExtend63.Parameter("Context2") = 3
    oLineExtend63.Evaluate

    Dim oArcByFillet64 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet64 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "064-ArcByFillet")
    oArcByFillet64.Inputs("Curve1").Add oLineExtend56
    oArcByFillet64.Inputs("Curve2").Add oLineExtend59
    oArcByFillet64.Parameter("Radius") = oParamDistConstant5
    oArcByFillet64.Evaluate

    Dim oLineExtend65 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend65 = oGCFactory.CreateEntity("LineExtend", pPOM, "065-LineExtend")
    oLineExtend65.Inputs("Line").Add oLineExtend56
    oLineExtend65.Inputs("Curve1").Add oArcByFillet64
    oLineExtend65.Parameter("Context1") = 1
    oLineExtend65.Parameter("Context2") = 3
    oLineExtend65.Evaluate

    Dim oLineExtend66 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend66 = oGCFactory.CreateEntity("LineExtend", pPOM, "066-LineExtend")
    oLineExtend66.Inputs("Line").Add oLineExtend59
    oLineExtend66.Inputs("Curve1").Add oArcByFillet64
    oLineExtend66.Parameter("Context1") = 1
    oLineExtend66.Parameter("Context2") = 3
    oLineExtend66.Evaluate

    Dim oCurveByCurves67 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves67 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "067-CurveByCurves")
    oCurveByCurves67.Inputs("Curves").Add oLineExtend63
    oCurveByCurves67.Inputs("Curves").Add oArcByFillet61
    oCurveByCurves67.Inputs("Curves").Add oLineExtend62
    oCurveByCurves67.Parameter("ConstructionSurface") = 0
    oCurveByCurves67.Evaluate

    Dim oCurveByCurves68 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves68 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "068-CurveByCurves")
    oCurveByCurves68.Inputs("Curves").Add oLineExtend65
    oCurveByCurves68.Inputs("Curves").Add oArcByFillet64
    oCurveByCurves68.Inputs("Curves").Add oLineExtend66
    oCurveByCurves68.Parameter("ConstructionSurface") = 0
    oCurveByCurves68.Evaluate

    Dim oSurfByLinearExtrusion69 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion69 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "069-SurfByLinearExtrusion")
    oSurfByLinearExtrusion69.Inputs("PlanarCrossSection").Add oCurveByCurves67
    oSurfByLinearExtrusion69.Inputs("ExtrusionLine").Add oLineByPoints45
    oSurfByLinearExtrusion69.Evaluate

    Dim oSurfByLinearExtrusion70 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion70 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "070-SurfByLinearExtrusion")
    oSurfByLinearExtrusion70.Inputs("PlanarCrossSection").Add oCurveByCurves68
    oSurfByLinearExtrusion70.Inputs("ExtrusionLine").Add oLineByPoints45
    oSurfByLinearExtrusion70.Evaluate

    Dim oLineExtend71 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend71 = oGCFactory.CreateEntity("LineExtend", pPOM, "071-LineExtend")
    oLineExtend71.Inputs("Line").Add oLineByPoints37
    oLineExtend71.Inputs("Curve1").Add oCurveByProjection29
    oLineExtend71.Inputs("Curve2").Add oLineExtend57
    oLineExtend71.Parameter("Context1") = 1
    oLineExtend71.Parameter("Context2") = 3
    oLineExtend71.Evaluate

    Dim oArcByFillet72 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet72 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "072-ArcByFillet")
    oArcByFillet72.Inputs("Curve1").Add oLineExtend57
    oArcByFillet72.Inputs("Curve2").Add oLineExtend71
    oArcByFillet72.Parameter("Radius") = oParamDistConstant5
    oArcByFillet72.Evaluate

    Dim oLineExtend73 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend73 = oGCFactory.CreateEntity("LineExtend", pPOM, "073-LineExtend")
    oLineExtend73.Inputs("Line").Add oLineExtend57
    oLineExtend73.Inputs("Curve1").Add oArcByFillet72
    oLineExtend73.Parameter("Context1") = 1
    oLineExtend73.Parameter("Context2") = 3
    oLineExtend73.Evaluate

    Dim oLineExtend74 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend74 = oGCFactory.CreateEntity("LineExtend", pPOM, "074-LineExtend")
    oLineExtend74.Inputs("Line").Add oLineExtend71
    oLineExtend74.Inputs("Curve1").Add oArcByFillet72
    oLineExtend74.Inputs("Curve2").Add oCurveByProjection29
    oLineExtend74.Parameter("Context1") = 1
    oLineExtend74.Parameter("Context2") = 3
    oLineExtend74.Evaluate

    Dim oCurveByCurves75 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves75 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "075-CurveByCurves")
    oCurveByCurves75.Inputs("Curves").Add oLineExtend73
    oCurveByCurves75.Inputs("Curves").Add oArcByFillet72
    oCurveByCurves75.Inputs("Curves").Add oLineExtend74
    oCurveByCurves75.Parameter("ConstructionSurface") = 0
    oCurveByCurves75.Evaluate

    Dim oSurfByLinearExtrusion76 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion76 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "076-SurfByLinearExtrusion")
    oSurfByLinearExtrusion76.Inputs("PlanarCrossSection").Add oCurveByCurves75
    oSurfByLinearExtrusion76.Inputs("ExtrusionLine").Add oLineByPoints45
    oSurfByLinearExtrusion76.Evaluate
   
    Dim oExtractPort77 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPort77 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "077-ExtractPorts")
    oExtractPort77.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "1"
    oExtractPort77.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1)
    oExtractPort77.Parameter("Support") = 2 'pGeometricConstruction.Parameter("Direction")
    oExtractPort77.Parameter("Debug") = 0
    oExtractPort77.Evaluate
    pGeometricConstruction.ControlledInputs("Port1_1_4").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_4").Add oExtractPort77.ControlledInputs("AxisPort").Item("1")
   
    Dim oCSByPlane78 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane78 = oGCFactory.CreateEntity("CSByPlane", pPOM, "078-CSByPlane")
    oCSByPlane78.Inputs("Plane").Add oExtractPort77.Output("Support", 1)
    oCSByPlane78.Inputs("Point").Add oOrderAxes21.Output("Node", 1)
    oCSByPlane78.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("CoordinateSystem", 1) = oCSByPlane78.Output
    oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
    oGCMacro.Output("Support", 1) = oExtractPort77.Output("Support", 1)
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion76.Output
    oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion55.Output
    oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion70.Output
    oGCMacro.Output("Boundary", 4) = oSurfByLinearExtrusion69.Output
    oGCMacro.Output("Boundary", 5) = oSurfFromGType18.Output
    oGCMacro.Output("Boundary", 6) = oSurfFromGType21.Output
    oGCMacro.Output("Boundary", 7) = oSurfFromGType25.Output
    oGCMacro.Output("Boundary", 8) = oSurfFromGType30.Output
End Sub
