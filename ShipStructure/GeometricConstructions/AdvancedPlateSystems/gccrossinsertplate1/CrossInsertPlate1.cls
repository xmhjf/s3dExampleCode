VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CrossInsertPlate1"
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
    Let Source = "GCAPSWithTrimBack3.CrossInsertPlate1"
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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select a first continous Member (in 1 or 2 pieces)", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 2, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersCheckIntersectionPlate;MembersSplitByPlate;MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile2", "Select a second intersecting Member (in 1 or 2 pieces)", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 2, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile2", "BehaviorOfPrimaryProfile2", 1, 0, 0, 0, 0, "MembersCheckIntersectionPlate;MembersSplitByPlate;MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_2", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_3")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Width", "Width", 8, 1, 59, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 0.2)
    Call pGeometricConstructionDefinition.AddParameter("Direction", "Direction", 4, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Direction", "Negative", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Direction", "Positive", 2)
    
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
    If pGeometricConstruction.Inputs("PrimaryProfile2").Count = 2 Then
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(2), "4"
    End If
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
    If pGeometricConstruction.Inputs("PrimaryProfile2").Count = 2 Then
        pGeometricConstruction.ControlledInputs("Port1_1_2").Add oOrderAxes21.ControlledInputs("Port1").Item("4")
    End If
   
    'Filed DI-CP-233119 on Geometric Constructions to createOrder Axes with Stiffener Part Axis
    'once the DI is done then below fix can be removed
    'For Stiffeners Update the OrderAxes with Stiffener Parts Axis
    If TypeOf pGeometricConstruction.Inputs("PrimaryProfile1").Item(1) Is IJStiffener Then
        UpdateOrderAxeswithStiffenerParts oOrderAxes21, pGeometricConstruction
    End If

    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = pGeometricConstruction.Parameter("Length")
    oParamDistConstant2.Evaluate

    Dim oParamDouble3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble3 = oGCFactory.CreateEntity("ParamDouble", pPOM, "003-ParamDouble")
    oParamDouble3.Parameter("Value") = 2#
    oParamDouble3.Evaluate

    Dim oPlaneFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS4 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "004-PlaneFromCS")
    oPlaneFromCS4.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS4.Parameter("LookingAxis") = 3
    oPlaneFromCS4.Parameter("Offset") = 0#
    oPlaneFromCS4.Parameter("Range") = 3#
    oPlaneFromCS4.Evaluate

    Dim oParamDouble5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDouble5 = oGCFactory.CreateEntity("ParamDouble", pPOM, "005-ParamDouble")
    oParamDouble5.Parameter("Value") = pGeometricConstruction.Parameter("Radius")
    oParamDouble5.Evaluate

    Dim oParamDistFunctionWith2Args6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args6 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "006-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args6.Inputs("Arg1").Add oParamDistConstant2
    oParamDistFunctionWith2Args6.Inputs("Arg2").Add oParamDouble3
    oParamDistFunctionWith2Args6.Parameter("Value") = pGeometricConstruction.Parameter("Width")
    oParamDistFunctionWith2Args6.Parameter("Operation") = 4
    oParamDistFunctionWith2Args6.Evaluate

    Dim oParamDistFunctionWith1Arg7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg7 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "007-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg7.Inputs("Arg").Add oParamDistConstant2
    oParamDistFunctionWith1Arg7.Parameter("Value") = -1#
    oParamDistFunctionWith1Arg7.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg7.Evaluate

    Dim oPlaneFromCS8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS8 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "008-PlaneFromCS")
    oPlaneFromCS8.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS8.Parameter("LookingAxis") = 1
    oPlaneFromCS8.Parameter("Offset") = oParamDistConstant2
    oPlaneFromCS8.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS8.Evaluate

    Dim oPlaneFromCS9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS9 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "009-PlaneFromCS")
    oPlaneFromCS9.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS9.Parameter("LookingAxis") = 2
    oPlaneFromCS9.Parameter("Offset") = oParamDistConstant2
    oPlaneFromCS9.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS9.Evaluate

    Dim oPointFromCS10 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS10 = oGCFactory.CreateEntity("PointFromCS", pPOM, "010-PointFromCS")
    oPointFromCS10.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS10.Parameter("X") = 0#
    oPointFromCS10.Parameter("Y") = 0#
    oPointFromCS10.Parameter("Z") = pGeometricConstruction.Parameter("Length") + 1
    oPointFromCS10.Evaluate

    Dim oParamDistFunctionWith1Arg11 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg11 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "011-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg11.Inputs("Arg").Add oParamDistFunctionWith2Args6
    oParamDistFunctionWith1Arg11.Parameter("Value") = -0.5
    oParamDistFunctionWith1Arg11.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg11.Evaluate

    Dim oPlaneFromCS12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS12 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "012-PlaneFromCS")
    oPlaneFromCS12.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS12.Parameter("LookingAxis") = 1
    oPlaneFromCS12.Parameter("Offset") = oParamDistFunctionWith1Arg7
    oPlaneFromCS12.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS12.Evaluate

    Dim oPlaneFromCS13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS13 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "013-PlaneFromCS")
    oPlaneFromCS13.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS13.Parameter("LookingAxis") = 2
    oPlaneFromCS13.Parameter("Offset") = oParamDistFunctionWith1Arg7
    oPlaneFromCS13.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS13.Evaluate

    Dim oPointFromCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS14 = oGCFactory.CreateEntity("PointFromCS", pPOM, "014-PointFromCS")
    oPointFromCS14.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS14.Parameter("X") = oParamDistFunctionWith2Args6
    oPointFromCS14.Parameter("Y") = 0#
    oPointFromCS14.Parameter("Z") = 0#
    oPointFromCS14.Evaluate

    Dim oPointFromCS15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS15 = oGCFactory.CreateEntity("PointFromCS", pPOM, "015-PointFromCS")
    oPointFromCS15.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS15.Parameter("X") = 0#
    oPointFromCS15.Parameter("Y") = oParamDistFunctionWith2Args6
    oPointFromCS15.Parameter("Z") = 0#
    oPointFromCS15.Evaluate

    Dim oPointFromCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS16 = oGCFactory.CreateEntity("PointFromCS", pPOM, "016-PointFromCS")
    oPointFromCS16.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS16.Parameter("X") = 0#
    oPointFromCS16.Parameter("Y") = 0#
    oPointFromCS16.Parameter("Z") = -(pGeometricConstruction.Parameter("Length") + 1)
    oPointFromCS16.Evaluate

    Dim oPointFromCS17 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS17 = oGCFactory.CreateEntity("PointFromCS", pPOM, "017-PointFromCS")
    oPointFromCS17.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS17.Parameter("X") = oParamDistFunctionWith1Arg11
    oPointFromCS17.Parameter("Y") = 0#
    oPointFromCS17.Parameter("Z") = 0#
    oPointFromCS17.Evaluate

    Dim oPointFromCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS18 = oGCFactory.CreateEntity("PointFromCS", pPOM, "018-PointFromCS")
    oPointFromCS18.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS18.Parameter("X") = 0#
    oPointFromCS18.Parameter("Y") = oParamDistFunctionWith1Arg11
    oPointFromCS18.Parameter("Z") = 0#
    oPointFromCS18.Evaluate

    Dim oSurfFromGType19 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType19 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "019-SurfFromGType")
    oSurfFromGType19.Inputs("Surface").Add oPlaneFromCS8
    oSurfFromGType19.Evaluate

    Dim oPointByProjectOnSurf20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf20 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "020-PointByProjectOnSurf")
    oPointByProjectOnSurf20.Inputs("Point").Add oPointFromCS18
    oPointByProjectOnSurf20.Inputs("Surface").Add oPlaneFromCS8
    oPointByProjectOnSurf20.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf20.Evaluate

    Dim oPointByProjectOnSurf21 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf21 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "021-PointByProjectOnSurf")
    oPointByProjectOnSurf21.Inputs("Point").Add oPointFromCS15
    oPointByProjectOnSurf21.Inputs("Surface").Add oPlaneFromCS8
    oPointByProjectOnSurf21.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf21.Evaluate

    Dim oSurfFromGType22 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType22 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "022-SurfFromGType")
    oSurfFromGType22.Inputs("Surface").Add oPlaneFromCS12
    oSurfFromGType22.Evaluate

    Dim oPointByProjectOnSurf23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf23 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "023-PointByProjectOnSurf")
    oPointByProjectOnSurf23.Inputs("Point").Add oPointFromCS18
    oPointByProjectOnSurf23.Inputs("Surface").Add oPlaneFromCS12
    oPointByProjectOnSurf23.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf23.Evaluate

    Dim oPointByProjectOnSurf24 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf24 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "024-PointByProjectOnSurf")
    oPointByProjectOnSurf24.Inputs("Point").Add oPointFromCS15
    oPointByProjectOnSurf24.Inputs("Surface").Add oPlaneFromCS12
    oPointByProjectOnSurf24.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf24.Evaluate

    Dim oSurfFromGType25 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType25 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "025-SurfFromGType")
    oSurfFromGType25.Inputs("Surface").Add oPlaneFromCS13
    oSurfFromGType25.Evaluate

    Dim oPointByProjectOnSurf26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf26 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "026-PointByProjectOnSurf")
    oPointByProjectOnSurf26.Inputs("Point").Add oPointFromCS14
    oPointByProjectOnSurf26.Inputs("Surface").Add oPlaneFromCS13
    oPointByProjectOnSurf26.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf26.Evaluate

    Dim oPointByProjectOnSurf27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf27 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "027-PointByProjectOnSurf")
    oPointByProjectOnSurf27.Inputs("Point").Add oPointFromCS17
    oPointByProjectOnSurf27.Inputs("Surface").Add oPlaneFromCS13
    oPointByProjectOnSurf27.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf27.Evaluate

    Dim oSurfFromGType28 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType28 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "028-SurfFromGType")
    oSurfFromGType28.Inputs("Surface").Add oPlaneFromCS9
    oSurfFromGType28.Evaluate

    Dim oPointByProjectOnSurf29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf29 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "029-PointByProjectOnSurf")
    oPointByProjectOnSurf29.Inputs("Point").Add oPointFromCS17
    oPointByProjectOnSurf29.Inputs("Surface").Add oPlaneFromCS9
    oPointByProjectOnSurf29.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf29.Evaluate

    Dim oPointByProjectOnSurf30 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf30 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "030-PointByProjectOnSurf")
    oPointByProjectOnSurf30.Inputs("Point").Add oPointFromCS14
    oPointByProjectOnSurf30.Inputs("Surface").Add oPlaneFromCS9
    oPointByProjectOnSurf30.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf30.Evaluate

    Dim oLineByPoints31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints31 = oGCFactory.CreateEntity("LineByPoints", pPOM, "031-LineByPoints")
    oLineByPoints31.Inputs("StartPoint").Add oPointByProjectOnSurf26
    oLineByPoints31.Inputs("EndPoint").Add oPointFromCS14
    oLineByPoints31.Evaluate

    Dim oLineByPoints32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints32 = oGCFactory.CreateEntity("LineByPoints", pPOM, "032-LineByPoints")
    oLineByPoints32.Inputs("StartPoint").Add oPointFromCS14
    oLineByPoints32.Inputs("EndPoint").Add oPointByProjectOnSurf30
    oLineByPoints32.Evaluate

    Dim oLineByPoints33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints33 = oGCFactory.CreateEntity("LineByPoints", pPOM, "033-LineByPoints")
    oLineByPoints33.Inputs("StartPoint").Add oPointByProjectOnSurf27
    oLineByPoints33.Inputs("EndPoint").Add oPointFromCS17
    oLineByPoints33.Evaluate

    Dim oLineByPoints34 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints34 = oGCFactory.CreateEntity("LineByPoints", pPOM, "034-LineByPoints")
    oLineByPoints34.Inputs("StartPoint").Add oPointFromCS17
    oLineByPoints34.Inputs("EndPoint").Add oPointByProjectOnSurf29
    oLineByPoints34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointByProjectOnSurf24
    oLineByPoints35.Inputs("EndPoint").Add oPointFromCS15
    oLineByPoints35.Evaluate

    Dim oLineByPoints36 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints36 = oGCFactory.CreateEntity("LineByPoints", pPOM, "036-LineByPoints")
    oLineByPoints36.Inputs("StartPoint").Add oPointByProjectOnSurf21
    oLineByPoints36.Inputs("EndPoint").Add oPointFromCS15
    oLineByPoints36.Evaluate

    Dim oLineByPoints37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints37 = oGCFactory.CreateEntity("LineByPoints", pPOM, "037-LineByPoints")
    oLineByPoints37.Inputs("StartPoint").Add oPointByProjectOnSurf20
    oLineByPoints37.Inputs("EndPoint").Add oPointFromCS18
    oLineByPoints37.Evaluate

    Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "038-LineByPoints")
    oLineByPoints38.Inputs("StartPoint").Add oPointByProjectOnSurf23
    oLineByPoints38.Inputs("EndPoint").Add oPointFromCS18
    oLineByPoints38.Evaluate

    Dim oLineByPoints39 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints39 = oGCFactory.CreateEntity("LineByPoints", pPOM, "039-LineByPoints")
    oLineByPoints39.Inputs("StartPoint").Add oPointByProjectOnSurf20
    oLineByPoints39.Inputs("EndPoint").Add oPointByProjectOnSurf26
    oLineByPoints39.Evaluate

    Dim oLineByPoints40 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints40 = oGCFactory.CreateEntity("LineByPoints", pPOM, "040-LineByPoints")
    oLineByPoints40.Inputs("StartPoint").Add oPointByProjectOnSurf23
    oLineByPoints40.Inputs("EndPoint").Add oPointByProjectOnSurf27
    oLineByPoints40.Evaluate

    Dim oLineByPoints41 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints41 = oGCFactory.CreateEntity("LineByPoints", pPOM, "041-LineByPoints")
    oLineByPoints41.Inputs("StartPoint").Add oPointByProjectOnSurf29
    oLineByPoints41.Inputs("EndPoint").Add oPointByProjectOnSurf24
    oLineByPoints41.Evaluate

    Dim oLineByPoints42 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints42 = oGCFactory.CreateEntity("LineByPoints", pPOM, "042-LineByPoints")
    oLineByPoints42.Inputs("StartPoint").Add oPointByProjectOnSurf30
    oLineByPoints42.Inputs("EndPoint").Add oPointByProjectOnSurf21
    oLineByPoints42.Evaluate

    Dim oLineExtend43 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend43 = oGCFactory.CreateEntity("LineExtend", pPOM, "043-LineExtend")
    oLineExtend43.Inputs("Line").Add oLineByPoints37
    oLineExtend43.Inputs("Curve1").Add oLineByPoints39
    oLineExtend43.Inputs("Curve2").Add oLineByPoints31
    oLineExtend43.Parameter("Context1") = 1
    oLineExtend43.Parameter("Context2") = 3
    oLineExtend43.Evaluate

    Dim oLineExtend44 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend44 = oGCFactory.CreateEntity("LineExtend", pPOM, "044-LineExtend")
    oLineExtend44.Inputs("Line").Add oLineByPoints31
    oLineExtend44.Inputs("Curve1").Add oLineExtend43
    oLineExtend44.Inputs("Curve2").Add oLineByPoints39
    oLineExtend44.Parameter("Context1") = 1
    oLineExtend44.Parameter("Context2") = 3
    oLineExtend44.Evaluate

    Dim oLineExtend45 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend45 = oGCFactory.CreateEntity("LineExtend", pPOM, "045-LineExtend")
    oLineExtend45.Inputs("Line").Add oLineByPoints38
    oLineExtend45.Inputs("Curve1").Add oLineByPoints33
    oLineExtend45.Inputs("Curve2").Add oLineByPoints40
    oLineExtend45.Parameter("Context1") = 1
    oLineExtend45.Parameter("Context2") = 3
    oLineExtend45.Evaluate

    Dim oLineExtend46 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend46 = oGCFactory.CreateEntity("LineExtend", pPOM, "046-LineExtend")
    oLineExtend46.Inputs("Line").Add oLineByPoints33
    oLineExtend46.Inputs("Curve1").Add oLineExtend45
    oLineExtend46.Inputs("Curve2").Add oLineByPoints40
    oLineExtend46.Parameter("Context1") = 1
    oLineExtend46.Parameter("Context2") = 3
    oLineExtend46.Evaluate

    Dim oLineExtend47 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend47 = oGCFactory.CreateEntity("LineExtend", pPOM, "047-LineExtend")
    oLineExtend47.Inputs("Line").Add oLineByPoints35
    oLineExtend47.Inputs("Curve1").Add oLineByPoints34
    oLineExtend47.Inputs("Curve2").Add oLineByPoints41
    oLineExtend47.Parameter("Context1") = 1
    oLineExtend47.Parameter("Context2") = 3
    oLineExtend47.Evaluate

    Dim oLineExtend48 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend48 = oGCFactory.CreateEntity("LineExtend", pPOM, "048-LineExtend")
    oLineExtend48.Inputs("Line").Add oLineByPoints34
    oLineExtend48.Inputs("Curve1").Add oLineExtend47
    oLineExtend48.Inputs("Curve2").Add oLineByPoints41
    oLineExtend48.Parameter("Context1") = 1
    oLineExtend48.Parameter("Context2") = 3
    oLineExtend48.Evaluate

    Dim oLineExtend49 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend49 = oGCFactory.CreateEntity("LineExtend", pPOM, "049-LineExtend")
    oLineExtend49.Inputs("Line").Add oLineByPoints36
    oLineExtend49.Inputs("Curve1").Add oLineByPoints32
    oLineExtend49.Inputs("Curve2").Add oLineByPoints42
    oLineExtend49.Parameter("Context1") = 1
    oLineExtend49.Parameter("Context2") = 3
    oLineExtend49.Evaluate

    Dim oLineExtend50 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend50 = oGCFactory.CreateEntity("LineExtend", pPOM, "050-LineExtend")
    oLineExtend50.Inputs("Line").Add oLineByPoints32
    oLineExtend50.Inputs("Curve1").Add oLineExtend49
    oLineExtend50.Inputs("Curve2").Add oLineByPoints42
    oLineExtend50.Parameter("Context1") = 1
    oLineExtend50.Parameter("Context2") = 3
    oLineExtend50.Evaluate

    Dim oArcByFillet51 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet51 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "051-ArcByFillet")
    oArcByFillet51.Inputs("Curve1").Add oLineExtend43
    oArcByFillet51.Inputs("Curve2").Add oLineExtend44
    oArcByFillet51.Parameter("Radius") = oParamDouble5
    oArcByFillet51.Evaluate

    Dim oLineExtend52 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend52 = oGCFactory.CreateEntity("LineExtend", pPOM, "052-LineExtend")
    oLineExtend52.Inputs("Line").Add oLineExtend44
    oLineExtend52.Inputs("Curve1").Add oArcByFillet51
    oLineExtend52.Parameter("Context1") = 1
    oLineExtend52.Parameter("Context2") = 3
    oLineExtend52.Evaluate

    Dim oArcByFillet53 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet53 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "053-ArcByFillet")
    oArcByFillet53.Inputs("Curve1").Add oLineExtend46
    oArcByFillet53.Inputs("Curve2").Add oLineExtend45
    oArcByFillet53.Parameter("Radius") = oParamDouble5
    oArcByFillet53.Evaluate

    Dim oLineExtend54 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend54 = oGCFactory.CreateEntity("LineExtend", pPOM, "054-LineExtend")
    oLineExtend54.Inputs("Line").Add oLineExtend45
    oLineExtend54.Inputs("Curve1").Add oArcByFillet53
    oLineExtend54.Parameter("Context1") = 1
    oLineExtend54.Parameter("Context2") = 3
    oLineExtend54.Evaluate

    Dim oLineExtend55 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend55 = oGCFactory.CreateEntity("LineExtend", pPOM, "055-LineExtend")
    oLineExtend55.Inputs("Line").Add oLineExtend46
    oLineExtend55.Inputs("Curve1").Add oArcByFillet53
    oLineExtend55.Parameter("Context1") = 1
    oLineExtend55.Parameter("Context2") = 3
    oLineExtend55.Evaluate

    Dim oArcByFillet56 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet56 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "056-ArcByFillet")
    oArcByFillet56.Inputs("Curve1").Add oLineExtend48
    oArcByFillet56.Inputs("Curve2").Add oLineExtend47
    oArcByFillet56.Parameter("Radius") = oParamDouble5
    oArcByFillet56.Evaluate

    Dim oLineExtend57 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend57 = oGCFactory.CreateEntity("LineExtend", pPOM, "057-LineExtend")
    oLineExtend57.Inputs("Line").Add oLineExtend47
    oLineExtend57.Inputs("Curve1").Add oArcByFillet56
    oLineExtend57.Parameter("Context1") = 1
    oLineExtend57.Parameter("Context2") = 3
    oLineExtend57.Evaluate

    Dim oLineExtend58 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend58 = oGCFactory.CreateEntity("LineExtend", pPOM, "058-LineExtend")
    oLineExtend58.Inputs("Line").Add oLineExtend48
    oLineExtend58.Inputs("Curve1").Add oArcByFillet56
    oLineExtend58.Parameter("Context1") = 1
    oLineExtend58.Parameter("Context2") = 3
    oLineExtend58.Evaluate

    Dim oArcByFillet59 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet59 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "059-ArcByFillet")
    oArcByFillet59.Inputs("Curve1").Add oLineExtend49
    oArcByFillet59.Inputs("Curve2").Add oLineExtend50
    oArcByFillet59.Parameter("Radius") = oParamDouble5
    oArcByFillet59.Evaluate

    Dim oLineExtend60 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend60 = oGCFactory.CreateEntity("LineExtend", pPOM, "060-LineExtend")
    oLineExtend60.Inputs("Line").Add oLineExtend49
    oLineExtend60.Inputs("Curve1").Add oArcByFillet59
    oLineExtend60.Parameter("Context1") = 1
    oLineExtend60.Parameter("Context2") = 3
    oLineExtend60.Evaluate

    Dim oLineExtend61 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend61 = oGCFactory.CreateEntity("LineExtend", pPOM, "061-LineExtend")
    oLineExtend61.Inputs("Line").Add oLineExtend50
    oLineExtend61.Inputs("Curve1").Add oArcByFillet59
    oLineExtend61.Parameter("Context1") = 1
    oLineExtend61.Parameter("Context2") = 3
    oLineExtend61.Evaluate

    Dim oCurveByCurves62 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves62 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "062-CurveByCurves")
    oCurveByCurves62.Inputs("Curves").Add oLineExtend55
    oCurveByCurves62.Inputs("Curves").Add oArcByFillet53
    oCurveByCurves62.Inputs("Curves").Add oLineExtend54
    oCurveByCurves62.Parameter("ConstructionSurface") = 0
    oCurveByCurves62.Evaluate

    Dim oCurveByCurves63 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves63 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "063-CurveByCurves")
    oCurveByCurves63.Inputs("Curves").Add oLineExtend58
    oCurveByCurves63.Inputs("Curves").Add oArcByFillet56
    oCurveByCurves63.Inputs("Curves").Add oLineExtend57
    oCurveByCurves63.Parameter("ConstructionSurface") = 0
    oCurveByCurves63.Evaluate

    Dim oCurveByCurves64 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves64 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "064-CurveByCurves")
    oCurveByCurves64.Inputs("Curves").Add oLineExtend60
    oCurveByCurves64.Inputs("Curves").Add oArcByFillet59
    oCurveByCurves64.Inputs("Curves").Add oLineExtend61
    oCurveByCurves64.Parameter("ConstructionSurface") = 0
    oCurveByCurves64.Evaluate

    Dim oLineByPoints65 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints65 = oGCFactory.CreateEntity("LineByPoints", pPOM, "065-LineByPoints")
    oLineByPoints65.Inputs("StartPoint").Add oPointFromCS16
    oLineByPoints65.Inputs("EndPoint").Add oPointFromCS10
    oLineByPoints65.Evaluate

    Dim oSurfByLinearExtrusion66 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion66 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "066-SurfByLinearExtrusion")
    oSurfByLinearExtrusion66.Inputs("PlanarCrossSection").Add oCurveByCurves62
    oSurfByLinearExtrusion66.Inputs("ExtrusionLine").Add oLineByPoints65
    oSurfByLinearExtrusion66.Evaluate

    Dim oSurfByLinearExtrusion67 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion67 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "067-SurfByLinearExtrusion")
    oSurfByLinearExtrusion67.Inputs("PlanarCrossSection").Add oCurveByCurves63
    oSurfByLinearExtrusion67.Inputs("ExtrusionLine").Add oLineByPoints65
    oSurfByLinearExtrusion67.Evaluate

    Dim oSurfByLinearExtrusion68 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion68 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "068-SurfByLinearExtrusion")
    oSurfByLinearExtrusion68.Inputs("PlanarCrossSection").Add oCurveByCurves64
    oSurfByLinearExtrusion68.Inputs("ExtrusionLine").Add oLineByPoints65
    oSurfByLinearExtrusion68.Evaluate
   
    Dim oLineExtend70 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend70 = oGCFactory.CreateEntity("LineExtend", pPOM, "070-LineExtend")
    oLineExtend70.Inputs("Line").Add oLineExtend43
    oLineExtend70.Inputs("Curve1").Add oArcByFillet51
    oLineExtend70.Inputs("Curve2").Add oLineByPoints39
    oLineExtend70.Parameter("Context1") = 1
    oLineExtend70.Parameter("Context2") = 3
    oLineExtend70.Evaluate

    Dim oCurveByCurves71 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves71 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "071-CurveByCurves")
    oCurveByCurves71.Inputs("Curves").Add oLineExtend70
    oCurveByCurves71.Inputs("Curves").Add oArcByFillet51
    oCurveByCurves71.Inputs("Curves").Add oLineExtend52
    oCurveByCurves71.Parameter("ConstructionSurface") = 0
    oCurveByCurves71.Evaluate

    Dim oSurfByLinearExtrusion72 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion72 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "072-SurfByLinearExtrusion")
    oSurfByLinearExtrusion72.Inputs("PlanarCrossSection").Add oCurveByCurves71
    oSurfByLinearExtrusion72.Inputs("ExtrusionLine").Add oLineByPoints65
    oSurfByLinearExtrusion72.Evaluate
    
    Dim oExtractPort73 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPort73 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "073-ExtractPorts")
    oExtractPort73.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "1"
    oExtractPort73.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1)
    oExtractPort73.Parameter("Support") = 2 ' pGeometricConstruction.Parameter("Direction")
    oExtractPort73.Parameter("Debug") = 0
    oExtractPort73.Evaluate
    pGeometricConstruction.ControlledInputs("Port1_1_3").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_3").Add oExtractPort73.ControlledInputs("AxisPort").Item("1")

    Dim oCSByPlane74 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane74 = oGCFactory.CreateEntity("CSByPlane", pPOM, "074-CSByPlane")
    oCSByPlane74.Inputs("Plane").Add oExtractPort73.Output("Support", 1)
    oCSByPlane74.Inputs("Point").Add oOrderAxes21.Output("Node", 1)
    oCSByPlane74.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion72.Output
    oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion66.Output
    oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion67.Output
    oGCMacro.Output("Boundary", 4) = oSurfByLinearExtrusion68.Output
    oGCMacro.Output("Boundary", 5) = oSurfFromGType19.Output
    oGCMacro.Output("Boundary", 6) = oSurfFromGType25.Output
    oGCMacro.Output("Boundary", 7) = oSurfFromGType22.Output
    oGCMacro.Output("Boundary", 8) = oSurfFromGType28.Output
    oGCMacro.Output("Support", 1) = oExtractPort73.Output("Support", 1)
    oGCMacro.Output("CoordinateSystem", 1) = oCSByPlane74.Output
    oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
End Sub
