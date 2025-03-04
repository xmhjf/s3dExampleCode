VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TInsertPlate"
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
    Let Source = "GCAPSWithTrimBack3.TInsertPlate"
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
    pGC.RaiseError (Err.Number)
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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile2", "Select a supported Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, _
                                                   "IJSplitNotify IJStructSplit IJDProfileMoldedConventions IJStiffenerAttachmentMethod ISPSMemberSystemXSectionNotify ISPSMemberSystemSuppingNotify1")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile2", "BehaviorOfPrimaryProfile2", 1, 0, 0, 0, 0, "MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_2", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_3")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem2")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem2")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Width", "Width", 8, 1, 59, 0, 0, 0.5)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 0.1)
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

    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = 2#
    oParamDistConstant3.Evaluate

    Dim oParamDistConstant4 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant4 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "004-ParamDistConstant")
    oParamDistConstant4.Parameter("Value") = pGeometricConstruction.Parameter("Width")
    oParamDistConstant4.Evaluate

    Dim oParamDistConstant5 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant5 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "005-ParamDistConstant")
    oParamDistConstant5.Parameter("Value") = 0.33
    oParamDistConstant5.Evaluate

    Dim oPlaneFromCS6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS6 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "006-PlaneFromCS")
    oPlaneFromCS6.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS6.Parameter("LookingAxis") = 3
    oPlaneFromCS6.Parameter("Offset") = 0#
    oPlaneFromCS6.Parameter("Range") = 3#
    oPlaneFromCS6.Evaluate

    Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "007-ParamDistConstant")
    oParamDistConstant7.Parameter("Value") = pGeometricConstruction.Parameter("Radius")
    oParamDistConstant7.Evaluate

    Dim oParamDistFunctionWith2Args8 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args8 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "008-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args8.Inputs("Arg1").Add oParamDistConstant2
    oParamDistFunctionWith2Args8.Inputs("Arg2").Add oParamDistConstant3
    oParamDistFunctionWith2Args8.Parameter("Value") = 0.5
    oParamDistFunctionWith2Args8.Parameter("Operation") = 4
    oParamDistFunctionWith2Args8.Evaluate

    Dim oParamDistFunctionWith1Arg9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg9 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "009-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg9.Inputs("Arg").Add oParamDistConstant2
    oParamDistFunctionWith1Arg9.Parameter("Value") = -1#
    oParamDistFunctionWith1Arg9.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg9.Evaluate

    Dim oParamDistFunctionWith2Args10 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args10 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "010-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args10.Inputs("Arg1").Add oParamDistConstant2
    oParamDistFunctionWith2Args10.Inputs("Arg2").Add oParamDistConstant5
    oParamDistFunctionWith2Args10.Parameter("Value") = 0.33
    oParamDistFunctionWith2Args10.Parameter("Operation") = 3
    oParamDistFunctionWith2Args10.Evaluate

    Dim oPlaneFromCS11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS11 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "011-PlaneFromCS")
    oPlaneFromCS11.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS11.Parameter("LookingAxis") = 2
    oPlaneFromCS11.Parameter("Offset") = oParamDistConstant2
    oPlaneFromCS11.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS11.Evaluate

    Dim oPointFromCS12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS12 = oGCFactory.CreateEntity("PointFromCS", pPOM, "012-PointFromCS")
    oPointFromCS12.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS12.Parameter("X") = 0#
    oPointFromCS12.Parameter("Y") = 0#
    oPointFromCS12.Parameter("Z") = pGeometricConstruction.Parameter("Length") + 1
    oPointFromCS12.Evaluate

    Dim oParamDistFunctionWith2Args13 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args13 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "013-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args13.Inputs("Arg1").Add oParamDistConstant4
    oParamDistFunctionWith2Args13.Inputs("Arg2").Add oParamDistConstant3
    oParamDistFunctionWith2Args13.Parameter("Value") = 0.25
    oParamDistFunctionWith2Args13.Parameter("Operation") = 4
    oParamDistFunctionWith2Args13.Evaluate

    Dim oPlaneFromCS14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS14 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "014-PlaneFromCS")
    oPlaneFromCS14.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS14.Parameter("LookingAxis") = 1
    oPlaneFromCS14.Parameter("Offset") = oParamDistFunctionWith1Arg9
    oPlaneFromCS14.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS14.Evaluate

    Dim oParamDistFunctionWith1Arg15 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith1Arg15 = oGCFactory.CreateEntity("ParamDistFunctionWith1Arg", pPOM, "015-ParamDistFunctionWith1Arg")
    oParamDistFunctionWith1Arg15.Inputs("Arg").Add oParamDistConstant4
    oParamDistFunctionWith1Arg15.Parameter("Value") = -0.5
    oParamDistFunctionWith1Arg15.Parameter("Operation") = 2
    oParamDistFunctionWith1Arg15.Evaluate

    Dim oPointFromCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS16 = oGCFactory.CreateEntity("PointFromCS", pPOM, "016-PointFromCS")
    oPointFromCS16.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS16.Parameter("X") = oParamDistConstant4
    oPointFromCS16.Parameter("Y") = 0#
    oPointFromCS16.Parameter("Z") = 0#
    oPointFromCS16.Evaluate

    Dim oPointFromCS17 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS17 = oGCFactory.CreateEntity("PointFromCS", pPOM, "017-PointFromCS")
    oPointFromCS17.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS17.Parameter("X") = 0#
    oPointFromCS17.Parameter("Y") = oParamDistConstant4
    oPointFromCS17.Parameter("Z") = 0#
    oPointFromCS17.Evaluate

    Dim oPointFromCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS18 = oGCFactory.CreateEntity("PointFromCS", pPOM, "018-PointFromCS")
    oPointFromCS18.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS18.Parameter("X") = oParamDistFunctionWith1Arg15
    oPointFromCS18.Parameter("Y") = 0#
    oPointFromCS18.Parameter("Z") = 0#
    oPointFromCS18.Evaluate

    Dim oPointFromCS19 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS19 = oGCFactory.CreateEntity("PointFromCS", pPOM, "019-PointFromCS")
    oPointFromCS19.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPointFromCS19.Parameter("X") = 0#
    oPointFromCS19.Parameter("Y") = 0#
    oPointFromCS19.Parameter("Z") = -(pGeometricConstruction.Parameter("Length") + 1)
    oPointFromCS19.Evaluate

    Dim oPointByProjectOnSurf20 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf20 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "020-PointByProjectOnSurf")
    oPointByProjectOnSurf20.Inputs("Point").Add oPointFromCS17
    oPointByProjectOnSurf20.Inputs("Surface").Add oPlaneFromCS14
    oPointByProjectOnSurf20.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf20.Evaluate

    Dim oSurfFromGType21 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType21 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "021-SurfFromGType")
    oSurfFromGType21.Inputs("Surface").Add oPlaneFromCS14
    oSurfFromGType21.Evaluate

    Dim oPointByProjectOnSurf22 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf22 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "022-PointByProjectOnSurf")
    oPointByProjectOnSurf22.Inputs("Point").Add oPointFromCS16
    oPointByProjectOnSurf22.Inputs("Surface").Add oPlaneFromCS11
    oPointByProjectOnSurf22.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf22.Evaluate

    Dim oPointByProjectOnSurf23 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf23 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "023-PointByProjectOnSurf")
    oPointByProjectOnSurf23.Inputs("Point").Add oPointFromCS18
    oPointByProjectOnSurf23.Inputs("Surface").Add oPlaneFromCS11
    oPointByProjectOnSurf23.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf23.Evaluate

    Dim oSurfFromGType24 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType24 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "024-SurfFromGType")
    oSurfFromGType24.Inputs("Surface").Add oPlaneFromCS11
    oSurfFromGType24.Evaluate

    Dim oLineByPoints25 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints25 = oGCFactory.CreateEntity("LineByPoints", pPOM, "025-LineByPoints")
    oLineByPoints25.Inputs("StartPoint").Add oPointByProjectOnSurf22
    oLineByPoints25.Inputs("EndPoint").Add oPointFromCS16
    oLineByPoints25.Evaluate

    Dim oLineByPoints26 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints26 = oGCFactory.CreateEntity("LineByPoints", pPOM, "026-LineByPoints")
    oLineByPoints26.Inputs("StartPoint").Add oPointByProjectOnSurf23
    oLineByPoints26.Inputs("EndPoint").Add oPointFromCS18
    oLineByPoints26.Evaluate

    Dim oLineByPoints27 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints27 = oGCFactory.CreateEntity("LineByPoints", pPOM, "027-LineByPoints")
    oLineByPoints27.Inputs("StartPoint").Add oPointByProjectOnSurf20
    oLineByPoints27.Inputs("EndPoint").Add oPointFromCS17
    oLineByPoints27.Evaluate

    Dim oLineByPoints28 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints28 = oGCFactory.CreateEntity("LineByPoints", pPOM, "028-LineByPoints")
    oLineByPoints28.Inputs("StartPoint").Add oPointByProjectOnSurf23
    oLineByPoints28.Inputs("EndPoint").Add oPointByProjectOnSurf20
    oLineByPoints28.Evaluate

    Dim oLineExtend29 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend29 = oGCFactory.CreateEntity("LineExtend", pPOM, "029-LineExtend")
    oLineExtend29.Inputs("Line").Add oLineByPoints27
    oLineExtend29.Inputs("Curve1").Add oLineByPoints26
    oLineExtend29.Inputs("Curve2").Add oLineByPoints28
    oLineExtend29.Parameter("Context1") = 1
    oLineExtend29.Parameter("Context2") = 3
    oLineExtend29.Evaluate

    Dim oLineExtend30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend30 = oGCFactory.CreateEntity("LineExtend", pPOM, "030-LineExtend")
    oLineExtend30.Inputs("Line").Add oLineByPoints26
    oLineExtend30.Inputs("Curve1").Add oLineExtend29
    oLineExtend30.Inputs("Curve2").Add oLineByPoints28
    oLineExtend30.Parameter("Context1") = 1
    oLineExtend30.Parameter("Context2") = 3
    oLineExtend30.Evaluate

    Dim oArcByFillet31 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet31 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "031-ArcByFillet")
    oArcByFillet31.Inputs("Curve1").Add oLineExtend29
    oArcByFillet31.Inputs("Curve2").Add oLineExtend30
    oArcByFillet31.Parameter("Radius") = oParamDistConstant7
    oArcByFillet31.Evaluate

    Dim oLineExtend32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend32 = oGCFactory.CreateEntity("LineExtend", pPOM, "032-LineExtend")
    oLineExtend32.Inputs("Line").Add oLineExtend29
    oLineExtend32.Inputs("Curve1").Add oArcByFillet31
    oLineExtend32.Parameter("Context1") = 1
    oLineExtend32.Parameter("Context2") = 3
    oLineExtend32.Evaluate

    Dim oLineExtend33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend33 = oGCFactory.CreateEntity("LineExtend", pPOM, "033-LineExtend")
    oLineExtend33.Inputs("Line").Add oLineExtend30
    oLineExtend33.Inputs("Curve1").Add oArcByFillet31
    oLineExtend33.Parameter("Context1") = 1
    oLineExtend33.Parameter("Context2") = 3
    oLineExtend33.Evaluate

    Dim oCurveByCurves34 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves34 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "034-CurveByCurves")
    oCurveByCurves34.Inputs("Curves").Add oLineExtend33
    oCurveByCurves34.Inputs("Curves").Add oArcByFillet31
    oCurveByCurves34.Inputs("Curves").Add oLineExtend32
    oCurveByCurves34.Parameter("ConstructionSurface") = 0
    oCurveByCurves34.Evaluate

    Dim oLineByPoints35 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints35 = oGCFactory.CreateEntity("LineByPoints", pPOM, "035-LineByPoints")
    oLineByPoints35.Inputs("StartPoint").Add oPointFromCS12
    oLineByPoints35.Inputs("EndPoint").Add oPointFromCS19
    oLineByPoints35.Evaluate

   
    Dim oPlaneFromCS37 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS37 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "037-PlaneFromCS")
    oPlaneFromCS37.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS37.Parameter("LookingAxis") = 2
    oPlaneFromCS37.Parameter("Offset") = oParamDistFunctionWith1Arg15
    oPlaneFromCS37.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS37.Evaluate

    Dim oPlaneFromCS38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneFromCS38 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, "038-PlaneFromCS")
    oPlaneFromCS38.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1)
    oPlaneFromCS38.Parameter("LookingAxis") = 1
    oPlaneFromCS38.Parameter("Offset") = oParamDistConstant2
    oPlaneFromCS38.Parameter("Range") = pGeometricConstruction.Parameter("Length") + 1
    oPlaneFromCS38.Evaluate

    Dim oParamDistFunctionWith2Args39 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWith2Args39 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "039-ParamDistFunctionWith2Args")
    oParamDistFunctionWith2Args39.Inputs("Arg1").Add oParamDistFunctionWith1Arg15
    oParamDistFunctionWith2Args39.Inputs("Arg2").Add oParamDistConstant3
    oParamDistFunctionWith2Args39.Parameter("Value") = -0.25
    oParamDistFunctionWith2Args39.Parameter("Operation") = 4
    oParamDistFunctionWith2Args39.Evaluate

    Dim oSurfByLinearExtrusion40 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion40 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "040-SurfByLinearExtrusion")
    oSurfByLinearExtrusion40.Inputs("PlanarCrossSection").Add oCurveByCurves34
    oSurfByLinearExtrusion40.Inputs("ExtrusionLine").Add oLineByPoints35
    oSurfByLinearExtrusion40.Evaluate

    Dim oSurfFromGType41 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType41 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "041-SurfFromGType")
    oSurfFromGType41.Inputs("Surface").Add oPlaneFromCS37
    oSurfFromGType41.Evaluate

    Dim oPointByProjectOnSurf42 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf42 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "042-PointByProjectOnSurf")
    oPointByProjectOnSurf42.Inputs("Point").Add oPointFromCS17
    oPointByProjectOnSurf42.Inputs("Surface").Add oPlaneFromCS38
    oPointByProjectOnSurf42.Parameter("TrackFlag") = 1
    oPointByProjectOnSurf42.Evaluate

    Dim oSurfFromGType43 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfFromGType43 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "043-SurfFromGType")
    oSurfFromGType43.Inputs("Surface").Add oPlaneFromCS38
    oSurfFromGType43.Evaluate

    Dim oLineByPoints44 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints44 = oGCFactory.CreateEntity("LineByPoints", pPOM, "044-LineByPoints")
    oLineByPoints44.Inputs("StartPoint").Add oPointByProjectOnSurf42
    oLineByPoints44.Inputs("EndPoint").Add oPointFromCS17
    oLineByPoints44.Evaluate

    Dim oLineByPoints45 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints45 = oGCFactory.CreateEntity("LineByPoints", pPOM, "045-LineByPoints")
    oLineByPoints45.Inputs("StartPoint").Add oPointByProjectOnSurf22
    oLineByPoints45.Inputs("EndPoint").Add oPointByProjectOnSurf42
    oLineByPoints45.Evaluate

    Dim oLineExtend46 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend46 = oGCFactory.CreateEntity("LineExtend", pPOM, "046-LineExtend")
    oLineExtend46.Inputs("Line").Add oLineByPoints25
    oLineExtend46.Inputs("Curve1").Add oLineByPoints44
    oLineExtend46.Inputs("Curve2").Add oLineByPoints45
    oLineExtend46.Parameter("Context1") = 1
    oLineExtend46.Parameter("Context2") = 3
    oLineExtend46.Evaluate

    Dim oLineExtend47 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend47 = oGCFactory.CreateEntity("LineExtend", pPOM, "047-LineExtend")
    oLineExtend47.Inputs("Line").Add oLineByPoints44
    oLineExtend47.Inputs("Curve1").Add oLineExtend46
    oLineExtend47.Inputs("Curve2").Add oLineByPoints45
    oLineExtend47.Parameter("Context1") = 1
    oLineExtend47.Parameter("Context2") = 3
    oLineExtend47.Evaluate

    Dim oArcByFillet48 As SP3DGeometricConstruction.GeometricConstruction
    Set oArcByFillet48 = oGCFactory.CreateEntity("ArcByFillet", pPOM, "048-ArcByFillet")
    oArcByFillet48.Inputs("Curve1").Add oLineExtend46
    oArcByFillet48.Inputs("Curve2").Add oLineExtend47
    oArcByFillet48.Parameter("Radius") = oParamDistConstant7
    oArcByFillet48.Evaluate

    Dim oLineExtend49 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend49 = oGCFactory.CreateEntity("LineExtend", pPOM, "049-LineExtend")
    oLineExtend49.Inputs("Line").Add oLineExtend46
    oLineExtend49.Inputs("Curve1").Add oArcByFillet48
    oLineExtend49.Parameter("Context1") = 1
    oLineExtend49.Parameter("Context2") = 3
    oLineExtend49.Evaluate

    Dim oLineExtend50 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineExtend50 = oGCFactory.CreateEntity("LineExtend", pPOM, "050-LineExtend")
    oLineExtend50.Inputs("Line").Add oLineExtend47
    oLineExtend50.Inputs("Curve1").Add oArcByFillet48
    oLineExtend50.Parameter("Context1") = 1
    oLineExtend50.Parameter("Context2") = 3
    oLineExtend50.Evaluate

    Dim oCurveByCurves51 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByCurves51 = oGCFactory.CreateEntity("CurveByCurves", pPOM, "051-CurveByCurves")
    oCurveByCurves51.Inputs("Curves").Add oLineExtend49
    oCurveByCurves51.Inputs("Curves").Add oArcByFillet48
    oCurveByCurves51.Inputs("Curves").Add oLineExtend50
    oCurveByCurves51.Parameter("ConstructionSurface") = 0
    oCurveByCurves51.Evaluate

    Dim oSurfByLinearExtrusion52 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion52 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "052-SurfByLinearExtrusion")
    oSurfByLinearExtrusion52.Inputs("PlanarCrossSection").Add oCurveByCurves51
    oSurfByLinearExtrusion52.Inputs("ExtrusionLine").Add oLineByPoints35
    oSurfByLinearExtrusion52.Evaluate
  
    Dim oExtractPorts53 As SP3DGeometricConstruction.GeometricConstruction
    Set oExtractPorts53 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "053-ExtractPorts")
    oExtractPorts53.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "1"
    oExtractPorts53.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1)
    oExtractPorts53.Parameter("Support") = 2 'pGeometricConstruction.Parameter("Direction")
    oExtractPorts53.Parameter("Debug") = 0
    oExtractPorts53.Evaluate
    pGeometricConstruction.ControlledInputs("Port1_1_3").Clear
    pGeometricConstruction.ControlledInputs("Port1_1_3").AddElements oExtractPorts53.ControlledInputs("AxisPort")
    pGeometricConstruction.ControlledInputs("FacePort2").Clear
    pGeometricConstruction.ControlledInputs("FacePort2").AddElements oExtractPorts53.ControlledInputs("FacePort")
    pGeometricConstruction.ControlledInputs("RootPlateSystem2").Clear
    pGeometricConstruction.ControlledInputs("RootPlateSystem2").AddElements oExtractPorts53.ControlledInputs("RootPlateSystem")
    pGeometricConstruction.ControlledInputs("LeafPlateSystem2").Clear
    pGeometricConstruction.ControlledInputs("LeafPlateSystem2").AddElements oExtractPorts53.ControlledInputs("LeafPlateSystem")
    
    Dim oCSByPlane54 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPlane54 = oGCFactory.CreateEntity("CSByPlane", pPOM, "054-CSByPlane")
    oCSByPlane54.Inputs("Plane").Add oExtractPorts53.Output("Support", 1)
    oCSByPlane54.Inputs("Point").Add oOrderAxes21.Output("Node", 1)
    oCSByPlane54.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
    oGCMacro.Output("CoordinateSystem", 1) = oCSByPlane54.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion40.Output
    oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion52.Output
    oGCMacro.Output("Boundary", 3) = oSurfFromGType41.Output
    oGCMacro.Output("Boundary", 4) = oSurfFromGType21.Output
    oGCMacro.Output("Boundary", 5) = oSurfFromGType43.Output
    oGCMacro.Output("Boundary", 6) = oSurfFromGType24.Output
    oGCMacro.Output("Support", 1) = oExtractPorts53.Output("Support", 1)
End Sub

