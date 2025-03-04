VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TInsertPlate1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oGenericPlateHelper As Object

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

             pGC.Parameter("MPlateThickness") = 0.02

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
    Call pGC.RaiseError(Err.Number)
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
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile1", "Select a Continuous Beam Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile1", "BehaviorOfPrimaryProfile1", 1, 0, 0, 0, 0, "MembersSplitByPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("PrimaryProfile2", "Select a Bounded Beam Member", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] AND [GCFilters.Filters,IsLinearStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfPrimaryProfile2", "BehaviorOfPrimaryProfile2", 1, 0, 0, 0, 0, "MembersBoundByEdgesOfPlate", 0)
    Call pGeometricConstructionDefinition.AddInput("SecondaryProfile1", "Select an Optional Auxilary Member(Column or Brace)", "ISPSMemberPartPrismatic or ISPSDesignedMember", 0, 8, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfSecondaryProfile1", "BehaviorOfSecondaryProfile1", 1, 0, 0, 0, 0, "MembersBoundByPlate", 0)
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_1", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("Port1_1_2", "ISPSSplitAxisAlongPort")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_8_1")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort_13_1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1_83_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort3")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_85_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_86_1")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem_62_1", "IJContactNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem_62_1", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_69_1")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(1, "Node")
    Call pGeometricConstructionDefinition.AddParameter("offset1", "offset1", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("offset2", "offset2", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("offset", "offset", 8, 1, 59, 0, 0, 0.0001)
    Call pGeometricConstructionDefinition.AddParameter("Location", "Location", 4, 0, 0, 0, 0, 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Negative", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Location", "Positive", 2)
    Call pGeometricConstructionDefinition.AddParameter("MPlateThickness", "MPlateThickness", 8, 1, 59, 0, 0, 0.05, False)
    
    Call GeometricConstructionDefinition_CopyErrors(pGeometricConstructionDefinition, GeometricConstructionDefinition_GetFromTypeName("OrderAxes2"), 100)

    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oMemberPart As Object: Set oMemberPart = pGeometricConstruction.Inputs("PrimaryProfile1").Item(1)
    If Not TypeOf oMemberPart Is IJStiffenerSystem Then
    
        On Error Resume Next
        Dim oOrderAxes21 As SP3DGeometricConstruction.GeometricConstruction
        Set oOrderAxes21 = oGCFactory.CreateEntity("OrderAxes2", pPOM, "0001-OrderAxes2")
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "2"
        oOrderAxes21.Parameter("Support") = pGeometricConstruction.Parameter("Location")
        oOrderAxes21.Evaluate
        Dim lError As Long: Let lError = Err.Number
        On Error GoTo 0
        If lError <> 0 Then
            Err.Raise m_oGenericPlateHelper.GetErrorNumber(oOrderAxes21) + 100
        End If

        pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_1").Add oOrderAxes21.ControlledInputs("Port1").Item("1")
        pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_2").Add oOrderAxes21.ControlledInputs("Port1").Item("2")
        Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0002-LineFromCS")
        oLineFromCS2.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1), "1"
        oLineFromCS2.Parameter("LookingAxis") = 3
        oLineFromCS2.Parameter("Length") = -2#
        oLineFromCS2.Parameter("CSOrientation") = 1
        oLineFromCS2.Parameter("LineJustification") = 1
        oLineFromCS2.Evaluate
    
        Dim oExtractPorts3 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts3 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0003-ExtractPorts")
        oExtractPorts3.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oExtractPorts3.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
        oExtractPorts3.Parameter("Support") = 2
        oExtractPorts3.Parameter("Debug") = 0
        oExtractPorts3.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_1").Add oExtractPorts3.ControlledInputs("AxisPort").Item("1")
        pGeometricConstruction.ControlledInputs("FacePort_8_1").Clear
        pGeometricConstruction.ControlledInputs("FacePort_8_1").Add oExtractPorts3.ControlledInputs("FacePort").Item("1")
        Dim oExtractPorts4 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts4 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0004-ExtractPorts")
        oExtractPorts4.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "2"
        oExtractPorts4.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
        oExtractPorts4.Parameter("Support") = 2
        oExtractPorts4.Parameter("Debug") = 0
        oExtractPorts4.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_2").Add oExtractPorts4.ControlledInputs("AxisPort").Item("1")
        pGeometricConstruction.ControlledInputs("FacePort_13_1").Clear
        pGeometricConstruction.ControlledInputs("FacePort_13_1").Add oExtractPorts4.ControlledInputs("FacePort").Item("1")
        Dim oParamDistMeasureParallel5 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel5 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0005-ParamDistMeasureParallel")
        oParamDistMeasureParallel5.Inputs("Graphics").Add oExtractPorts3.Output("CurveRight", 1), "1"
        oParamDistMeasureParallel5.Inputs("Graphics").Add oExtractPorts3.Output("MemberAxis", 1), "2"
        oParamDistMeasureParallel5.Parameter("Value") = 0.2032
        oParamDistMeasureParallel5.Evaluate
    
        Dim oPointAtCurveStart6 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart6 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0006-PointAtCurveStart")
        oPointAtCurveStart6.Inputs("Curve").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPointAtCurveStart6.Evaluate
    
        Dim oParamDistMeasureParallel7 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel7 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0007-ParamDistMeasureParallel")
        oParamDistMeasureParallel7.Inputs("Graphics").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oParamDistMeasureParallel7.Inputs("Graphics").Add oExtractPorts4.Output("CurveRight", 1), "2"
        oParamDistMeasureParallel7.Parameter("Value") = 0.2032
        oParamDistMeasureParallel7.Evaluate
        
        Dim pElementsOfSecondaryMembers As IJElements
        Set pElementsOfSecondaryMembers = pGeometricConstruction.Inputs("SecondaryProfile1")
        Dim iCountOfSecondaryMembers As Integer
        Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
        If iCountOfSecondaryMembers > 0 Then
    
            Dim oExtractCircle8 As SP3DGeometricConstruction.GeometricConstruction
            Set oExtractCircle8 = oGCFactory.CreateEntity("ExtractCircle", pPOM, "0008-ExtractCircle")
            oExtractCircle8.Inputs("MemberPart1").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
            oExtractCircle8.Inputs("MemberPart2").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "1"
            oExtractCircle8.Parameter("Debug") = 0
            oExtractCircle8.Evaluate
        
            pGeometricConstruction.ControlledInputs("AxisPort1_83_1").Clear
            pGeometricConstruction.ControlledInputs("AxisPort1_83_1").Add oExtractCircle8.ControlledInputs("AxisPort1").Item("1")
            pGeometricConstruction.ControlledInputs("FacePort3").Clear
            pGeometricConstruction.ControlledInputs("FacePort3").Add oExtractCircle8.ControlledInputs("FacePort1").Item("1")
            pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
            pGeometricConstruction.ControlledInputs("Port1_1_2").Add oExtractCircle8.ControlledInputs("AxisPort2").Item("1")
        
        End If
        
        Dim oPointAtCurveMiddle9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle9 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0009-PointAtCurveMiddle")
        oPointAtCurveMiddle9.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointAtCurveMiddle9.Evaluate
    
        Dim oParamDistConstant10 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant10 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0010-ParamDistConstant")
        oParamDistConstant10.Parameter("Value") = pGeometricConstruction.Parameter("offset1")
        oParamDistConstant10.Evaluate
    
        Dim oParamDistConstant11 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant11 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0011-ParamDistConstant")
        oParamDistConstant11.Parameter("Value") = pGeometricConstruction.Parameter("offset2")
        oParamDistConstant11.Evaluate
    
        Dim oParamDistConstant12 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant12 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0012-ParamDistConstant")
        oParamDistConstant12.Parameter("Value") = pGeometricConstruction.Parameter("offset")
        oParamDistConstant12.Evaluate
    
        Dim oParamDistConstant13 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant13 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0013-ParamDistConstant")
        oParamDistConstant13.Parameter("Value") = 1
        oParamDistConstant13.Evaluate
    
        Dim oPointAtCurveExtremity14 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity14 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0014-PointAtCurveExtremity")
        oPointAtCurveExtremity14.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointAtCurveExtremity14.Inputs("TrackPoint").Add oOrderAxes21.Output("Node", 1), "1"
        oPointAtCurveExtremity14.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity14.Evaluate
    
        Dim oVectorTangentToCurve15 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve15 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "0015-VectorTangentToCurve")
        oVectorTangentToCurve15.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oVectorTangentToCurve15.Inputs("Point").Add oPointAtCurveExtremity14, "1"
        oVectorTangentToCurve15.Inputs("TrackPoint").Add oPointAtCurveMiddle9, "1"
        oVectorTangentToCurve15.Parameter("Range") = 2#
        oVectorTangentToCurve15.Parameter("Orientation") = 1
        oVectorTangentToCurve15.Parameter("TrackFlag") = 2
        oVectorTangentToCurve15.Evaluate
    
        Dim oPointAtCurveExtremity16 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity16 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0016-PointAtCurveExtremity")
        oPointAtCurveExtremity16.Inputs("Curve").Add oVectorTangentToCurve15, "1"
        oPointAtCurveExtremity16.Inputs("TrackPoint").Add oPointAtCurveExtremity14, "1"
        oPointAtCurveExtremity16.Parameter("TrackFlag") = 2
        oPointAtCurveExtremity16.Evaluate
    
        Dim oParamDistFunctionWith2Args17 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args17 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0017-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args17.Inputs("Arg1").Add oParamDistMeasureParallel5, "1"
        oParamDistFunctionWith2Args17.Inputs("Arg2").Add oParamDistConstant10, "1"
        oParamDistFunctionWith2Args17.Parameter("Value") = 0.3032
        oParamDistFunctionWith2Args17.Parameter("Operation") = 1
        oParamDistFunctionWith2Args17.Evaluate
    
        Dim oParamDistFunctionWith2Args18 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args18 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0018-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args18.Inputs("Arg1").Add oParamDistMeasureParallel5, "1"
        oParamDistFunctionWith2Args18.Inputs("Arg2").Add oParamDistConstant12, "1"
        oParamDistFunctionWith2Args18.Parameter("Value") = 0.2033
        oParamDistFunctionWith2Args18.Parameter("Operation") = 1
        oParamDistFunctionWith2Args18.Evaluate
    
        Dim oParamDistFunctionWith2Args19 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args19 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0019-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args19.Inputs("Arg1").Add oParamDistFunctionWith2Args18, "1"
        oParamDistFunctionWith2Args19.Inputs("Arg2").Add oParamDistConstant13, "1"
        oParamDistFunctionWith2Args19.Parameter("Value") = 0.2533
        oParamDistFunctionWith2Args19.Parameter("Operation") = 1
        oParamDistFunctionWith2Args19.Evaluate
    
        Dim oParamDistFunctionWith2Args20 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args20 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0020-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args20.Inputs("Arg1").Add oParamDistMeasureParallel7, "1"
        oParamDistFunctionWith2Args20.Inputs("Arg2").Add oParamDistConstant11, "1"
        oParamDistFunctionWith2Args20.Parameter("Value") = 0.3032
        oParamDistFunctionWith2Args20.Parameter("Operation") = 1
        oParamDistFunctionWith2Args20.Evaluate
    
        Dim oParamDistFunctionWith2Args21 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args21 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0021-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args21.Inputs("Arg1").Add oParamDistMeasureParallel7, "1"
        oParamDistFunctionWith2Args21.Inputs("Arg2").Add oParamDistConstant12, "1"
        oParamDistFunctionWith2Args21.Parameter("Value") = 0.2033
        oParamDistFunctionWith2Args21.Parameter("Operation") = 1
        oParamDistFunctionWith2Args21.Evaluate
    
        Dim oParamDistFunctionWith2Args22 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args22 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0022-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args22.Inputs("Arg1").Add oParamDistFunctionWith2Args21, "1"
        oParamDistFunctionWith2Args22.Inputs("Arg2").Add oParamDistConstant13, "1"
        oParamDistFunctionWith2Args22.Parameter("Value") = 0.2533
        oParamDistFunctionWith2Args22.Parameter("Operation") = 1
        oParamDistFunctionWith2Args22.Evaluate
    
        Dim oLineParallel23 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel23 = oGCFactory.CreateEntity("LineParallel", pPOM, "0023-LineParallel")
        oLineParallel23.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel23.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel23.Inputs("TrackPoint").Add oPointAtCurveMiddle9, "1"
        oLineParallel23.Parameter("Distance") = oParamDistFunctionWith2Args17
        oLineParallel23.Parameter("TrackFlag") = 2
        oLineParallel23.Evaluate
    
        Dim oLineParallel24 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel24 = oGCFactory.CreateEntity("LineParallel", pPOM, "0024-LineParallel")
        oLineParallel24.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel24.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel24.Inputs("TrackPoint").Add oPointAtCurveMiddle9, "1"
        oLineParallel24.Parameter("Distance") = oParamDistFunctionWith2Args17
        oLineParallel24.Parameter("TrackFlag") = 1
        oLineParallel24.Evaluate
    
        Dim oLineParallel25 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel25 = oGCFactory.CreateEntity("LineParallel", pPOM, "0025-LineParallel")
        oLineParallel25.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel25.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel25.Inputs("TrackPoint").Add oPointAtCurveMiddle9, "1"
        oLineParallel25.Parameter("Distance") = oParamDistFunctionWith2Args18
        oLineParallel25.Parameter("TrackFlag") = 2
        oLineParallel25.Evaluate
    
        Dim oLineParallel26 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel26 = oGCFactory.CreateEntity("LineParallel", pPOM, "0026-LineParallel")
        oLineParallel26.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel26.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel26.Inputs("TrackPoint").Add oPointAtCurveMiddle9, "1"
        oLineParallel26.Parameter("Distance") = oParamDistFunctionWith2Args18
        oLineParallel26.Parameter("TrackFlag") = 1
        oLineParallel26.Evaluate
    
        Dim oLineByPoints27 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints27 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0027-LineByPoints")
        oLineByPoints27.Inputs("StartPoint").Add oPointAtCurveExtremity16, "1"
        oLineByPoints27.Inputs("EndPoint").Add oPointAtCurveMiddle9, "1"
        oLineByPoints27.Evaluate
    
        Dim oPointByCurves28 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves28 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0028-PointByCurves")
        oPointByCurves28.Inputs("Curve1").Add oExtractPorts4.Output("MemberAxis", 1), "2"
        oPointByCurves28.Inputs("Curve2").Add oLineParallel24, "1"
        oPointByCurves28.Parameter("TrackFlag") = 1
        oPointByCurves28.Evaluate
    
        Dim oPlaneByPointNormal29 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal29 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0029-PlaneByPointNormal")
        oPlaneByPointNormal29.Inputs("Point").Add oPointByCurves28, "2"
        oPlaneByPointNormal29.Inputs("Line").Add oExtractPorts4.Output("MemberAxis", 1), "2"
        oPlaneByPointNormal29.Parameter("Range") = oParamDistFunctionWith2Args22
        oPlaneByPointNormal29.Evaluate
    
        Dim oSurfFromGType30 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType30 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0030-SurfFromGType")
        oSurfFromGType30.Inputs("Surface").Add oPlaneByPointNormal29, "2"
        oSurfFromGType30.Evaluate
    
        Dim oLineParallel31 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel31 = oGCFactory.CreateEntity("LineParallel", pPOM, "0031-LineParallel")
        oLineParallel31.Inputs("Line").Add oLineByPoints27, "1"
        oLineParallel31.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel31.Inputs("TrackPoint").Add oPointAtCurveStart6, "1"
        oLineParallel31.Parameter("Distance") = oParamDistFunctionWith2Args20
        oLineParallel31.Parameter("TrackFlag") = 1
        oLineParallel31.Evaluate
    
        Dim oLineParallel32 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel32 = oGCFactory.CreateEntity("LineParallel", pPOM, "0032-LineParallel")
        oLineParallel32.Inputs("Line").Add oLineByPoints27, "1"
        oLineParallel32.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel32.Inputs("TrackPoint").Add oPointAtCurveStart6, "1"
        oLineParallel32.Parameter("Distance") = oParamDistFunctionWith2Args20
        oLineParallel32.Parameter("TrackFlag") = 2
        oLineParallel32.Evaluate
    
        Dim oLineParallel33 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel33 = oGCFactory.CreateEntity("LineParallel", pPOM, "0033-LineParallel")
        oLineParallel33.Inputs("Line").Add oLineByPoints27, "1"
        oLineParallel33.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel33.Inputs("TrackPoint").Add oPointAtCurveStart6, "1"
        oLineParallel33.Parameter("Distance") = oParamDistFunctionWith2Args21
        oLineParallel33.Parameter("TrackFlag") = 1
        oLineParallel33.Evaluate
    
        Dim oLineParallel34 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel34 = oGCFactory.CreateEntity("LineParallel", pPOM, "0034-LineParallel")
        oLineParallel34.Inputs("Line").Add oLineByPoints27, "1"
        oLineParallel34.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel34.Inputs("TrackPoint").Add oPointAtCurveStart6, "1"
        oLineParallel34.Parameter("Distance") = oParamDistFunctionWith2Args21
        oLineParallel34.Parameter("TrackFlag") = 2
        oLineParallel34.Evaluate
    
        Dim oPointByCurves35 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves35 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0035-PointByCurves")
        oPointByCurves35.Inputs("Curve1").Add oLineParallel34, "1"
        oPointByCurves35.Inputs("Curve2").Add oLineParallel23, "1"
        oPointByCurves35.Parameter("TrackFlag") = 1
        oPointByCurves35.Evaluate
    
        Dim oPointByCurves36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves36 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0036-PointByCurves")
        oPointByCurves36.Inputs("Curve1").Add oLineParallel33, "1"
        oPointByCurves36.Inputs("Curve2").Add oLineParallel23, "1"
        oPointByCurves36.Parameter("TrackFlag") = 1
        oPointByCurves36.Evaluate
    
        Dim oPointByCurves37 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves37 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0037-PointByCurves")
        oPointByCurves37.Inputs("Curve1").Add oLineParallel34, "2"
        oPointByCurves37.Inputs("Curve2").Add oLineParallel24, "1"
        oPointByCurves37.Parameter("TrackFlag") = 1
        oPointByCurves37.Evaluate
    
        Dim oPointByCurves38 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves38 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0038-PointByCurves")
        oPointByCurves38.Inputs("Curve1").Add oLineParallel33, "2"
        oPointByCurves38.Inputs("Curve2").Add oLineParallel24, "1"
        oPointByCurves38.Parameter("TrackFlag") = 1
        oPointByCurves38.Evaluate
    
        Dim oPointByCurves39 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves39 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0039-PointByCurves")
        oPointByCurves39.Inputs("Curve1").Add oLineParallel26, "1"
        oPointByCurves39.Inputs("Curve2").Add oLineParallel31, "1"
        oPointByCurves39.Parameter("TrackFlag") = 1
        oPointByCurves39.Evaluate
    
        Dim oPointByCurves40 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves40 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0040-PointByCurves")
        oPointByCurves40.Inputs("Curve1").Add oLineParallel25, "1"
        oPointByCurves40.Inputs("Curve2").Add oLineParallel31, "1"
        oPointByCurves40.Parameter("TrackFlag") = 1
        oPointByCurves40.Evaluate
    
        Dim oPointByCurves41 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves41 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0041-PointByCurves")
        oPointByCurves41.Inputs("Curve1").Add oLineParallel26, "1"
        oPointByCurves41.Inputs("Curve2").Add oLineParallel32, "1"
        oPointByCurves41.Parameter("TrackFlag") = 1
        oPointByCurves41.Evaluate
    
        Dim oPointByCurves42 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves42 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0042-PointByCurves")
        oPointByCurves42.Inputs("Curve1").Add oLineParallel25, "2"
        oPointByCurves42.Inputs("Curve2").Add oLineParallel32, "2"
        oPointByCurves42.Parameter("TrackFlag") = 1
        oPointByCurves42.Evaluate
    
        Dim oLineByPoints43 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints43 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0043-LineByPoints")
        oLineByPoints43.Inputs("StartPoint").Add oPointByCurves37, "2"
        oLineByPoints43.Inputs("EndPoint").Add oPointByCurves41, "1"
        oLineByPoints43.Evaluate
    
        Dim oPointAlongCurve44 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve44 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0044-PointAlongCurve")
        oPointAlongCurve44.Inputs("Curve").Add oLineByPoints43, "2"
        oPointAlongCurve44.Inputs("Point").Add oPointByCurves41, "2"
        oPointAlongCurve44.Inputs("TrackPoint").Add oPointByCurves37, "2"
        oPointAlongCurve44.Parameter("Distance") = 0.05
        oPointAlongCurve44.Parameter("TrackFlag") = 2
        oPointAlongCurve44.Evaluate
    
        Dim oPointAlongCurve45 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve45 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0045-PointAlongCurve")
        oPointAlongCurve45.Inputs("Curve").Add oLineByPoints43, "1"
        oPointAlongCurve45.Inputs("Point").Add oPointByCurves37, "2"
        oPointAlongCurve45.Inputs("TrackPoint").Add oPointByCurves41, "2"
        oPointAlongCurve45.Parameter("Distance") = 0.05
        oPointAlongCurve45.Parameter("TrackFlag") = 2
        oPointAlongCurve45.Evaluate
    
        Dim oLineByPoints46 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints46 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0046-LineByPoints")
        oLineByPoints46.Inputs("StartPoint").Add oPointByCurves39, "2"
        oLineByPoints46.Inputs("EndPoint").Add oPointByCurves38, "1"
        oLineByPoints46.Evaluate
    
        Dim oPointAlongCurve47 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve47 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0047-PointAlongCurve")
        oPointAlongCurve47.Inputs("Curve").Add oLineByPoints46, "2"
        oPointAlongCurve47.Inputs("Point").Add oPointByCurves38, "2"
        oPointAlongCurve47.Inputs("TrackPoint").Add oPointByCurves39, "2"
        oPointAlongCurve47.Parameter("Distance") = 0.05
        oPointAlongCurve47.Parameter("TrackFlag") = 2
        oPointAlongCurve47.Evaluate
    
        Dim oPointAlongCurve48 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve48 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0048-PointAlongCurve")
        oPointAlongCurve48.Inputs("Curve").Add oLineByPoints46, "1"
        oPointAlongCurve48.Inputs("Point").Add oPointByCurves39, "2"
        oPointAlongCurve48.Inputs("TrackPoint").Add oPointByCurves38, "2"
        oPointAlongCurve48.Parameter("Distance") = 0.05
        oPointAlongCurve48.Parameter("TrackFlag") = 2
        oPointAlongCurve48.Evaluate
    
        Dim oLineByPoints49 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints49 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0049-LineByPoints")
        oLineByPoints49.Inputs("StartPoint").Add oPointByCurves42, "1"
        oLineByPoints49.Inputs("EndPoint").Add oPointByCurves35, "1"
        oLineByPoints49.Evaluate
    
        Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0050-PointAlongCurve")
        oPointAlongCurve50.Inputs("Curve").Add oLineByPoints49, "1"
        oPointAlongCurve50.Inputs("Point").Add oPointByCurves42, "1"
        oPointAlongCurve50.Inputs("TrackPoint").Add oPointByCurves35, "1"
        oPointAlongCurve50.Parameter("Distance") = 0.05
        oPointAlongCurve50.Parameter("TrackFlag") = 2
        oPointAlongCurve50.Evaluate
    
        Dim oLineByPoints51 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints51 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0051-LineByPoints")
        oLineByPoints51.Inputs("StartPoint").Add oPointByCurves36, "1"
        oLineByPoints51.Inputs("EndPoint").Add oPointByCurves40, "1"
        oLineByPoints51.Evaluate
    
        Dim oPointAlongCurve52 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve52 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0052-PointAlongCurve")
        oPointAlongCurve52.Inputs("Curve").Add oLineByPoints51, "2"
        oPointAlongCurve52.Inputs("Point").Add oPointByCurves40, "2"
        oPointAlongCurve52.Inputs("TrackPoint").Add oPointByCurves36, "2"
        oPointAlongCurve52.Parameter("Distance") = 0.05
        oPointAlongCurve52.Parameter("TrackFlag") = 2
        oPointAlongCurve52.Evaluate
    
        Dim oLineByPoints53 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints53 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0053-LineByPoints")
        oLineByPoints53.Inputs("StartPoint").Add oPointByCurves35, "1"
        oLineByPoints53.Inputs("EndPoint").Add oPointByCurves36, "1"
        oLineByPoints53.Evaluate
    
        Dim oLineByPoints54 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints54 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0054-LineByPoints")
        oLineByPoints54.Inputs("StartPoint").Add oPointAlongCurve50, "1"
        oLineByPoints54.Inputs("EndPoint").Add oPointByCurves35, "1"
        oLineByPoints54.Evaluate
    
        Dim oLineByPoints55 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints55 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0055-LineByPoints")
        oLineByPoints55.Inputs("StartPoint").Add oPointByCurves36, "2"
        oLineByPoints55.Inputs("EndPoint").Add oPointAlongCurve52, "1"
        oLineByPoints55.Evaluate
    
        Dim oLineByPoints56 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints56 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0056-LineByPoints")
        oLineByPoints56.Inputs("StartPoint").Add oPointAlongCurve48, "2"
        oLineByPoints56.Inputs("EndPoint").Add oPointAlongCurve47, "1"
        oLineByPoints56.Evaluate
    
        Dim oLineByPoints57 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints57 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0057-LineByPoints")
        oLineByPoints57.Inputs("StartPoint").Add oPointAlongCurve45, "2"
        oLineByPoints57.Inputs("EndPoint").Add oPointAlongCurve44, "1"
        oLineByPoints57.Evaluate
    
        Dim oSurfByLinearExtrusion58 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion58 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0058-SurfByLinearExtrusion")
        oSurfByLinearExtrusion58.Inputs("PlanarCrossSection").Add oLineByPoints56, "1"
        oSurfByLinearExtrusion58.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion58.Evaluate
    
        Dim oSurfByLinearExtrusion59 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion59 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0059-SurfByLinearExtrusion")
        oSurfByLinearExtrusion59.Inputs("PlanarCrossSection").Add oLineByPoints57, "1"
        oSurfByLinearExtrusion59.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion59.Evaluate
    
        Dim oCpxStringByCurves60 As SP3DGeometricConstruction.GeometricConstruction
        Set oCpxStringByCurves60 = oGCFactory.CreateEntity("CpxStringByCurves", pPOM, "0060-CpxStringByCurves")
        oCpxStringByCurves60.Inputs("Curves").Add oLineByPoints54, "1"
        oCpxStringByCurves60.Inputs("Curves").Add oLineByPoints53, "2"
        oCpxStringByCurves60.Inputs("Curves").Add oLineByPoints55, "3"
        oCpxStringByCurves60.Evaluate
    
        Dim oPointByCurves61 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves61 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0061-PointByCurves")
        oPointByCurves61.Inputs("Curve1").Add oExtractPorts3.Output("MemberAxis", 1), "2"
        oPointByCurves61.Inputs("Curve2").Add oLineParallel32, "1"
        oPointByCurves61.Parameter("TrackFlag") = 1
        oPointByCurves61.Evaluate
    
        Dim oPointByCurves62 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves62 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0062-PointByCurves")
        oPointByCurves62.Inputs("Curve1").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPointByCurves62.Inputs("Curve2").Add oLineParallel31, "1"
        oPointByCurves62.Parameter("TrackFlag") = 1
        oPointByCurves62.Evaluate
    
        Dim oPlaneByPointNormal63 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal63 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0063-PlaneByPointNormal")
        oPlaneByPointNormal63.Inputs("Point").Add oPointByCurves61, "1"
        oPlaneByPointNormal63.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal63.Parameter("Range") = oParamDistFunctionWith2Args19
        oPlaneByPointNormal63.Evaluate
    
        Dim oPlaneByPointNormal64 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal64 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0064-PlaneByPointNormal")
        oPlaneByPointNormal64.Inputs("Point").Add oPointByCurves62, "2"
        oPlaneByPointNormal64.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal64.Parameter("Range") = oParamDistFunctionWith2Args19
        oPlaneByPointNormal64.Evaluate
    
        Dim oSurfByLinearExtrusion65 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion65 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0065-SurfByLinearExtrusion")
        oSurfByLinearExtrusion65.Inputs("PlanarCrossSection").Add oCpxStringByCurves60, "1"
        oSurfByLinearExtrusion65.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion65.Evaluate
    
        Dim oSurfFromGType66 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType66 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0066-SurfFromGType")
        oSurfFromGType66.Inputs("Surface").Add oPlaneByPointNormal63, "1"
        oSurfFromGType66.Evaluate
    
        Dim oSurfFromGType67 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType67 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0067-SurfFromGType")
        oSurfFromGType67.Inputs("Surface").Add oPlaneByPointNormal64, "2"
        oSurfFromGType67.Evaluate
        
        Dim oCSFromMember68 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSFromMember68 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0068-CSFromMember")
        oCSFromMember68.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oCSFromMember68.Evaluate
        
        Dim oFacePortExtractor69 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor69 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0069-FacePortExtractor")
        oFacePortExtractor69.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oFacePortExtractor69.Inputs("CoordinateSystem").Add oCSFromMember68, "1"
        oFacePortExtractor69.Parameter("GeometrySelector") = 2
        oFacePortExtractor69.Parameter("FacesContext") = 4
        oFacePortExtractor69.Parameter("LookingAxis") = 3
        oFacePortExtractor69.Parameter("IntersectingPlane") = 0
        oFacePortExtractor69.Parameter("SurfaceType") = 1
        oFacePortExtractor69.Parameter("TrackFlag") = 1
        oFacePortExtractor69.Parameter("Offset") = 0#
        oFacePortExtractor69.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port_85_1").Clear
        pGeometricConstruction.ControlledInputs("Port_85_1").Add oFacePortExtractor69.ControlledInputs("Port").Item("1")
        Dim oFacePortExtractor70 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor70 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0070-FacePortExtractor")
        oFacePortExtractor70.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oFacePortExtractor70.Inputs("CoordinateSystem").Add oCSFromMember68, "1"
        oFacePortExtractor70.Parameter("GeometrySelector") = 2
        oFacePortExtractor70.Parameter("FacesContext") = 2
        oFacePortExtractor70.Parameter("LookingAxis") = 3
        oFacePortExtractor70.Parameter("IntersectingPlane") = 0
        oFacePortExtractor70.Parameter("SurfaceType") = 1
        oFacePortExtractor70.Parameter("TrackFlag") = 1
        oFacePortExtractor70.Parameter("Offset") = 0#
        oFacePortExtractor70.Evaluate
        pGeometricConstruction.ControlledInputs("Port_86_1").Clear
        pGeometricConstruction.ControlledInputs("Port_86_1").Add oFacePortExtractor70.ControlledInputs("Port").Item("1")
        
        Dim oMember As Object
        Set oMember = pGeometricConstruction.ControlledInputs("PrimaryProfile1")(1)
        Dim oSD_Member As New StructDetailObjects.MemberPart
        Set oSD_Member.object = oMember
        If oSD_Member.SectionType = "S" Then
        
            Dim oFacePortExtractor95 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortExtractor95 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0095-FacePortExtractor")
            oFacePortExtractor95.Inputs("Connectable").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
            oFacePortExtractor95.Inputs("CoordinateSystem").Add oCSFromMember68, "1"
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
            oCurveByIntersection96.Inputs("Surface1").Add oFacePortExtractor69, "1"
            oCurveByIntersection96.Inputs("Surface2").Add oFacePortExtractor95, "1"
            oCurveByIntersection96.Parameter("TrackFlag") = 1
            oCurveByIntersection96.Parameter("ConstructionSurface") = 0
            oCurveByIntersection96.Evaluate
        
            Dim oCurveByProjection97 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByProjection97 = oGCFactory.CreateEntity("CurveByProjection", pPOM, "0097-CurveByProjection")
            oCurveByProjection97.Inputs("Curve").Add oCurveByIntersection96, "1"
            oCurveByProjection97.Inputs("Surface").Add oFacePortExtractor70, "2"
            oCurveByProjection97.Inputs("Line").Add oLineFromCS2, "1"
            oCurveByProjection97.Parameter("TrackFlag") = 1
            oCurveByProjection97.Parameter("ConstructionSurface") = 0
            oCurveByProjection97.Evaluate
        
            Dim oParamDistMeasureParallel71 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel71 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0071-ParamDistMeasureParallel")
            oParamDistMeasureParallel71.Inputs("Graphics").Add oCurveByIntersection96, "1"
            oParamDistMeasureParallel71.Inputs("Graphics").Add oCurveByProjection97, "2"
            oParamDistMeasureParallel71.Parameter("Value") = 3.53569989408697E-02
            oParamDistMeasureParallel71.Evaluate
    
        Else
        
    '        Dim oParamDistMeasureParallel71 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel71 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0071-ParamDistMeasureParallel")
            oParamDistMeasureParallel71.Inputs("Graphics").Add oFacePortExtractor69, "1"
            oParamDistMeasureParallel71.Inputs("Graphics").Add oFacePortExtractor70, "2"
            oParamDistMeasureParallel71.Parameter("Value") = 4.49580000000012E-02
            oParamDistMeasureParallel71.Evaluate
            
        End If
        
        dT = oParamDistMeasureParallel71.Parameter("Value")
    
        Dim oGCMacro As IJGeometricConstructionMacro
        Set oGCMacro = pGeometricConstruction
    
        oGCMacro.Output("Support", 1) = oExtractPorts3.Output("Support", 1)
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion65.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion59.Output
        oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion58.Output
        oGCMacro.Output("Boundary", 4) = oSurfFromGType66.Output
        oGCMacro.Output("Boundary", 5) = oSurfFromGType67.Output
        oGCMacro.Output("Boundary", 6) = oSurfFromGType30.Output
    '    If iCountOfSecondaryMembers > 0 Then
    '        oGCMacro.Output("Boundary", "FacePort3") = oExtractCircle8.Output("Boundary", 1)'(Function:Create LC Between Tube and APS)
    '    End If
        oGCMacro.Output("CoordinateSystem", 1) = oOrderAxes21.Output("CoordinateSystem", 1)
        oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
    
    Else
    
'        Dim oOrderAxes21 As SP3DGeometricConstruction.GeometricConstruction
        Set oOrderAxes21 = oGCFactory.CreateEntity("OrderAxes2", pPOM, "0001-OrderAxes2")
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oOrderAxes21.Inputs("PrimaryProfiles").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "2"
        oOrderAxes21.Parameter("Support") = 2
        oOrderAxes21.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_1").Add oOrderAxes21.ControlledInputs("Port1").Item("1")
        pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_2").Add oOrderAxes21.ControlledInputs("Port1").Item("2")
'        Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "0002-LineFromCS")
        oLineFromCS2.Inputs("CoordinateSystem").Add oOrderAxes21.Output("CoordinateSystem", 1), "1"
        oLineFromCS2.Parameter("LookingAxis") = 3
        oLineFromCS2.Parameter("Length") = -2#
        oLineFromCS2.Parameter("CSOrientation") = 1
        oLineFromCS2.Parameter("LineJustification") = 1
        oLineFromCS2.Evaluate
    
'        Dim oExtractPorts3 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts3 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0003-ExtractPorts")
        oExtractPorts3.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oExtractPorts3.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
        oExtractPorts3.Parameter("Support") = 2
        oExtractPorts3.Parameter("Debug") = 0
        oExtractPorts3.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port1_1_1").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_1").Add oExtractPorts3.ControlledInputs("AxisPort").Item("1")
'        Dim oExtractPorts4 As SP3DGeometricConstruction.GeometricConstruction
        Set oExtractPorts4 = oGCFactory.CreateEntity("ExtractPorts", pPOM, "0004-ExtractPorts")
        oExtractPorts4.Inputs("MemberPart").Add pGeometricConstruction.Inputs("PrimaryProfile2").Item(1), "2"
        oExtractPorts4.Inputs("ZAxis").Add oOrderAxes21.Output("Normal", 1), "1"
        oExtractPorts4.Parameter("Support") = 2
        oExtractPorts4.Parameter("Debug") = 0
        oExtractPorts4.Evaluate
    
        pGeometricConstruction.ControlledInputs("Port1_1_2").Clear
        pGeometricConstruction.ControlledInputs("Port1_1_2").Add oExtractPorts4.ControlledInputs("AxisPort").Item("1")
'        Dim oParamDistMeasureParallel5 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel5 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0005-ParamDistMeasureParallel")
        oParamDistMeasureParallel5.Inputs("Graphics").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oParamDistMeasureParallel5.Inputs("Graphics").Add oExtractPorts3.Output("CurveRight", 1), "2"
        oParamDistMeasureParallel5.Parameter("Value") = 7.49999999999993E-02
        oParamDistMeasureParallel5.Evaluate
    
        Dim oParamDistMeasureParallel6 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel6 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "0006-ParamDistMeasureParallel")
        oParamDistMeasureParallel6.Inputs("Graphics").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oParamDistMeasureParallel6.Inputs("Graphics").Add oExtractPorts4.Output("CurveRight", 1), "2"
        oParamDistMeasureParallel6.Parameter("Value") = 7.50000000000028E-02
        oParamDistMeasureParallel6.Evaluate
    
        Dim oPointAtCurveMiddle7 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle7 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "0007-PointAtCurveMiddle")
        oPointAtCurveMiddle7.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointAtCurveMiddle7.Evaluate
    
        Dim oParamDistConstant8 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant8 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0008-ParamDistConstant")
        oParamDistConstant8.Parameter("Value") = pGeometricConstruction.Parameter("offset1")
        oParamDistConstant8.Evaluate
    
        Dim oParamDistConstant9 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant9 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0009-ParamDistConstant")
        oParamDistConstant9.Parameter("Value") = pGeometricConstruction.Parameter("offset2")
        oParamDistConstant9.Evaluate
    
'        Dim oParamDistConstant10 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant10 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0010-ParamDistConstant")
        oParamDistConstant10.Parameter("Value") = pGeometricConstruction.Parameter("offset")
        oParamDistConstant10.Evaluate
    
        Dim oPointAtCurveExtremity11 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveExtremity11 = oGCFactory.CreateEntity("PointAtCurveExtremity", pPOM, "0011-PointAtCurveExtremity")
        oPointAtCurveExtremity11.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointAtCurveExtremity11.Inputs("TrackPoint").Add oOrderAxes21.Output("Node", 1), "1"
        oPointAtCurveExtremity11.Parameter("TrackFlag") = 1
        oPointAtCurveExtremity11.Evaluate
    
        Dim oPointAtCurveStart12 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveStart12 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "0012-PointAtCurveStart")
        oPointAtCurveStart12.Inputs("Curve").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPointAtCurveStart12.Evaluate
    
'        Dim oParamDistConstant13 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistConstant13 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "0013-ParamDistConstant")
        oParamDistConstant13.Parameter("Value") = 1#
        oParamDistConstant13.Evaluate
    
        Dim oCSByStiffener14 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByStiffener14 = oGCFactory.CreateEntity("CSByStiffener", pPOM, "0014-CSByStiffener")
        oCSByStiffener14.Inputs("Stiffener").Add pGeometricConstruction.Inputs("PrimaryProfile1").Item(1), "1"
        oCSByStiffener14.Parameter("Ortho") = -1
        oCSByStiffener14.Evaluate
    
        pGeometricConstruction.ControlledInputs("RootPlateSystem_62_1").Clear
        pGeometricConstruction.ControlledInputs("RootPlateSystem_62_1").Add oCSByStiffener14.ControlledInputs("RootPlateSystem").Item("1")
        pGeometricConstruction.ControlledInputs("LeafPlateSystem_62_1").Clear
        pGeometricConstruction.ControlledInputs("LeafPlateSystem_62_1").Add oCSByStiffener14.ControlledInputs("LeafPlateSystem").Item(1)
    
        Dim oParamDistFunctionWith2Args16 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args16 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0016-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args16.Inputs("Arg1").Add oParamDistMeasureParallel5, "1"
        oParamDistFunctionWith2Args16.Inputs("Arg2").Add oParamDistConstant8, "1"
        oParamDistFunctionWith2Args16.Parameter("Value") = 0.174999999999999
        oParamDistFunctionWith2Args16.Parameter("Operation") = 1
        oParamDistFunctionWith2Args16.Evaluate
    
'        Dim oParamDistFunctionWith2Args17 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args17 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0017-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args17.Inputs("Arg1").Add oParamDistMeasureParallel5, "1"
        oParamDistFunctionWith2Args17.Inputs("Arg2").Add oParamDistConstant10, "1"
        oParamDistFunctionWith2Args17.Parameter("Value") = 7.50999999999993E-02
        oParamDistFunctionWith2Args17.Parameter("Operation") = 1
        oParamDistFunctionWith2Args17.Evaluate
    
'        Dim oParamDistFunctionWith2Args18 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args18 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0018-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args18.Inputs("Arg1").Add oParamDistMeasureParallel6, "1"
        oParamDistFunctionWith2Args18.Inputs("Arg2").Add oParamDistConstant9, "1"
        oParamDistFunctionWith2Args18.Parameter("Value") = 0.175000000000003
        oParamDistFunctionWith2Args18.Parameter("Operation") = 1
        oParamDistFunctionWith2Args18.Evaluate
    
'        Dim oParamDistFunctionWith2Args19 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args19 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0019-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args19.Inputs("Arg1").Add oParamDistMeasureParallel6, "1"
        oParamDistFunctionWith2Args19.Inputs("Arg2").Add oParamDistConstant10, "1"
        oParamDistFunctionWith2Args19.Parameter("Value") = 7.51000000000028E-02
        oParamDistFunctionWith2Args19.Parameter("Operation") = 1
        oParamDistFunctionWith2Args19.Evaluate
    
        Dim oPointAlongCurve20 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve20 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0020-PointAlongCurve")
        oPointAlongCurve20.Inputs("Curve").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointAlongCurve20.Inputs("Point").Add oPointAtCurveExtremity11, "1"
        oPointAlongCurve20.Inputs("TrackPoint").Add oPointAtCurveMiddle7, "1"
        oPointAlongCurve20.Parameter("Distance") = 2#
        oPointAlongCurve20.Parameter("TrackFlag") = 2
        oPointAlongCurve20.Evaluate
    
        Dim oLineByPoints21 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints21 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0021-LineByPoints")
        oLineByPoints21.Inputs("StartPoint").Add oPointAlongCurve20, "1"
        oLineByPoints21.Inputs("EndPoint").Add oPointAtCurveMiddle7, "1"
        oLineByPoints21.Evaluate
    
        Dim oLineParallel22 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel22 = oGCFactory.CreateEntity("LineParallel", pPOM, "0022-LineParallel")
        oLineParallel22.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel22.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel22.Inputs("TrackPoint").Add oPointAtCurveMiddle7, "1"
        oLineParallel22.Parameter("Distance") = oParamDistFunctionWith2Args16
        oLineParallel22.Parameter("TrackFlag") = 1
        oLineParallel22.Evaluate
    
'        Dim oLineParallel23 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel23 = oGCFactory.CreateEntity("LineParallel", pPOM, "0023-LineParallel")
        oLineParallel23.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel23.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel23.Inputs("TrackPoint").Add oPointAtCurveMiddle7, "1"
        oLineParallel23.Parameter("Distance") = oParamDistFunctionWith2Args16
        oLineParallel23.Parameter("TrackFlag") = 2
        oLineParallel23.Evaluate
    
'        Dim oLineParallel24 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel24 = oGCFactory.CreateEntity("LineParallel", pPOM, "0024-LineParallel")
        oLineParallel24.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel24.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel24.Inputs("TrackPoint").Add oPointAtCurveMiddle7, "1"
        oLineParallel24.Parameter("Distance") = oParamDistFunctionWith2Args17
        oLineParallel24.Parameter("TrackFlag") = 2
        oLineParallel24.Evaluate
    
'        Dim oLineParallel25 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel25 = oGCFactory.CreateEntity("LineParallel", pPOM, "0025-LineParallel")
        oLineParallel25.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oLineParallel25.Inputs("Plane").Add oExtractPorts3.Output("Support", 1), "1"
        oLineParallel25.Inputs("TrackPoint").Add oPointAtCurveMiddle7, "1"
        oLineParallel25.Parameter("Distance") = oParamDistFunctionWith2Args17
        oLineParallel25.Parameter("TrackFlag") = 1
        oLineParallel25.Evaluate
    
'        Dim oLineParallel26 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel26 = oGCFactory.CreateEntity("LineParallel", pPOM, "0026-LineParallel")
        oLineParallel26.Inputs("Line").Add oLineByPoints21, "1"
        oLineParallel26.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel26.Inputs("TrackPoint").Add oPointAtCurveStart12, "1"
        oLineParallel26.Parameter("Distance") = oParamDistFunctionWith2Args18
        oLineParallel26.Parameter("TrackFlag") = 1
        oLineParallel26.Evaluate
    
        Dim oLineParallel27 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel27 = oGCFactory.CreateEntity("LineParallel", pPOM, "0027-LineParallel")
        oLineParallel27.Inputs("Line").Add oLineByPoints21, "1"
        oLineParallel27.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel27.Inputs("TrackPoint").Add oPointAtCurveStart12, "1"
        oLineParallel27.Parameter("Distance") = oParamDistFunctionWith2Args18
        oLineParallel27.Parameter("TrackFlag") = 2
        oLineParallel27.Evaluate
    
        Dim oLineParallel28 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel28 = oGCFactory.CreateEntity("LineParallel", pPOM, "0028-LineParallel")
        oLineParallel28.Inputs("Line").Add oLineByPoints21, "1"
        oLineParallel28.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel28.Inputs("TrackPoint").Add oPointAtCurveStart12, "1"
        oLineParallel28.Parameter("Distance") = oParamDistFunctionWith2Args19
        oLineParallel28.Parameter("TrackFlag") = 2
        oLineParallel28.Evaluate
    
        Dim oLineParallel29 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineParallel29 = oGCFactory.CreateEntity("LineParallel", pPOM, "0029-LineParallel")
        oLineParallel29.Inputs("Line").Add oLineByPoints21, "1"
        oLineParallel29.Inputs("Plane").Add oExtractPorts4.Output("Support", 1), "1"
        oLineParallel29.Inputs("TrackPoint").Add oPointAtCurveStart12, "1"
        oLineParallel29.Parameter("Distance") = oParamDistFunctionWith2Args19
        oLineParallel29.Parameter("TrackFlag") = 1
        oLineParallel29.Evaluate
    
        Dim oParamDistFunctionWith2Args30 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args30 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, "0030-ParamDistFunctionWith2Args")
        oParamDistFunctionWith2Args30.Inputs("Arg1").Add oParamDistConstant13, "1"
        oParamDistFunctionWith2Args30.Inputs("Arg2").Add oParamDistFunctionWith2Args19, "1"
        oParamDistFunctionWith2Args30.Parameter("Value") = 1.0751
        oParamDistFunctionWith2Args30.Parameter("Operation") = 1
        oParamDistFunctionWith2Args30.Evaluate
    
        Dim oPointByCurves31 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves31 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0031-PointByCurves")
        oPointByCurves31.Inputs("Curve1").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPointByCurves31.Inputs("Curve2").Add oLineParallel22, "1"
        oPointByCurves31.Parameter("TrackFlag") = 1
        oPointByCurves31.Evaluate
    
        Dim oPointByCurves32 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves32 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0032-PointByCurves")
        oPointByCurves32.Inputs("Curve1").Add oLineParallel28, "1"
        oPointByCurves32.Inputs("Curve2").Add oLineParallel22, "1"
        oPointByCurves32.Parameter("TrackFlag") = 1
        oPointByCurves32.Evaluate
    
        Dim oPointByCurves33 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves33 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0033-PointByCurves")
        oPointByCurves33.Inputs("Curve1").Add oLineParallel29, "2"
        oPointByCurves33.Inputs("Curve2").Add oLineParallel22, "1"
        oPointByCurves33.Parameter("TrackFlag") = 1
        oPointByCurves33.Evaluate
    
        Dim oPointByCurves34 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves34 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0034-PointByCurves")
        oPointByCurves34.Inputs("Curve1").Add oLineParallel29, "1"
        oPointByCurves34.Inputs("Curve2").Add oLineParallel23, "1"
        oPointByCurves34.Parameter("TrackFlag") = 1
        oPointByCurves34.Evaluate
    
'        Dim oPointByCurves35 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves35 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0035-PointByCurves")
        oPointByCurves35.Inputs("Curve1").Add oLineParallel28, "1"
        oPointByCurves35.Inputs("Curve2").Add oLineParallel23, "1"
        oPointByCurves35.Parameter("TrackFlag") = 1
        oPointByCurves35.Evaluate
    
'        Dim oPointByCurves36 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves36 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0036-PointByCurves")
        oPointByCurves36.Inputs("Curve1").Add oLineParallel24, "2"
        oPointByCurves36.Inputs("Curve2").Add oLineParallel26, "1"
        oPointByCurves36.Parameter("TrackFlag") = 1
        oPointByCurves36.Evaluate
    
'        Dim oPointByCurves37 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves37 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0037-PointByCurves")
        oPointByCurves37.Inputs("Curve1").Add oLineParallel24, "1"
        oPointByCurves37.Inputs("Curve2").Add oLineParallel27, "1"
        oPointByCurves37.Parameter("TrackFlag") = 1
        oPointByCurves37.Evaluate
    
'        Dim oPointByCurves38 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves38 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0038-PointByCurves")
        oPointByCurves38.Inputs("Curve1").Add oLineParallel25, "2"
        oPointByCurves38.Inputs("Curve2").Add oLineParallel26, "1"
        oPointByCurves38.Parameter("TrackFlag") = 1
        oPointByCurves38.Evaluate
    
'        Dim oPointByCurves39 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves39 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0039-PointByCurves")
        oPointByCurves39.Inputs("Curve1").Add oLineParallel25, "1"
        oPointByCurves39.Inputs("Curve2").Add oLineParallel27, "1"
        oPointByCurves39.Parameter("TrackFlag") = 1
        oPointByCurves39.Evaluate
    
'        Dim oPointByCurves40 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves40 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0040-PointByCurves")
        oPointByCurves40.Inputs("Curve1").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPointByCurves40.Inputs("Curve2").Add oLineParallel26, "1"
        oPointByCurves40.Parameter("TrackFlag") = 1
        oPointByCurves40.Evaluate
    
'        Dim oPointByCurves41 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves41 = oGCFactory.CreateEntity("PointByCurves", pPOM, "0041-PointByCurves")
        oPointByCurves41.Inputs("Curve1").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPointByCurves41.Inputs("Curve2").Add oLineParallel27, "1"
        oPointByCurves41.Parameter("TrackFlag") = 1
        oPointByCurves41.Evaluate
    
        Dim oPlaneByPointNormal42 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal42 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0042-PlaneByPointNormal")
        oPlaneByPointNormal42.Inputs("Point").Add oPointByCurves31, "1"
        oPlaneByPointNormal42.Inputs("Line").Add oExtractPorts4.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal42.Parameter("Range") = oParamDistFunctionWith2Args30
        oPlaneByPointNormal42.Evaluate
    
        Dim oPlaneByPointNormal43 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal43 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0043-PlaneByPointNormal")
        oPlaneByPointNormal43.Inputs("Point").Add oPointByCurves40, "1"
        oPlaneByPointNormal43.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal43.Parameter("Range") = oParamDistFunctionWith2Args30
        oPlaneByPointNormal43.Evaluate
    
        Dim oPlaneByPointNormal44 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal44 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0044-PlaneByPointNormal")
        oPlaneByPointNormal44.Inputs("Point").Add oPointByCurves41, "2"
        oPlaneByPointNormal44.Inputs("Line").Add oExtractPorts3.Output("MemberAxis", 1), "1"
        oPlaneByPointNormal44.Parameter("Range") = oParamDistFunctionWith2Args30
        oPlaneByPointNormal44.Evaluate
    
        Dim oSurfFromGType45 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType45 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0045-SurfFromGType")
        oSurfFromGType45.Inputs("Surface").Add oPlaneByPointNormal42, "1"
        oSurfFromGType45.Evaluate
    
'        Dim oLineByPoints46 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints46 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0046-LineByPoints")
        oLineByPoints46.Inputs("StartPoint").Add oPointByCurves35, "1"
        oLineByPoints46.Inputs("EndPoint").Add oPointByCurves37, "1"
        oLineByPoints46.Evaluate
    
'        Dim oPointAlongCurve47 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve47 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0047-PointAlongCurve")
        oPointAlongCurve47.Inputs("Curve").Add oLineByPoints46, "2"
        oPointAlongCurve47.Inputs("Point").Add oPointByCurves37, "2"
        oPointAlongCurve47.Inputs("TrackPoint").Add oPointByCurves35, "2"
        oPointAlongCurve47.Parameter("Distance") = 0.05
        oPointAlongCurve47.Parameter("TrackFlag") = 2
        oPointAlongCurve47.Evaluate
    
        Dim oLineByPoints48 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints48 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0048-LineByPoints")
        oLineByPoints48.Inputs("StartPoint").Add oPointByCurves32, "2"
        oLineByPoints48.Inputs("EndPoint").Add oPointByCurves39, "1"
        oLineByPoints48.Evaluate
    
        Dim oPointAlongCurve49 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0049-PointAlongCurve")
        oPointAlongCurve49.Inputs("Curve").Add oLineByPoints48, "2"
        oPointAlongCurve49.Inputs("Point").Add oPointByCurves32, "2"
        oPointAlongCurve49.Inputs("TrackPoint").Add oPointByCurves39, "2"
        oPointAlongCurve49.Parameter("Distance") = 0.05
        oPointAlongCurve49.Parameter("TrackFlag") = 2
        oPointAlongCurve49.Evaluate
    
'        Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0050-PointAlongCurve")
        oPointAlongCurve50.Inputs("Curve").Add oLineByPoints48, "1"
        oPointAlongCurve50.Inputs("Point").Add oPointByCurves39, "2"
        oPointAlongCurve50.Inputs("TrackPoint").Add oPointByCurves32, "1"
        oPointAlongCurve50.Parameter("Distance") = 0.05
        oPointAlongCurve50.Parameter("TrackFlag") = 2
        oPointAlongCurve50.Evaluate
    
'        Dim oLineByPoints51 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints51 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0051-LineByPoints")
        oLineByPoints51.Inputs("StartPoint").Add oPointByCurves38, "2"
        oLineByPoints51.Inputs("EndPoint").Add oPointByCurves33, "1"
        oLineByPoints51.Evaluate
    
'        Dim oPointAlongCurve52 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve52 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0052-PointAlongCurve")
        oPointAlongCurve52.Inputs("Curve").Add oLineByPoints51, "1"
        oPointAlongCurve52.Inputs("Point").Add oPointByCurves38, "1"
        oPointAlongCurve52.Inputs("TrackPoint").Add oPointByCurves33, "1"
        oPointAlongCurve52.Parameter("Distance") = 0.05
        oPointAlongCurve52.Parameter("TrackFlag") = 2
        oPointAlongCurve52.Evaluate
    
        Dim oPointAlongCurve53 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0053-PointAlongCurve")
        oPointAlongCurve53.Inputs("Curve").Add oLineByPoints51, "1"
        oPointAlongCurve53.Inputs("Point").Add oPointByCurves33, "2"
        oPointAlongCurve53.Inputs("TrackPoint").Add oPointByCurves38, "1"
        oPointAlongCurve53.Parameter("Distance") = 0.05
        oPointAlongCurve53.Parameter("TrackFlag") = 2
        oPointAlongCurve53.Evaluate
    
'        Dim oLineByPoints54 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints54 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0054-LineByPoints")
        oLineByPoints54.Inputs("StartPoint").Add oPointByCurves36, "1"
        oLineByPoints54.Inputs("EndPoint").Add oPointByCurves34, "1"
        oLineByPoints54.Evaluate
    
        Dim oPointAlongCurve55 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve55 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0055-PointAlongCurve")
        oPointAlongCurve55.Inputs("Curve").Add oLineByPoints54, "2"
        oPointAlongCurve55.Inputs("Point").Add oPointByCurves36, "2"
        oPointAlongCurve55.Inputs("TrackPoint").Add oPointByCurves34, "2"
        oPointAlongCurve55.Parameter("Distance") = 0.05
        oPointAlongCurve55.Parameter("TrackFlag") = 2
        oPointAlongCurve55.Evaluate
    
'        Dim oLineByPoints56 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints56 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0056-LineByPoints")
        oLineByPoints56.Inputs("StartPoint").Add oPointAlongCurve49, "2"
        oLineByPoints56.Inputs("EndPoint").Add oPointAlongCurve50, "2"
        oLineByPoints56.Evaluate
    
'        Dim oLineByPoints57 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints57 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0057-LineByPoints")
        oLineByPoints57.Inputs("StartPoint").Add oPointAlongCurve52, "1"
        oLineByPoints57.Inputs("EndPoint").Add oPointAlongCurve53, "1"
        oLineByPoints57.Evaluate
    
        Dim oSurfFromGType58 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType58 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0058-SurfFromGType")
        oSurfFromGType58.Inputs("Surface").Add oPlaneByPointNormal43, "2"
        oSurfFromGType58.Evaluate
    
        Dim oSurfFromGType59 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType59 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0059-SurfFromGType")
        oSurfFromGType59.Inputs("Surface").Add oPlaneByPointNormal44, "2"
        oSurfFromGType59.Evaluate
    
        Dim oPointFromCS60 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointFromCS60 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0060-PointFromCS")
        oPointFromCS60.Inputs("CoordinateSystem").Add oCSByStiffener14, "1"
        oPointFromCS60.Parameter("X") = 0#
        oPointFromCS60.Parameter("Y") = 0#
        oPointFromCS60.Parameter("Z") = 0#
        oPointFromCS60.Evaluate
    
        Dim oPointAlongCurve61 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve61 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "0061-PointAlongCurve")
        oPointAlongCurve61.Inputs("Curve").Add oOrderAxes21.Output("Normal", 1), "1"
        oPointAlongCurve61.Inputs("Point").Add oOrderAxes21.Output("Node", 1), "1"
        oPointAlongCurve61.Inputs("TrackPoint").Add oPointFromCS60, "1"
        oPointAlongCurve61.Parameter("Distance") = 0.3
        oPointAlongCurve61.Parameter("TrackFlag") = 1
        oPointAlongCurve61.Evaluate
    
        Dim oLineByPoints62 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints62 = oGCFactory.CreateEntity("LineByPoints", pPOM, "0062-LineByPoints")
        oLineByPoints62.Inputs("StartPoint").Add oOrderAxes21.Output("Node", 1), "1"
        oLineByPoints62.Inputs("EndPoint").Add oPointAlongCurve61, "1"
        oLineByPoints62.Evaluate
    
'        Dim oPlaneByPointNormal63 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPointNormal63 = oGCFactory.CreateEntity("PlaneByPointNormal", pPOM, "0063-PlaneByPointNormal")
        oPlaneByPointNormal63.Inputs("Point").Add oOrderAxes21.Output("Node", 1), "1"
        oPlaneByPointNormal63.Inputs("Line").Add oLineByPoints62, "1"
        oPlaneByPointNormal63.Parameter("Range") = 1#
        oPlaneByPointNormal63.Evaluate
    
        Dim oSurfFromGType64 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfFromGType64 = oGCFactory.CreateEntity("SurfFromGType", pPOM, "0064-SurfFromGType")
        oSurfFromGType64.Inputs("Surface").Add oPlaneByPointNormal63, "1"
        oSurfFromGType64.Evaluate
        
'        Dim pElementsOfSecondaryMembers As IJElements
        Set pElementsOfSecondaryMembers = pGeometricConstruction.Inputs("SecondaryProfile1")
'        Dim iCountOfSecondaryMembers As Integer
        Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
        If iCountOfSecondaryMembers > 0 Then
        
            Dim oCSFromMember15 As SP3DGeometricConstruction.GeometricConstruction
            Set oCSFromMember15 = oGCFactory.CreateEntity("CSFromMember", pPOM, "0015-CSFromMember")
            oCSFromMember15.Inputs("MemberPart").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "1"
            oCSFromMember15.Evaluate
        
            Dim oFacePortExtractor65 As SP3DGeometricConstruction.GeometricConstruction
            Set oFacePortExtractor65 = oGCFactory.CreateEntity("FacePortExtractor", pPOM, "0065-FacePortExtractor")
            oFacePortExtractor65.Inputs("Connectable").Add pGeometricConstruction.Inputs("SecondaryProfile1").Item(1), "2"
            oFacePortExtractor65.Inputs("CoordinateSystem").Add oCSFromMember15, "1"
            oFacePortExtractor65.Parameter("GeometrySelector") = 4
            oFacePortExtractor65.Parameter("FacesContext") = 2
            oFacePortExtractor65.Parameter("LookingAxis") = 2
            oFacePortExtractor65.Parameter("IntersectingPlane") = 0
            oFacePortExtractor65.Parameter("SurfaceType") = 1
            oFacePortExtractor65.Parameter("TrackFlag") = 1
            oFacePortExtractor65.Parameter("Offset") = 0#
            oFacePortExtractor65.Evaluate
        
            pGeometricConstruction.ControlledInputs("Port_69_1").Clear
            pGeometricConstruction.ControlledInputs("Port_69_1").Add oFacePortExtractor65.ControlledInputs("Port").Item("1")
        
        End If
        
        Dim oSurfByLinearExtrusion66 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion66 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0066-SurfByLinearExtrusion")
        oSurfByLinearExtrusion66.Inputs("PlanarCrossSection").Add oLineByPoints57, "1"
        oSurfByLinearExtrusion66.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion66.Evaluate
    
        Dim oSurfByLinearExtrusion67 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion67 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0067-SurfByLinearExtrusion")
        oSurfByLinearExtrusion67.Inputs("PlanarCrossSection").Add oLineByPoints56, "1"
        oSurfByLinearExtrusion67.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion67.Evaluate
    
        Dim oLineStringByPoints68 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineStringByPoints68 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "0068-LineStringByPoints")
        oLineStringByPoints68.Inputs("Points").Add oPointAlongCurve55, "1"
        oLineStringByPoints68.Inputs("Points").Add oPointByCurves34, "2"
        oLineStringByPoints68.Inputs("Points").Add oPointByCurves35, "3"
        oLineStringByPoints68.Inputs("Points").Add oPointAlongCurve47, "4"
        oLineStringByPoints68.Parameter("ClosedFlag") = 0
        oLineStringByPoints68.Evaluate
    
        Dim oSurfByLinearExtrusion69 As SP3DGeometricConstruction.GeometricConstruction
        Set oSurfByLinearExtrusion69 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "0069-SurfByLinearExtrusion")
        oSurfByLinearExtrusion69.Inputs("PlanarCrossSection").Add oLineStringByPoints68, "1"
        oSurfByLinearExtrusion69.Inputs("ExtrusionLine").Add oLineFromCS2, "1"
        oSurfByLinearExtrusion69.Evaluate
    
'        Dim oGCMacro As IJGeometricConstructionMacro
        Set oGCMacro = pGeometricConstruction
    
        oGCMacro.Output("Support", 1) = oSurfFromGType64.Output
        oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion69.Output
        oGCMacro.Output("Boundary", 2) = oSurfByLinearExtrusion66.Output
        oGCMacro.Output("Boundary", 3) = oSurfByLinearExtrusion67.Output
        oGCMacro.Output("Boundary", 4) = oSurfFromGType45.Output
        oGCMacro.Output("Boundary", 5) = oSurfFromGType58.Output
        oGCMacro.Output("Boundary", 6) = oSurfFromGType59.Output
        oGCMacro.Output("CoordinateSystem", 1) = oOrderAxes21.Output("CoordinateSystem", 1)
        oGCMacro.Output("Node", 1) = oOrderAxes21.Output("Node", 1)
    
    End If
    
End Sub
