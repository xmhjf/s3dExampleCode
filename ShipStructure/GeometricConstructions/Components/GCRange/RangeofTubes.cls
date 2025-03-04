VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeOfTubes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("PrimaryMember", "Select a MemberPart", "ISPSMemberPartPrismatic OR ISPSDesignedMember OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddInput("Tubes", "Select a Tube", "ISPSMemberPartPrismatic or ISPSDesignedMember", 1, 100, "IJForeignIFC")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_5_1_27_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_6_1_27_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_7_1_27_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_8_1_27_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_6_1_35_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_8_1_35_1", "IJGeometry")
    Call pGeometricConstructionDefinition.AddOutput(5, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(2, "Lines")
    Call pGeometricConstructionDefinition.AddOutput(1, "Points")
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oRangeOfTubes() As SP3DGeometricConstruction.GeometricConstruction
    ReDim oRangeOfTubes(pGeometricConstruction.Inputs("Tubes").Count)
    Dim oLineFromCS3 As SP3DGeometricConstruction.GeometricConstruction
    Dim oLineFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointFromCS5 As SP3DGeometricConstruction.GeometricConstruction
    
    Dim oParamDistFunctionWithNArgs21 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs21 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "021-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs21.Parameter("Operation") = 1
    
    Dim oParamDistFunctionWithNArgs22 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs22 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "022-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs22.Parameter("Operation") = 2

    Dim oParamDistFunctionWithNArgs23 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs23 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "023-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs23.Parameter("Operation") = 1

    Dim oParamDistFunctionWithNArgs24 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs24 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "024-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs24.Parameter("Operation") = 2
    
    pGeometricConstruction.ControlledInputs("Port_5_1_27_1").Clear
    pGeometricConstruction.ControlledInputs("Port_6_1_27_1").Clear
    pGeometricConstruction.ControlledInputs("Port_7_1_27_1").Clear
    pGeometricConstruction.ControlledInputs("Port_8_1_27_1").Clear
    Dim i As Integer
    For i = 1 To pGeometricConstruction.Inputs("Tubes").Count
        Set oRangeOfTubes(i) = oGCFactory.CreateEntity("RangeOfTube", pPOM, "001." + CStr(i) + "-RangeOfTube")
        oRangeOfTubes(i).Inputs("PrimaryMember").Add pGeometricConstruction.Inputs("PrimaryMember").Item(1), "1"
        oRangeOfTubes(i).Inputs("Tube").Add pGeometricConstruction.Inputs("Tubes").Item(i), "1"
        If i > 1 Then
            oRangeOfTubes(i).Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1), "1"
        End If
        oRangeOfTubes(i).Evaluate
        If i = 1 Then
            pGeometricConstruction.ControlledInputs("Port_5_1_27_1").Add oRangeOfTubes(i).ControlledInputs("Port_5_1").Item(1)
            pGeometricConstruction.ControlledInputs("Port_7_1_27_1").Add oRangeOfTubes(i).ControlledInputs("Port_7_1").Item(1)
        End If
        pGeometricConstruction.ControlledInputs("Port_6_1_27_1").Add oRangeOfTubes(i).ControlledInputs("Port_6_1").Item(1)
        pGeometricConstruction.ControlledInputs("Port_8_1_27_1").Add oRangeOfTubes(i).ControlledInputs("Port_8_1").Item(1)

        If i = 1 Then
            Set oLineFromCS3 = oGCFactory.CreateEntity("LineFromCS", pPOM, "003-LineFromCS")
            oLineFromCS3.Inputs("CoordinateSystem").Add oRangeOfTubes(i).Output("CoordinateSystem", 1)
            oLineFromCS3.Parameter("LookingAxis") = 1
            oLineFromCS3.Parameter("Length") = -10#
            oLineFromCS3.Parameter("CSOrientation") = 1
            oLineFromCS3.Parameter("LineJustification") = 1
            oLineFromCS3.Evaluate

            Set oLineFromCS4 = oGCFactory.CreateEntity("LineFromCS", pPOM, "004-LineFromCS")
            oLineFromCS4.Inputs("CoordinateSystem").Add oRangeOfTubes(i).Output("CoordinateSystem", 1)
            oLineFromCS4.Parameter("LookingAxis") = 2
            oLineFromCS4.Parameter("Length") = -10#
            oLineFromCS4.Parameter("CSOrientation") = 1
            oLineFromCS4.Parameter("LineJustification") = 1
            oLineFromCS4.Evaluate
        
            Set oPointFromCS5 = oGCFactory.CreateEntity("PointFromCS", pPOM, "005-PointFromCS")
            oPointFromCS5.Inputs("CoordinateSystem").Add oRangeOfTubes(i).Output("CoordinateSystem", 1)
            oPointFromCS5.Parameter("X") = 0#
            oPointFromCS5.Parameter("Y") = 0#
            oPointFromCS5.Parameter("Z") = 0#
            oPointFromCS5.Evaluate
        End If
    
        Dim oPointAlongCurve6 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve6 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "006." + CStr(i) + "-PointAlongCurve")
        oPointAlongCurve6.Inputs("Curve").Add oLineFromCS3
        oPointAlongCurve6.Inputs("Point").Add oRangeOfTubes(i).Output("Points", "Low")
        oPointAlongCurve6.Parameter("Distance") = 0#
        oPointAlongCurve6.Parameter("TrackFlag") = 1
        oPointAlongCurve6.Evaluate
    
        Dim oPointAlongCurve7 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve7 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "007." + CStr(i) + "-PointAlongCurve")
        oPointAlongCurve7.Inputs("Curve").Add oLineFromCS3
        oPointAlongCurve7.Inputs("Point").Add oRangeOfTubes(i).Output("Points", "High")
        oPointAlongCurve7.Parameter("Distance") = 0#
        oPointAlongCurve7.Parameter("TrackFlag") = 1
        oPointAlongCurve7.Evaluate
    
        Dim oParamDistMeasureParallel10 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel10 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "010." + CStr(i) + "-ParamDistMeasureParallel")
        oParamDistMeasureParallel10.Inputs("Graphics").Add oPointFromCS5
        oParamDistMeasureParallel10.Inputs("Graphics").Add oPointAlongCurve6
        oParamDistMeasureParallel10.Inputs("OrientationLine").Add oLineFromCS3
        oParamDistMeasureParallel10.Parameter("Value") = -0.1942263
        oParamDistMeasureParallel10.Evaluate
    
        Dim oParamDistMeasureParallel11 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel11 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "011." + CStr(i) + "-ParamDistMeasureParallel")
        oParamDistMeasureParallel11.Inputs("Graphics").Add oPointFromCS5
        oParamDistMeasureParallel11.Inputs("Graphics").Add oPointAlongCurve7
        oParamDistMeasureParallel11.Inputs("OrientationLine").Add oLineFromCS3
        oParamDistMeasureParallel11.Parameter("Value") = 0.1942263
        oParamDistMeasureParallel11.Evaluate
        
        Dim oPointAlongCurve14 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve14 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "014." + CStr(i) + "-PointAlongCurve")
        oPointAlongCurve14.Inputs("Curve").Add oLineFromCS4
        oPointAlongCurve14.Inputs("Point").Add oRangeOfTubes(i).Output("Points", "Low")
        oPointAlongCurve14.Parameter("Distance") = 0#
        oPointAlongCurve14.Parameter("TrackFlag") = 1
        oPointAlongCurve14.Evaluate
    
        Dim oPointAlongCurve15 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve15 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "015." + CStr(i) + "-PointAlongCurve")
        oPointAlongCurve15.Inputs("Curve").Add oLineFromCS4
        oPointAlongCurve15.Inputs("Point").Add oRangeOfTubes(i).Output("Points", "High")
        oPointAlongCurve15.Parameter("Distance") = 0#
        oPointAlongCurve15.Parameter("TrackFlag") = 1
        oPointAlongCurve15.Evaluate
        
        Dim oParamDistMeasureParallel18 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel18 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "018." + CStr(i) + "-ParamDistMeasureParallel")
        oParamDistMeasureParallel18.Inputs("Graphics").Add oPointFromCS5
        oParamDistMeasureParallel18.Inputs("Graphics").Add oPointAlongCurve14
        oParamDistMeasureParallel18.Inputs("OrientationLine").Add oLineFromCS4
        oParamDistMeasureParallel18.Parameter("Value") = -0.175
        oParamDistMeasureParallel18.Evaluate
    
        Dim oParamDistMeasureParallel19 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel19 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "019." + CStr(i) + "-ParamDistMeasureParallel")
        oParamDistMeasureParallel19.Inputs("Graphics").Add oPointFromCS5
        oParamDistMeasureParallel19.Inputs("Graphics").Add oPointAlongCurve15
        oParamDistMeasureParallel19.Inputs("OrientationLine").Add oLineFromCS4
        oParamDistMeasureParallel19.Parameter("Value") = 0.175
        oParamDistMeasureParallel19.Evaluate
    
        oParamDistFunctionWithNArgs21.Inputs("Args").Add oParamDistMeasureParallel10
        oParamDistFunctionWithNArgs22.Inputs("Args").Add oParamDistMeasureParallel11
        oParamDistFunctionWithNArgs23.Inputs("Args").Add oParamDistMeasureParallel18
        oParamDistFunctionWithNArgs24.Inputs("Args").Add oParamDistMeasureParallel19
    Next
    
    oParamDistFunctionWithNArgs21.Evaluate
    oParamDistFunctionWithNArgs22.Evaluate
    oParamDistFunctionWithNArgs23.Evaluate
    oParamDistFunctionWithNArgs24.Evaluate

    Dim oPointFromCS26 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS26 = oGCFactory.CreateEntity("PointFromCS", pPOM, "026-PointFromCS")
    oPointFromCS26.Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1)
    oPointFromCS26.Parameter("X") = oParamDistFunctionWithNArgs21
    oPointFromCS26.Parameter("Y") = oParamDistFunctionWithNArgs23
    oPointFromCS26.Parameter("Z") = 0#
    oPointFromCS26.Evaluate

    Dim oPointFromCS27 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS27 = oGCFactory.CreateEntity("PointFromCS", pPOM, "027-PointFromCS")
    oPointFromCS27.Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1)
    oPointFromCS27.Parameter("X") = oParamDistFunctionWithNArgs22
    oPointFromCS27.Parameter("Y") = oParamDistFunctionWithNArgs23
    oPointFromCS27.Parameter("Z") = 0#
    oPointFromCS27.Evaluate

    Dim oPointFromCS28 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS28 = oGCFactory.CreateEntity("PointFromCS", pPOM, "028-PointFromCS")
    oPointFromCS28.Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1)
    oPointFromCS28.Parameter("X") = oParamDistFunctionWithNArgs22
    oPointFromCS28.Parameter("Y") = oParamDistFunctionWithNArgs24
    oPointFromCS28.Parameter("Z") = 0#
    oPointFromCS28.Evaluate

    Dim oPointFromCS29 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS29 = oGCFactory.CreateEntity("PointFromCS", pPOM, "029-PointFromCS")
    oPointFromCS29.Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1)
    oPointFromCS29.Parameter("X") = oParamDistFunctionWithNArgs21
    oPointFromCS29.Parameter("Y") = oParamDistFunctionWithNArgs24
    oPointFromCS29.Parameter("Z") = 0#
    oPointFromCS29.Evaluate

    Dim oLineByPoints30 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints30 = oGCFactory.CreateEntity("LineByPoints", pPOM, "030-LineByPoints")
    oLineByPoints30.Inputs("StartPoint").Add oPointFromCS26
    oLineByPoints30.Inputs("EndPoint").Add oPointFromCS29
    oLineByPoints30.Evaluate

    Dim oLineByPoints31 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints31 = oGCFactory.CreateEntity("LineByPoints", pPOM, "031-LineByPoints")
    oLineByPoints31.Inputs("StartPoint").Add oPointFromCS27
    oLineByPoints31.Inputs("EndPoint").Add oPointFromCS28
    oLineByPoints31.Evaluate

    Dim oLineByPoints32 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints32 = oGCFactory.CreateEntity("LineByPoints", pPOM, "032-LineByPoints")
    oLineByPoints32.Inputs("StartPoint").Add oPointFromCS26
    oLineByPoints32.Inputs("EndPoint").Add oPointFromCS27
    oLineByPoints32.Evaluate

    Dim oLineByPoints33 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints33 = oGCFactory.CreateEntity("LineByPoints", pPOM, "033-LineByPoints")
    oLineByPoints33.Inputs("StartPoint").Add oPointFromCS29
    oLineByPoints33.Inputs("EndPoint").Add oPointFromCS28
    oLineByPoints33.Evaluate

    Dim oLineByPoints34 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints34 = oGCFactory.CreateEntity("LineByPoints", pPOM, "034-LineByPoints")
    oLineByPoints34.Inputs("StartPoint").Add oPointFromCS26
    oLineByPoints34.Inputs("EndPoint").Add oPointFromCS28
    oLineByPoints34.Evaluate

    Dim oPointAtCurveMiddle35 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle35 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "035-PointAtCurveMiddle")
    oPointAtCurveMiddle35.Inputs("Curve").Add oLineByPoints34
    oPointAtCurveMiddle35.Evaluate

    Dim oCSByCS36 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS36 = oGCFactory.CreateEntity("CSByCS", pPOM, "036-CSByCS")
    oCSByCS36.Inputs("CoordinateSystem").Add oRangeOfTubes(1).Output("CoordinateSystem", 1)
    oCSByCS36.Inputs("Origin").Add oPointAtCurveMiddle35
    oCSByCS36.Parameter("AxesRoles") = 1
    oCSByCS36.Parameter("CSOrientation") = 1
    oCSByCS36.Parameter("TrackFlag") = 1
    oCSByCS36.Evaluate

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("CoordinateSystem", 1) = oCSByCS36.Output
    oGCMacro.Output("Lines", "XLow") = oLineByPoints30.Output
    oGCMacro.Output("Lines", "XHigh") = oLineByPoints31.Output
    oGCMacro.Output("Lines", "YLow") = oLineByPoints32.Output
    oGCMacro.Output("Lines", "YHigh") = oLineByPoints33.Output
    oGCMacro.Output("Points", "Low") = oPointFromCS26.Output
    oGCMacro.Output("Points", "High") = oPointFromCS28.Output
End Sub
