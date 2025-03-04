VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckledBracket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const sSTIFFENER_SYSTEM1 = "StiffenerSystem1"
Const sSTIFFENER_SYSTEM2 = "StiffenerSystem2"
Const sPLATE_SYSTEM = "PlateSystem"
Const sCOORDINATE_SYSTEM = "CoordinateSystem"
Const sNO_LOCATE = "NoLocate"
Const sNO_LOCATE_PROMPT_ = "Auto preview is off"

Const sPREVIEW = "Preview"
Const sPREVIEW_PROMPT_ = "Auto preview"

Const sORIENTATION = "Orientation"
Const sWIDTH = "Width"
Const sKNUCKLE_TYPE = "KnuckleType"
Const sKNUCKLE_HEIGHT = "KnuckleHeight"
Const sJUSTIFICATION1 = "Justification1"
Const sJUSTIFICATION2 = "Justification2"
Const sOFFSET1 = "Offset1"
Const sOFFSET2 = "Offset2"

Private Enum errors
    lWIDTH_TOO_SMALL = 2
    lKNUCKLE_HEIGHT_TOO_SMALL = 3
    lKNUCKLE_HEIGHT_TOO_BIG = 4
    lSTIFFENER_AXIS1_NOT_INTERSECTING = 5
    lSTIFFENER_AXIS2_NOT_INTERSECTING = 6
    lSTIFFENER_AXES_TOO_CLOSE = 7
    lOFFSET1_BIGGER_THAN_FLANGE_WIDTH = 15
    lOFFSET2_BIGGER_THAN_FLANGE_WIDTH = 16
    lOFFSET1_BIGGER_THAN_HALF_FLANGE_WIDTH = 17
    lOFFSET2_BIGGER_THAN_HALF_FLANGE_WIDTH = 18
    lCANNOT_EVALUATE_KNUCKLED_PLANE = 19
    lCANNOT_CREATE_RIBBON_SURFACE = 23
    lCANNOT_CREATE_CS_AT_POINT_ON_STIFFENER = 24
    lSTIFFENERS_IDENTICAL = 26
    lCANNOT_GENERATE_SURFACE_FROM_PLATE = 27
    lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM1 = 28
    lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM2 = 29
    lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER1 = 30
    lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER2 = 31
    lCANNOT_POSITION_POINT_ALONG_STIFFENER1_AT_WIDTH = 32
    lCANNOT_POSITION_POINT_ALONG_STIFFENER2_AT_WIDTH = 33
    lSTIFFENERS_ON_SAME_UNDERLYING_PLATE_SYSTEM = 34
    lSTIFFENER1_ON_SUPPORTING_PLATE_SYSTEM = 35
    lSTIFFENER2_ON_SUPPORTING_PLATE_SYSTEM = 36
    lCANNOT_CREATE_TUBE_FROM_PROFILE1 = 37
    lCANNOT_CREATE_TUBE_FROM_PROFILE2 = 38
    lINVALID_ORIENTATION_AXIS = 39
    
    lUNEXPECTED_INTERNAL_ERROR = 100
End Enum


Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGeometricConstructionDynamic_IsOptional
Implements IJGCMirror
Implements IJKnuckledBracket
Implements IJGCToDoDelegate

Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As IJGeometricConstruction, ByVal vInfo As Variant, ByRef bIsOnRibbonBar As Boolean)
    bIsOnRibbonBar = True
    Select Case sName
        Case sNO_LOCATE: bIsOnRibbonBar = pGC.Parameter(sPREVIEW) = 0
        Case sKNUCKLE_HEIGHT: bIsOnRibbonBar = pGC.Parameter(sKNUCKLE_TYPE) = 0
    End Select
End Sub
Private Sub IJGeometricConstructionDynamic_IsOptional_PropertyValue(ByVal sName As String, ByVal pGC As IJGeometricConstruction, ByVal vInfo As Variant, ByRef bIsOptional As Boolean)
    bIsOptional = False
    Select Case sName
        Case sNO_LOCATE: bIsOptional = pGC.Parameter(sPREVIEW) = 1
    End Select
End Sub
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput(sSTIFFENER_SYSTEM1, "Select first stiffener system", "IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJGeometry IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    Call pGeometricConstructionDefinition.AddInput(sSTIFFENER_SYSTEM2, "Select second stiffener system", "IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, "IJGeometry IJSplitNotify IJDProfileMoldedConventions IJStiffenerAttachmentMethod")
    Call pGeometricConstructionDefinition.AddInput(sPLATE_SYSTEM, "Select supporting plate system", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddInput(sCOORDINATE_SYSTEM, "Select coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput(sNO_LOCATE, sNO_LOCATE_PROMPT_, "[GCKnuckledBracket.KnuckledBracket,RejectEverything]", 0, 1)
    
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort1")
    Call pGeometricConstructionDefinition.AddControlledInput("AxisPort2")
'   Call pGeometricConstructionDefinition.AddControlledInput("BKHFacePort0")
'   Call pGeometricConstructionDefinition.AddControlledInput("LowerDeckFacePort0")
'   Call pGeometricConstructionDefinition.AddControlledInput("UppderDeckFacePort0")
'    Call pGeometricConstructionDefinition.AddControlledInput("FacePort1")
'    Call pGeometricConstructionDefinition.AddControlledInput("FacePort2")
    Call pGeometricConstructionDefinition.AddControlledInput("FacePort")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem1", "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem1", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("RootPlateSystem2", "IJSplitNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("LeafPlateSystem2", "IJPlate IJDPlateMoldedConventions")
    Call pGeometricConstructionDefinition.AddControlledInput("SupportLeafPlateSystem", "IJPlate IJDPlateMoldedConventions")
    
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "")
    
    Call pGeometricConstructionDefinition.AddParameter(sPREVIEW, sPREVIEW_PROMPT_, GCCodeList, 0, 0, 0, 0, 1)
    If True Then
        Call pGeometricConstructionDefinition.AddParameterValue(sPREVIEW, "Off", 0)
        Call pGeometricConstructionDefinition.AddParameterValue(sPREVIEW, "On", 1)
    End If
    
    Call pGeometricConstructionDefinition.AddParameter(sORIENTATION, sORIENTATION, GCCodeList, 0, 0, 0, 0, 1)
    If True Then
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "+x", 1)
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "-x", 2)
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "+y", 3)
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "-y", 4)
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "+z", 5)
        Call pGeometricConstructionDefinition.AddParameterValue(sORIENTATION, "-z", 6)
    End If
    
    Call pGeometricConstructionDefinition.AddParameter(sWIDTH, sWIDTH, GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 2#)
    Call pGeometricConstructionDefinition.AddParameter(sKNUCKLE_TYPE, "Knuckle type", GCCodeList, 0, 0, 0, 0, 0)
    If True Then
        Call pGeometricConstructionDefinition.AddParameterValue(sKNUCKLE_TYPE, "Partial", 0)
        Call pGeometricConstructionDefinition.AddParameterValue(sKNUCKLE_TYPE, "Full", 1)
    End If
    Call pGeometricConstructionDefinition.AddParameter(sKNUCKLE_HEIGHT, "Knuckle height", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter(sJUSTIFICATION1, "Justif #1", GCCodeList, 0, 0, 0, 0, 2)
    If True Then
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION1, "Left", 1)
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION1, "Centered", 2)
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION1, "Right", 3)
        Call pGeometricConstructionDefinition.AddParameter(sOFFSET1, "U-offset #1", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0#)
    End If
    
    Call pGeometricConstructionDefinition.AddParameter(sJUSTIFICATION2, "Justif #2", GCCodeList, 0, 0, 0, 0, 2)
    If True Then
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION2, "Left", 1)
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION2, "Centered", 2)
        Call pGeometricConstructionDefinition.AddParameterValue(sJUSTIFICATION2, "Right", 3)
        Call pGeometricConstructionDefinition.AddParameter(sOFFSET2, "U-offset #2", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0#)
    End If
    
    Call pGeometricConstructionDefinition.AddErrorValue(lWIDTH_TOO_SMALL, "ErrorValue2", "Width too small")
    Call pGeometricConstructionDefinition.AddErrorValue(lKNUCKLE_HEIGHT_TOO_SMALL, "ErrorValue3", "Knuckle height too small")
    Call pGeometricConstructionDefinition.AddErrorValue(lKNUCKLE_HEIGHT_TOO_BIG, "ErrorValue4", "Knuckle height too big")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENER_AXIS1_NOT_INTERSECTING, "ErrorValue5", "The first stiffener does not intersect the plate")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENER_AXIS2_NOT_INTERSECTING, "ErrorValue6", "The second stiffener does not intersect the plate")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENER_AXES_TOO_CLOSE, "ErrorValue7", "The axes of the stiffener intersect the plate too closely")
    Call pGeometricConstructionDefinition.AddErrorValue(lOFFSET1_BIGGER_THAN_FLANGE_WIDTH, "ErrorValue15", "Offset1 bigger than flange width")
    Call pGeometricConstructionDefinition.AddErrorValue(lOFFSET2_BIGGER_THAN_FLANGE_WIDTH, "ErrorValue16", "Offset2 bigger than flange width")
    Call pGeometricConstructionDefinition.AddErrorValue(lOFFSET1_BIGGER_THAN_HALF_FLANGE_WIDTH, "ErrorValue17", "Offset1 bigger than half flange width")
    Call pGeometricConstructionDefinition.AddErrorValue(lOFFSET2_BIGGER_THAN_HALF_FLANGE_WIDTH, "ErrorValue18", "Offset2 bigger than half flange width")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_EVALUATE_KNUCKLED_PLANE, "ErrorValue19", "Cannot evaluate knuckled plane (see errors.log for more details)")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_CREATE_RIBBON_SURFACE, "ErrorValue23", "Cannot create ribbon surface")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_CREATE_CS_AT_POINT_ON_STIFFENER, "ErrorValue24", "Cannot create CS at point along stiffener")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENERS_IDENTICAL, "ErrorValue26", "The 2 stiffeners are identical")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_GENERATE_SURFACE_FROM_PLATE, "ErrorValue27", "Cannot generate a surface from the supporting plate system")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM1, "ErrorValue28", "Cannot generate a surface from the first stiffener system")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM2, "ErrorValue29", "Cannot generate a surface from the second stiffener system")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER1, "ErrorValue30", "Cannot intersect surfaces from plate and first stiffener")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER2, "ErrorValue31", "Cannot intersect surfaces from plate and second stiffener")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_POSITION_POINT_ALONG_STIFFENER1_AT_WIDTH, "ErrorValue32", "Cannot position point along first stiffener at distance = width (width too big)")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_POSITION_POINT_ALONG_STIFFENER2_AT_WIDTH, "ErrorValue33", "Cannot position point along second stiffener at distance = width (width too big)")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENERS_ON_SAME_UNDERLYING_PLATE_SYSTEM, "ErrorValue34", "Stiffeners on the same underlying plate system")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENER1_ON_SUPPORTING_PLATE_SYSTEM, "ErrorValue35", "First stiffener on the supporting plate")
    Call pGeometricConstructionDefinition.AddErrorValue(lSTIFFENER2_ON_SUPPORTING_PLATE_SYSTEM, "ErrorValue36", "Second stiffener on the supporting plate")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_CREATE_TUBE_FROM_PROFILE1, "ErrorValue37", "Cannot create tube geometry from first stiffener")
    Call pGeometricConstructionDefinition.AddErrorValue(lCANNOT_CREATE_TUBE_FROM_PROFILE2, "ErrorValue38", "Cannot create tube geometry from second stiffener")
    Call pGeometricConstructionDefinition.AddErrorValue(lINVALID_ORIENTATION_AXIS, "ErrorValue39", "Invalid Orientation axis (angle with plate normal should be < 60 degree)")

    
    Call pGeometricConstructionDefinition.AddErrorValue(lUNEXPECTED_INTERNAL_ERROR, "ErrorValue100", "Unexpected error in KnuckledBracket")
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
End Sub
Public Function RejectEverything(oObject As Object) As Integer
    RejectEverything = 0
End Function
Private Function IsInputER(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal sNameOfInput As String) As Boolean
    ' the IJERsystem interface cannot be used from VB6 so EdgeReinforcementSystem class will be used instead
    If TypeOf pGeometricConstruction.Input(sNameOfInput) Is EdgeReinforcementSystem Then
        IsInputER = True
    Else
        IsInputER = False
    End If
End Function
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    If pGeometricConstruction.Parameter(sPREVIEW) = 0 Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim bIsInMirror As Boolean
    bIsInMirror = GeometricConstructionMacro_IsBeingMirrored(pGeometricConstruction)
        
    ' check errors
    If pGeometricConstruction.Input(sSTIFFENER_SYSTEM1) Is pGeometricConstruction.Input(sSTIFFENER_SYSTEM2) Then Err.Raise lSTIFFENERS_IDENTICAL
    If pGeometricConstruction.Parameter(sWIDTH) < TOLERANCE Then Err.Raise lWIDTH_TOO_SMALL
    'Partial knuckled should have a valid knuckled height
    If pGeometricConstruction.Parameter(sKNUCKLE_TYPE) = 0 And pGeometricConstruction.Parameter(sKNUCKLE_HEIGHT) < TOLERANCE Then Err.Raise lKNUCKLE_HEIGHT_TOO_SMALL
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    ' determine if input stiffeners are edge reinforcements or regular stiffeners.
    ' it will be used for the evaluation of 024-CSByLines and 025-CSByLines
    Dim bIsStiffener1_ER As Boolean: bIsStiffener1_ER = IsInputER(pGeometricConstruction, sSTIFFENER_SYSTEM1)
    Dim bIsStiffener2_ER As Boolean: bIsStiffener2_ER = IsInputER(pGeometricConstruction, sSTIFFENER_SYSTEM2)
    
    Dim pGCGeomFactory As IJGCGeomFactory: Set pGCGeomFactory = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    Dim pGCGeomFactory2 As IJGCGeomFactory2: Set pGCGeomFactory2 = pGCGeomFactory
    Dim pGCGeomDotNetFactory As IJGCGeomDotNetFactory: Set pGCGeomDotNetFactory = CreateObject("GCCMNSTRDefinitions.GCGeomDotNetFactory")
    
    ' retrieve underlying root systems
    Dim oStiffenerSystem1 As Object: Set oStiffenerSystem1 = pGeometricConstruction.Input(sSTIFFENER_SYSTEM1): Call ShowName(oStiffenerSystem1)
    Dim oStiffenerSystem2 As Object: Set oStiffenerSystem2 = pGeometricConstruction.Input(sSTIFFENER_SYSTEM2): Call ShowName(oStiffenerSystem2)
    Dim oPlateSystem As Object: Set oPlateSystem = pGeometricConstruction.Input(sPLATE_SYSTEM): Call ShowName(oPlateSystem)
    Dim oPlateSystem1 As Object: Set oPlateSystem1 = GetDesignParent(oStiffenerSystem1): Call ShowName(oPlateSystem1)
    Dim oPlateSystem2 As Object: Set oPlateSystem2 = GetDesignParent(oStiffenerSystem2): Call ShowName(oPlateSystem2)
    If oPlateSystem1 Is oPlateSystem2 Then Err.Raise lSTIFFENER1_ON_SUPPORTING_PLATE_SYSTEM
    If oPlateSystem1 Is oPlateSystem Then Err.Raise lSTIFFENER1_ON_SUPPORTING_PLATE_SYSTEM
    If oPlateSystem2 Is oPlateSystem Then Err.Raise lSTIFFENER2_ON_SUPPORTING_PLATE_SYSTEM

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = 0.2
    oParamDistConstant1.Evaluate

    Dim oAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oAxisPortExtractor2 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "002-AxisPortExtractor")
    oAxisPortExtractor2.Inputs("Connectable").Add pGeometricConstruction.Inputs(sSTIFFENER_SYSTEM1).Item(1)
    oAxisPortExtractor2.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort1").Clear
    pGeometricConstruction.ControlledInputs("AxisPort1").AddElements oAxisPortExtractor2.ControlledInputs("Port")

    Dim oAxisPortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
    Set oAxisPortExtractor3 = oGCFactory.CreateEntity("AxisPortExtractor", pPOM, "003-AxisPortExtractor")
    oAxisPortExtractor3.Inputs("Connectable").Add pGeometricConstruction.Inputs(sSTIFFENER_SYSTEM2).Item(1)
    oAxisPortExtractor3.Evaluate
    pGeometricConstruction.ControlledInputs("AxisPort2").Clear
    pGeometricConstruction.ControlledInputs("AxisPort2").AddElements oAxisPortExtractor3.ControlledInputs("Port")

    Dim oFacePortExtractor04 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor04 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "004-FacePortExtractor0")
    oFacePortExtractor04.Inputs("Connectable").Add oPlateSystem
    oFacePortExtractor04.Parameter("Offset") = 0#
    oFacePortExtractor04.Evaluate
    pGeometricConstruction.ControlledInputs("FacePort").Clear
    pGeometricConstruction.ControlledInputs("FacePort").AddElements oFacePortExtractor04.ControlledInputs("Port")

    Dim oFacePortExtractor05 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor05 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "005-FacePortExtractor0")
    oFacePortExtractor05.Inputs("Connectable").Add oPlateSystem1
    oFacePortExtractor05.Parameter("Offset") = 0#
    oFacePortExtractor05.Evaluate
'    pGeometricConstruction.ControlledInputs("LowerDeckFacePort0").Clear
'    pGeometricConstruction.ControlledInputs("LowerDeckFacePort0").AddElements oFacePortExtractor05.ControlledInputs("Port")

    Dim oFacePortExtractor06 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor06 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "006-FacePortExtractor0")
    oFacePortExtractor06.Inputs("Connectable").Add oPlateSystem2
    oFacePortExtractor06.Parameter("Offset") = 0#
    oFacePortExtractor06.Evaluate
'    pGeometricConstruction.ControlledInputs("UppderDeckFacePort0").Clear
'    pGeometricConstruction.ControlledInputs("UppderDeckFacePort0").AddElements oFacePortExtractor06.ControlledInputs("Port")

    Dim oParamDistConstant7 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant7 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "007-ParamDistConstant")
    oParamDistConstant7.Parameter("Value") = pGeometricConstruction.Parameter(sWIDTH)
    oParamDistConstant7.Evaluate

    Dim oPointByCurveAndSurface8 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface8 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "008-PointByCurveAndSurface")
    oPointByCurveAndSurface8.Inputs("Curve1").Add oAxisPortExtractor2
    oPointByCurveAndSurface8.Inputs("Surface2").Add oFacePortExtractor04
    oPointByCurveAndSurface8.Parameter("TrackFlag") = GCNear
    Call GCEvaluate(oPointByCurveAndSurface8, lSTIFFENER_AXIS1_NOT_INTERSECTING)

    Dim oPointByCurveAndSurface9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface9 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, "009-PointByCurveAndSurface")
    oPointByCurveAndSurface9.Inputs("Curve1").Add oAxisPortExtractor3
    oPointByCurveAndSurface9.Inputs("Surface2").Add oFacePortExtractor04
    oPointByCurveAndSurface9.Parameter("TrackFlag") = GCNear
    Call GCEvaluate(oPointByCurveAndSurface9, lSTIFFENER_AXIS2_NOT_INTERSECTING)
    
    Dim oParamDistMeasureParallel10 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel10 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "010-ParamDistMeasureParallel")
    oParamDistMeasureParallel10.Inputs("Graphics").Add oPointByCurveAndSurface8
    oParamDistMeasureParallel10.Inputs("Graphics").Add oPointByCurveAndSurface9
    oParamDistMeasureParallel10.Parameter("Value") = 5.015529
    oParamDistMeasureParallel10.Evaluate
    
    If oParamDistMeasureParallel10.Parameter("Value") < EPSILON Then Err.Raise lSTIFFENER_AXES_TOO_CLOSE

    Dim oLineByPoints11 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints11 = oGCFactory.CreateEntity("LineByPoints", pPOM, "011-LineByPoints")
    oLineByPoints11.Inputs("StartPoint").Add oPointByCurveAndSurface8
    oLineByPoints11.Inputs("EndPoint").Add oPointByCurveAndSurface9
    oLineByPoints11.Evaluate
    
    Dim oPointAtCurveMiddle12 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle12 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "012-PointAtCurveMiddle")
    oPointAtCurveMiddle12.Inputs("Curve").Add oLineByPoints11
    oPointAtCurveMiddle12.Evaluate
    
    Dim oPointByProjectOnSurf13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByProjectOnSurf13 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "013-PointByProjectOnSurf")
    oPointByProjectOnSurf13.Inputs("Point").Add oPointAtCurveMiddle12
    oPointByProjectOnSurf13.Inputs("Surface").Add oFacePortExtractor04
    oPointByProjectOnSurf13.Evaluate
    
    ' axis z of oCSByLines15 (normal to Plate at midpoint of the line created by the 2 stiffeners)
    Dim oVectorNormalToSurface13a As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface13a = GCEvaluate(pGCGeomDotNetFactory.VectorNormalToSurface.Create(pPOM, _
                                                                                                 oFacePortExtractor04, _
                                                                                                 oPointByProjectOnSurf13, _
                                                                                                 Nothing, _
                                                                                                 1, GCNormal, GCNear), , "013a-VectorNormalToSurface")
    'test normal to plate to be in the direction of Orientation param
    Dim oNormalToSurface As IJDVector
    Set oNormalToSurface = Vector_FromLine(oVectorNormalToSurface13a)
    Dim pLocalCoordinateSystem As IJLocalCoordinateSystem: Set pLocalCoordinateSystem = pGeometricConstruction.Inputs(sCOORDINATE_SYSTEM).Item(1)
     
    Dim pVectorOfOrientation As IJDVector
    Select Case pGeometricConstruction.Parameter(sORIENTATION)
        Case 1: Set pVectorOfOrientation = pLocalCoordinateSystem.XAxis
        Case 2: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.XAxis, -1)
        Case 3: Set pVectorOfOrientation = pLocalCoordinateSystem.YAxis
        Case 4: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.YAxis, -1)
        Case 5: Set pVectorOfOrientation = pLocalCoordinateSystem.ZAxis
        Case 6: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.ZAxis, -1)
    End Select
    
    'orientation axis should be no more than 60 degrees from oNormalToSurface
    If Abs(oNormalToSurface.Dot(pVectorOfOrientation)) - 0.5 < -EPSILON Then
        Err.Raise lINVALID_ORIENTATION_AXIS
    End If
    
    ' axis y of oCSByLines15 (in the sense of stiffener1 toward stiffener2)
    Dim oLineByPoints14 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints14 = pGCGeomFactory2.LineByPoints(pPOM, _
                                                       oPointByProjectOnSurf13, _
                                                       oPointByCurveAndSurface9)
    If Not pPOM Is Nothing Then Call GeometricConstruction_NameByItem(oLineByPoints14, "014-LineByPoints")
                                                                        
    Dim oCSByLines15 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines15 = pGCGeomFactory2.CSByLines(pPOM, _
                                                 oLineByPoints14, _
                                                 oVectorNormalToSurface13a, _
                                                 Nothing, _
                                                 GCYZ, GCDirect)
    If Not pPOM Is Nothing Then Call GeometricConstruction_NameByItem(oCSByLines15, "015-CSByLines")

    ' axis x of oCSByLines15
    Dim oLineFromCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS16 = oGCFactory.CreateEntity("LineFromCS", pPOM, "016-LineFromCS")
    oLineFromCS16.Inputs("CoordinateSystem").Add oCSByLines15
    oLineFromCS16.Parameter("LookingAxis") = GCXAxis
    oLineFromCS16.Parameter("Length") = 100#
    oLineFromCS16.Parameter("CSOrientation") = GCDirect
    oLineFromCS16.Parameter("LineJustification") = 1
    oLineFromCS16.Evaluate
    
    Dim oCSByLines17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines17 = oGCFactory.CreateEntity("CSByLines", pPOM, "017-CSByLines")
    oCSByLines17.Inputs("AxisLine1").Add oLineFromCS16
    oCSByLines17.Inputs("AxisLine2").Add oLineByPoints14
    oCSByLines17.Parameter("AxesRoles") = GCXAxis
    oCSByLines17.Parameter("CSOrientation") = GCDirect
    oCSByLines17.Evaluate

    Dim oLineFromCS18 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS18 = oGCFactory.CreateEntity("LineFromCS", pPOM, "018-LineFromCS")
    oLineFromCS18.Inputs("CoordinateSystem").Add oCSByLines17
    oLineFromCS18.Parameter("LookingAxis") = GCZAxis
    oLineFromCS18.Parameter("Length") = 100#
    oLineFromCS18.Parameter("CSOrientation") = GCDirect
    oLineFromCS18.Parameter("LineJustification") = 1
    oLineFromCS18.Evaluate

    ' re-oriente oLineFromCS18 if necessary
    If True Then
        If Vector_FromLine(oLineFromCS18).Dot(pVectorOfOrientation) < 0 Then
            oCSByLines17.Parameter("CSOrientation") = GCIndirect
            oCSByLines17.Evaluate
            oLineFromCS18.Evaluate
        End If
    End If
    
    Dim oPointAtCurveEnd19 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd19 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "019-PointAtCurveEnd")
    oPointAtCurveEnd19.Inputs("Curve").Add oLineFromCS18
    oPointAtCurveEnd19.Evaluate
    
    'normal to parent of stiffener1 oriented between the 2 stiffeners
    Dim oVectorNormalToSurface20 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface20 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "020-VectorNormalToSurface")
    oVectorNormalToSurface20.Inputs("Surface").Add oFacePortExtractor05
    oVectorNormalToSurface20.Inputs("Point").Add oPointByCurveAndSurface8
    oVectorNormalToSurface20.Inputs("TrackPoint").Add oPointAtCurveMiddle12
    oVectorNormalToSurface20.Parameter("Range") = 1#
    oVectorNormalToSurface20.Parameter("Orientation") = GCNormal
    oVectorNormalToSurface20.Parameter("TrackFlag") = GCNear
    oVectorNormalToSurface20.Evaluate

    'normal to parent of stiffener2 oriented between the 2 stiffeners
    Dim oVectorNormalToSurface21 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorNormalToSurface21 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "021-VectorNormalToSurface")
    oVectorNormalToSurface21.Inputs("Surface").Add oFacePortExtractor06
    oVectorNormalToSurface21.Inputs("Point").Add oPointByCurveAndSurface9
    oVectorNormalToSurface21.Inputs("TrackPoint").Add oPointAtCurveMiddle12
    oVectorNormalToSurface21.Parameter("Range") = 1#
    oVectorNormalToSurface21.Parameter("Orientation") = GCNormal
    oVectorNormalToSurface21.Parameter("TrackFlag") = GCNear
    oVectorNormalToSurface21.Evaluate
    
    'tangent to stiffener axis oriented in the sense of oLineFromCS18 (orientation parameter)
    Dim oVectorTangentToCurve22 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve22 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "022-VectorTangentToCurve")
    oVectorTangentToCurve22.Inputs("Curve").Add oAxisPortExtractor2
    oVectorTangentToCurve22.Inputs("Point").Add oPointByCurveAndSurface8
    oVectorTangentToCurve22.Inputs("TrackPoint").Add oPointAtCurveEnd19
    oVectorTangentToCurve22.Parameter("Range") = 1#
    oVectorTangentToCurve22.Parameter("Orientation") = GCNormal
    oVectorTangentToCurve22.Parameter("TrackFlag") = GCNear
    oVectorTangentToCurve22.Evaluate

    Dim oVectorTangentToCurve23 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve23 = oGCFactory.CreateEntity("VectorTangentToCurve", pPOM, "023-VectorTangentToCurve")
    oVectorTangentToCurve23.Inputs("Curve").Add oAxisPortExtractor3
    oVectorTangentToCurve23.Inputs("Point").Add oPointByCurveAndSurface9
    oVectorTangentToCurve23.Inputs("TrackPoint").Add oPointAtCurveEnd19
    oVectorTangentToCurve23.Parameter("Range") = 1#
    oVectorTangentToCurve23.Parameter("Orientation") = GCNormal
    oVectorTangentToCurve23.Parameter("TrackFlag") = GCNear
    oVectorTangentToCurve23.Evaluate
    
    ' the evaluation differs between edge reinforcement system and stiffener system
    ' TR-CP-274540  Unable to control justify on knuckle bracket when support is Edge reinforcement
    Dim oCSByLines24 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines24 = oGCFactory.CreateEntity("CSByLines", pPOM, "024-CSByLines")
    If bIsStiffener1_ER Then
        ' be sure to have the Z direction pointing toward the midpoint
        Dim oCSByLines24a As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines24a = oGCFactory.CreateEntity("CSByLines", pPOM, "024a-CSByLines")
        oCSByLines24a.Inputs("AxisLine1").Add oVectorNormalToSurface20
        oCSByLines24a.Inputs("AxisLine2").Add oVectorTangentToCurve22
        oCSByLines24a.Inputs("TrackPoint").Add oPointAtCurveMiddle12
        oCSByLines24a.Parameter("AxesRoles") = GCYX
        oCSByLines24a.Parameter("CSOrientation") = GCDirect
        oCSByLines24a.Parameter("TrackFlag") = GCNear
        oCSByLines24a.Evaluate
        
        ' get the Z axis now we know that it is correctly oriented
        Dim oLineFromCS24b As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS24b = oGCFactory.CreateEntity("LineFromCS", pPOM, "024b-LineFromCS")
        oLineFromCS24b.Inputs("CoordinateSystem").Add oCSByLines24a
        oLineFromCS24b.Parameter("LookingAxis") = GCZAxis
        oLineFromCS24b.Parameter("Length") = 100#
        oLineFromCS24b.Parameter("CSOrientation") = GCDirect
        oLineFromCS24b.Parameter("LineJustification") = 1
        oLineFromCS24b.Evaluate
        
        ' use this Z axis instead of the normal of the supporting plate (or an ER)
        oCSByLines24.Inputs("AxisLine1").Add oLineFromCS24b
    Else
        oCSByLines24.Inputs("AxisLine1").Add oVectorNormalToSurface20
    End If
    oCSByLines24.Inputs("AxisLine2").Add oVectorTangentToCurve22
    oCSByLines24.Parameter("AxesRoles") = GCZX
    oCSByLines24.Parameter("CSOrientation") = GCDirect
    oCSByLines24.Parameter("TrackFlag") = GCNear
    oCSByLines24.Evaluate

    ' same comments related to ER specific case than in oCSByLines24
    Dim oCSByLines25 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByLines25 = oGCFactory.CreateEntity("CSByLines", pPOM, "025-CSByLines")
    If bIsStiffener2_ER Then
        Dim oCSByLines25a As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByLines25a = oGCFactory.CreateEntity("CSByLines", pPOM, "025a-CSByLines")
        oCSByLines25a.Inputs("AxisLine1").Add oVectorNormalToSurface21
        oCSByLines25a.Inputs("AxisLine2").Add oVectorTangentToCurve23
        oCSByLines25a.Inputs("TrackPoint").Add oPointAtCurveMiddle12
        oCSByLines25a.Parameter("AxesRoles") = GCYX
        oCSByLines25a.Parameter("CSOrientation") = GCIndirect
        oCSByLines25a.Parameter("TrackFlag") = GCNear
        oCSByLines25a.Evaluate
        
        Dim oLineFromCS25b As SP3DGeometricConstruction.GeometricConstruction
        Set oLineFromCS25b = oGCFactory.CreateEntity("LineFromCS", pPOM, "025b-LineFromCS")
        oLineFromCS25b.Inputs("CoordinateSystem").Add oCSByLines25a
        oLineFromCS25b.Parameter("LookingAxis") = GCZAxis
        oLineFromCS25b.Parameter("Length") = 100#
        oLineFromCS25b.Parameter("CSOrientation") = GCDirect
        oLineFromCS25b.Parameter("LineJustification") = 1
        oLineFromCS25b.Evaluate
        
        oCSByLines25.Inputs("AxisLine1").Add oLineFromCS25b
    Else
        oCSByLines25.Inputs("AxisLine1").Add oVectorNormalToSurface21
    End If
    oCSByLines25.Inputs("AxisLine2").Add oVectorTangentToCurve23
    oCSByLines25.Parameter("AxesRoles") = GCZX
    oCSByLines25.Parameter("CSOrientation") = GCIndirect
    oCSByLines25.Parameter("TrackFlag") = GCNear
    oCSByLines25.Evaluate
            
    'Create FacePortGenerator
    Dim oFacePortGenerator26 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortGenerator26 = oGCFactory.CreateEntity("FacePortGenerator", pPOM, "026-FacePortGenerator")
    oFacePortGenerator26.Inputs("RootPlateSystem").Add oPlateSystem
    oFacePortGenerator26.Inputs("ReferencePoint").Add oPointAtCurveEnd19
    Call GCEvaluate(oFacePortGenerator26, lCANNOT_GENERATE_SURFACE_FROM_PLATE)
    pGeometricConstruction.ControlledInputs("SupportLeafPlateSystem").Clear
    pGeometricConstruction.ControlledInputs("SupportLeafPlateSystem").AddElements oFacePortGenerator26.ControlledInputs("LeafPlateSystem")
  
    'Create TubeFromDetailedProfile for first stiffener
    Dim oTubeFromDetailedProfile27a As SP3DGeometricConstruction.GeometricConstruction
    Set oTubeFromDetailedProfile27a = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "027a-TubeFromDetailedProfile")
    oTubeFromDetailedProfile27a.Inputs("RootProfileSystem").Add oStiffenerSystem1
    oTubeFromDetailedProfile27a.Inputs("ReferencePoint").Add oPointByCurveAndSurface8
    oTubeFromDetailedProfile27a.Parameter("Length") = 0#
    Call GCEvaluate(oTubeFromDetailedProfile27a, lCANNOT_CREATE_TUBE_FROM_PROFILE1)
    pGeometricConstruction.ControlledInputs("RootPlateSystem1").Clear
    pGeometricConstruction.ControlledInputs("RootPlateSystem1").AddElements oTubeFromDetailedProfile27a.ControlledInputs("RootPlateSystem")
    pGeometricConstruction.ControlledInputs("LeafPlateSystem1").Clear
    pGeometricConstruction.ControlledInputs("LeafPlateSystem1").AddElements oTubeFromDetailedProfile27a.ControlledInputs("LeafPlateSystem")
    
    'Create TubeFromDetailedProfile for second stiffener
    Dim oTubeFromDetailedProfile27b As SP3DGeometricConstruction.GeometricConstruction
    Set oTubeFromDetailedProfile27b = oGCFactory.CreateEntity("TubeFromDetailedProfile", pPOM, "027b-TubeFromDetailedProfile")
    oTubeFromDetailedProfile27b.Inputs("RootProfileSystem").Add oStiffenerSystem2
    oTubeFromDetailedProfile27b.Inputs("ReferencePoint").Add oPointByCurveAndSurface9
    oTubeFromDetailedProfile27b.Parameter("Length") = 0#
    Call GCEvaluate(oTubeFromDetailedProfile27b, lCANNOT_CREATE_TUBE_FROM_PROFILE2)
    pGeometricConstruction.ControlledInputs("RootPlateSystem2").Clear
    pGeometricConstruction.ControlledInputs("RootPlateSystem2").AddElements oTubeFromDetailedProfile27b.ControlledInputs("RootPlateSystem")
    pGeometricConstruction.ControlledInputs("LeafPlateSystem2").Clear
    pGeometricConstruction.ControlledInputs("LeafPlateSystem2").AddElements oTubeFromDetailedProfile27b.ControlledInputs("LeafPlateSystem")
    
    'Create SurfaceBodyExtractor for first stiffener
    Dim oSurfaceBodyExtractor28a  As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfaceBodyExtractor28a = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "028a-SurfaceBodyExtractor")
    oSurfaceBodyExtractor28a.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile27a
    oSurfaceBodyExtractor28a.Inputs("CoordinateSystem").Add oCSByLines24
    oSurfaceBodyExtractor28a.Parameter("FacesContext") = 2
    oSurfaceBodyExtractor28a.Parameter("LookingAxis") = GCZAxis
    oSurfaceBodyExtractor28a.Parameter("IntersectingPlane") = 2
    oSurfaceBodyExtractor28a.Parameter("SurfaceType") = 1
    oSurfaceBodyExtractor28a.Parameter("TrackFlag") = GCFar
    oSurfaceBodyExtractor28a.Parameter("Offset") = 0#
    Call GCEvaluate(oSurfaceBodyExtractor28a, lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM1)
    
    'Create SurfaceBodyExtractor for second stiffener
    Dim oSurfaceBodyExtractor28b As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfaceBodyExtractor28b = oGCFactory.CreateEntity("SurfaceBodyExtractor", pPOM, "028b-SurfaceBodyExtractor")
    oSurfaceBodyExtractor28b.Inputs("SurfaceOrSolidBody").Add oTubeFromDetailedProfile27b
    oSurfaceBodyExtractor28b.Inputs("CoordinateSystem").Add oCSByLines25
    oSurfaceBodyExtractor28b.Parameter("FacesContext") = 2
    oSurfaceBodyExtractor28b.Parameter("LookingAxis") = GCZAxis
    oSurfaceBodyExtractor28b.Parameter("IntersectingPlane") = 2
    oSurfaceBodyExtractor28b.Parameter("SurfaceType") = 1
    oSurfaceBodyExtractor28b.Parameter("TrackFlag") = GCFar
    oSurfaceBodyExtractor28b.Parameter("Offset") = 0#
    Call GCEvaluate(oSurfaceBodyExtractor28b, lCANNOT_GENERATE_SURFACE_FROM_STIFFENER_SYSTEM2)

    Dim oCurveByIntersection29 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection29 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "029-CurveByIntersection")
    oCurveByIntersection29.Inputs("Surface1").Add oFacePortGenerator26
    oCurveByIntersection29.Inputs("Surface2").Add oSurfaceBodyExtractor28a
    oCurveByIntersection29.Inputs("CoordinateSystem").Add oCSByLines24
    oCurveByIntersection29.Parameter("TrackFlag") = GCNear
    oCurveByIntersection29.Parameter("ConstructionSurface") = 0
    Call GCEvaluate(oCurveByIntersection29, lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER1)

    Dim oCurveByIntersection30 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection30 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "030-CurveByIntersection")
    oCurveByIntersection30.Inputs("Surface1").Add oFacePortGenerator26
    oCurveByIntersection30.Inputs("Surface2").Add oSurfaceBodyExtractor28b
    oCurveByIntersection30.Inputs("CoordinateSystem").Add oCSByLines25
    oCurveByIntersection30.Parameter("TrackFlag") = GCNear
    oCurveByIntersection30.Parameter("ConstructionSurface") = 0
    Call GCEvaluate(oCurveByIntersection30, lCANNOT_INTERSECT_SURFACES_FROM_PLATE_AND_STIFFENER2)

    Dim oPointAtCurveMiddle31 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle31 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "031-PointAtCurveMiddle")
    oPointAtCurveMiddle31.Inputs("Curve").Add oCurveByIntersection29
    oPointAtCurveMiddle31.Evaluate

    Dim oPointAtCurveMiddle32 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle32 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "032-PointAtCurveMiddle")
    oPointAtCurveMiddle32.Inputs("Curve").Add oCurveByIntersection30
    oPointAtCurveMiddle32.Evaluate

    Dim oSphereByCenterRadius33 As SP3DGeometricConstruction.GeometricConstruction
    Set oSphereByCenterRadius33 = oGCFactory.CreateEntity("SphereByCenterRadius", pPOM, "033-SphereByCenterRadius")
    oSphereByCenterRadius33.Inputs("Center").Add oPointAtCurveMiddle31
    oSphereByCenterRadius33.Parameter("Radius") = oParamDistConstant7
    oSphereByCenterRadius33.Evaluate

    Dim oSphereByCenterRadius34 As SP3DGeometricConstruction.GeometricConstruction
    Set oSphereByCenterRadius34 = oGCFactory.CreateEntity("SphereByCenterRadius", pPOM, "034-SphereByCenterRadius")
    oSphereByCenterRadius34.Inputs("Center").Add oPointAtCurveMiddle32
    oSphereByCenterRadius34.Parameter("Radius") = oParamDistConstant7
    oSphereByCenterRadius34.Evaluate

    Dim oCurveByIntersection35 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection35 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "035-CurveByIntersection")
    oCurveByIntersection35.Inputs("Surface1").Add oSphereByCenterRadius33
    oCurveByIntersection35.Inputs("Surface2").Add oSurfaceBodyExtractor28a
    oCurveByIntersection35.Inputs("TrackPoint").Add oPointAtCurveEnd19
    oCurveByIntersection35.Parameter("TrackFlag") = GCNear
    oCurveByIntersection35.Parameter("ConstructionSurface") = 0
    Call GCEvaluate(oCurveByIntersection35, lCANNOT_POSITION_POINT_ALONG_STIFFENER1_AT_WIDTH)

    Dim oCurveByIntersection36 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection36 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "036-CurveByIntersection")
    oCurveByIntersection36.Inputs("Surface1").Add oSphereByCenterRadius34
    oCurveByIntersection36.Inputs("Surface2").Add oSurfaceBodyExtractor28b
    oCurveByIntersection36.Inputs("TrackPoint").Add oPointAtCurveEnd19
    oCurveByIntersection36.Parameter("TrackFlag") = GCNear
    oCurveByIntersection36.Parameter("ConstructionSurface") = 0
    Call GCEvaluate(oCurveByIntersection36, lCANNOT_POSITION_POINT_ALONG_STIFFENER2_AT_WIDTH)

    Dim oPointAtCurveMiddle37 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle37 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "037-PointAtCurveMiddle")
    oPointAtCurveMiddle37.Inputs("Curve").Add oCurveByIntersection35
    oPointAtCurveMiddle37.Evaluate

    Dim oPointAtCurveMiddle38 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle38 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "038-PointAtCurveMiddle")
    oPointAtCurveMiddle38.Inputs("Curve").Add oCurveByIntersection36
    oPointAtCurveMiddle38.Evaluate
    
    'Check that the 2 points are on the same side of the ref surface (can have only one intersection with sphere that can be on the wrong side)
    'We should eliminate this case.
    Dim oLineByPoints37 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints37 = oGCFactory.CreateEntity("LineByPoints", pPOM, "037-LineByPoints")
    oLineByPoints37.Inputs("StartPoint").Add oPointAtCurveMiddle31
    oLineByPoints37.Inputs("EndPoint").Add oPointAtCurveMiddle37
    oLineByPoints37.Evaluate
    
    If Vector_FromLine(oLineByPoints37).Dot(Vector_FromLine(oLineFromCS18)) < 0 Then
        Err.Raise lCANNOT_POSITION_POINT_ALONG_STIFFENER1_AT_WIDTH
    End If
    
    Dim oLineByPoints38 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints38 = oGCFactory.CreateEntity("LineByPoints", pPOM, "038-LineByPoints")
    oLineByPoints38.Inputs("StartPoint").Add oPointAtCurveMiddle32
    oLineByPoints38.Inputs("EndPoint").Add oPointAtCurveMiddle38
    oLineByPoints38.Evaluate
    
    If Vector_FromLine(oLineByPoints38).Dot(Vector_FromLine(oLineFromCS18)) < 0 Then
        Err.Raise lCANNOT_POSITION_POINT_ALONG_STIFFENER2_AT_WIDTH
    End If
    
    ' generate 2 small planes to intersect with the FacePortExtractors
    Dim oParamDistMeasureLength39 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength39 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "039-ParamDistMeasureLength")
    oParamDistMeasureLength39.Inputs("Curve").Add oCurveByIntersection35
    oParamDistMeasureLength39.Evaluate
    
    Dim oParamDistMeasureLength40 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureLength40 = oGCFactory.CreateEntity("ParamDistMeasureLength", pPOM, "040-ParamDistMeasureLength")
    oParamDistMeasureLength40.Inputs("Curve").Add oCurveByIntersection36
    oParamDistMeasureLength40.Evaluate
    
    Dim oVectorTangentToCurve41 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve41 = GCEvaluate(pGCGeomDotNetFactory.VectorTangentToCurve.Create(pPOM, _
                                                                                              oAxisPortExtractor2, _
                                                                                              oPointAtCurveMiddle37, _
                                                                                              oPointAtCurveEnd19, _
                                                                                              Nothing, 10, GCNormal, GCNear), , "041-VectorTangentToCurve")
                                                                                 
    Dim oVectorTangentToCurve42 As SP3DGeometricConstruction.GeometricConstruction
    Set oVectorTangentToCurve42 = GCEvaluate(pGCGeomDotNetFactory.VectorTangentToCurve.Create(pPOM, _
                                                                                              oAxisPortExtractor3, _
                                                                                              oPointAtCurveMiddle38, _
                                                                                              oPointAtCurveEnd19, _
                                                                                              Nothing, 10, GCNormal, GCNear), , "042-VectorTangentToCurve")
                                                                                 
    Dim oPlaneByPointNormal43 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal43 = GCEvaluate(pGCGeomFactory.PlaneByPointNormal.Create(pPOM, _
                                                                                    oPointAtCurveMiddle37, _
                                                                                    oVectorTangentToCurve41, _
                                                                                    2 * oParamDistMeasureLength39.Parameter("Value")), , "043-PlaneByPointNormal")
   
    Dim oPlaneByPointNormal44 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneByPointNormal44 = GCEvaluate(pGCGeomFactory.PlaneByPointNormal.Create(pPOM, _
                                                                                    oPointAtCurveMiddle38, _
                                                                                    oVectorTangentToCurve42, _
                                                                                    2 * oParamDistMeasureLength40.Parameter("Value")), , "044-PlaneByPointNormal")
     
    Dim oCurveByIntersection45 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection45 = GCEvaluate(pGCGeomFactory.CurveByIntersection.Create(pPOM, _
                                                                                      oPlaneByPointNormal43, _
                                                                                      oSurfaceBodyExtractor28a, _
                                                                                      oCSByLines24, _
                                                                                      oPointAtCurveMiddle37, _
                                                                                      GCNear), , "045-CurveByIntersection")
   
    Dim oCurveByIntersection46 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection46 = GCEvaluate(pGCGeomFactory.CurveByIntersection.Create(pPOM, _
                                                                                      oPlaneByPointNormal44, _
                                                                                      oSurfaceBodyExtractor28b, _
                                                                                      oCSByLines25, _
                                                                                      oPointAtCurveMiddle38, _
                                                                                      GCNear), , "046-CurveByIntersection")
    
    Dim oReferencePoint47 As SP3DGeometricConstruction.GeometricConstruction
    Dim oReferencePoint48 As SP3DGeometricConstruction.GeometricConstruction
    If True Then
        Select Case pGeometricConstruction.Parameter(sJUSTIFICATION1)
            Case 1:
                Set oReferencePoint47 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "047-PointAtCurveStart")
                oReferencePoint47.Inputs("Curve").Add oCurveByIntersection29
                oReferencePoint47.Evaluate
                
                Set oReferencePoint48 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "048-PointAtCurveStart")
                oReferencePoint48.Inputs("Curve").Add oCurveByIntersection45
                oReferencePoint48.Evaluate
            Case 2:
                Set oReferencePoint47 = oPointAtCurveMiddle31
                Set oReferencePoint48 = oPointAtCurveMiddle37
            Case 3:
                Set oReferencePoint47 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "047-PointAtCurveEnd")
                oReferencePoint47.Inputs("Curve").Add oCurveByIntersection29
                oReferencePoint47.Evaluate
                
                Set oReferencePoint48 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "048-PointAtCurveEnd")
                oReferencePoint48.Inputs("Curve").Add oCurveByIntersection45
                oReferencePoint48.Evaluate
        End Select
    End If
    
    Dim oPointAlongCurve49 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointAlongCurve50 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter(sOFFSET1) <> 0 Then
        Select Case pGeometricConstruction.Parameter(sJUSTIFICATION1)
            Case 1:
                If Abs(pGeometricConstruction.Parameter(sOFFSET1)) > oParamDistMeasureLength39.Parameter("Value") Then Err.Raise lOFFSET1_BIGGER_THAN_FLANGE_WIDTH
                Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
                oPointAlongCurve49.Inputs("Curve").Add oCurveByIntersection29
                oPointAlongCurve49.Inputs("Point").Add oReferencePoint47
                oPointAlongCurve49.Parameter("Distance") = Abs(pGeometricConstruction.Parameter(sOFFSET1))
                oPointAlongCurve49.Evaluate

                Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
                oPointAlongCurve50.Inputs("Curve").Add oCurveByIntersection45
                oPointAlongCurve50.Inputs("Point").Add oReferencePoint48
                oPointAlongCurve50.Parameter("Distance") = Abs(pGeometricConstruction.Parameter(sOFFSET1))
                oPointAlongCurve50.Evaluate

            Case 2:
                If Abs(pGeometricConstruction.Parameter(sOFFSET1)) > oParamDistMeasureLength39.Parameter("Value") / 2 Then Err.Raise lOFFSET1_BIGGER_THAN_HALF_FLANGE_WIDTH
                Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
                oPointAlongCurve49.Inputs("Curve").Add oCurveByIntersection29
                oPointAlongCurve49.Inputs("Point").Add oReferencePoint47
                oPointAlongCurve49.Parameter("Distance") = pGeometricConstruction.Parameter(sOFFSET1)
                oPointAlongCurve49.Evaluate

                Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
                oPointAlongCurve50.Inputs("Curve").Add oCurveByIntersection45
                oPointAlongCurve50.Inputs("Point").Add oReferencePoint48
                oPointAlongCurve50.Parameter("Distance") = pGeometricConstruction.Parameter(sOFFSET1)
                oPointAlongCurve50.Evaluate
            
            Case 3:
                If Abs(pGeometricConstruction.Parameter(sOFFSET1)) > oParamDistMeasureLength39.Parameter("Value") Then Err.Raise lOFFSET1_BIGGER_THAN_FLANGE_WIDTH
                Set oPointAlongCurve49 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "049-PointAlongCurve")
                oPointAlongCurve49.Inputs("Curve").Add oCurveByIntersection29
                oPointAlongCurve49.Inputs("Point").Add oReferencePoint47
                oPointAlongCurve49.Parameter("Distance") = -Abs(pGeometricConstruction.Parameter(sOFFSET1))
                oPointAlongCurve49.Evaluate

                Set oPointAlongCurve50 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "050-PointAlongCurve")
                oPointAlongCurve50.Inputs("Curve").Add oCurveByIntersection45
                oPointAlongCurve50.Inputs("Point").Add oReferencePoint48
                oPointAlongCurve50.Parameter("Distance") = -Abs(pGeometricConstruction.Parameter(sOFFSET1))
                oPointAlongCurve50.Evaluate
        End Select
    Else
        Set oPointAlongCurve49 = oReferencePoint47
        Set oPointAlongCurve50 = oReferencePoint48
    End If
   
    Dim oReferencePoint51 As SP3DGeometricConstruction.GeometricConstruction
    Dim oReferencePoint52 As SP3DGeometricConstruction.GeometricConstruction
    If True Then
        Select Case pGeometricConstruction.Parameter(sJUSTIFICATION2)
            Case 1:
                Set oReferencePoint51 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "051-PointAtCurveStart")
                oReferencePoint51.Inputs("Curve").Add oCurveByIntersection30
                oReferencePoint51.Evaluate

                Set oReferencePoint52 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "052-PointAtCurveStart")
                oReferencePoint52.Inputs("Curve").Add oCurveByIntersection46
                oReferencePoint52.Evaluate
            Case 2:
                Set oReferencePoint51 = oPointAtCurveMiddle32
                Set oReferencePoint52 = oPointAtCurveMiddle38
            Case 3:
                Set oReferencePoint51 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "051-PointAtCurveEnd")
                oReferencePoint51.Inputs("Curve").Add oCurveByIntersection30
                oReferencePoint51.Evaluate

                Set oReferencePoint52 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "052-PointAtCurveEnd")
                oReferencePoint52.Inputs("Curve").Add oCurveByIntersection46
                oReferencePoint52.Evaluate
        End Select
    End If
    
    Dim oPointAlongCurve53 As SP3DGeometricConstruction.GeometricConstruction
    Dim oPointAlongCurve54 As SP3DGeometricConstruction.GeometricConstruction
    If pGeometricConstruction.Parameter(sOFFSET2) <> 0 Then
        Select Case pGeometricConstruction.Parameter(sJUSTIFICATION2)
            Case 1:
               If Abs(pGeometricConstruction.Parameter(sOFFSET2)) > oParamDistMeasureLength40.Parameter("Value") Then Err.Raise lOFFSET2_BIGGER_THAN_FLANGE_WIDTH
                Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "053-PointAlongCurve")
                oPointAlongCurve53.Inputs("Curve").Add oCurveByIntersection30
                oPointAlongCurve53.Inputs("Point").Add oReferencePoint51
                oPointAlongCurve53.Parameter("Distance") = Abs(pGeometricConstruction.Parameter(sOFFSET2))
                oPointAlongCurve53.Evaluate

                Set oPointAlongCurve54 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "054-PointAlongCurve")
                oPointAlongCurve54.Inputs("Curve").Add oCurveByIntersection46
                oPointAlongCurve54.Inputs("Point").Add oReferencePoint52
                oPointAlongCurve54.Parameter("Distance") = Abs(pGeometricConstruction.Parameter(sOFFSET2))
                oPointAlongCurve54.Evaluate
            
            Case 2:
               If Abs(pGeometricConstruction.Parameter(sOFFSET2)) > oParamDistMeasureLength40.Parameter("Value") / 2 Then Err.Raise lOFFSET2_BIGGER_THAN_HALF_FLANGE_WIDTH
                Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "053-PointAlongCurve")
                oPointAlongCurve53.Inputs("Curve").Add oCurveByIntersection30
                oPointAlongCurve53.Inputs("Point").Add oReferencePoint51
                oPointAlongCurve53.Parameter("Distance") = pGeometricConstruction.Parameter(sOFFSET2)
                oPointAlongCurve53.Evaluate

                Set oPointAlongCurve54 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "054-PointAlongCurve")
                oPointAlongCurve54.Inputs("Curve").Add oCurveByIntersection46
                oPointAlongCurve54.Inputs("Point").Add oReferencePoint52
                oPointAlongCurve54.Parameter("Distance") = pGeometricConstruction.Parameter(sOFFSET2)
                oPointAlongCurve54.Evaluate
            
            Case 3:
               If Abs(pGeometricConstruction.Parameter(sOFFSET2)) > oParamDistMeasureLength40.Parameter("Value") Then Err.Raise lOFFSET2_BIGGER_THAN_FLANGE_WIDTH
                Set oPointAlongCurve53 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "053-PointAlongCurve")
                oPointAlongCurve53.Inputs("Curve").Add oCurveByIntersection30
                oPointAlongCurve53.Inputs("Point").Add oReferencePoint51
                oPointAlongCurve53.Parameter("Distance") = -Abs(pGeometricConstruction.Parameter(sOFFSET2))
                oPointAlongCurve53.Evaluate

                Set oPointAlongCurve54 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "054-PointAlongCurve")
                oPointAlongCurve54.Inputs("Curve").Add oCurveByIntersection46
                oPointAlongCurve54.Inputs("Point").Add oReferencePoint52
                oPointAlongCurve54.Parameter("Distance") = -Abs(pGeometricConstruction.Parameter(sOFFSET2))
                oPointAlongCurve54.Evaluate
        End Select
    Else
        Set oPointAlongCurve53 = oReferencePoint51
        Set oPointAlongCurve54 = oReferencePoint52
    End If
    
    Dim dKnuckleHeight As Double
    Select Case pGeometricConstruction.Parameter(sKNUCKLE_TYPE)
        Case 0:
            'Partial knuckled should have a valid knuckled height
            dKnuckleHeight = pGeometricConstruction.Parameter(sKNUCKLE_HEIGHT)
            If dKnuckleHeight > Length_FromCurve(oLineByPoints11) Then Err.Raise lKNUCKLE_HEIGHT_TOO_BIG
        Case 1: dKnuckleHeight = 0
    End Select
    
    Dim oKnuckledPlane55 As SP3DGeometricConstruction.GeometricConstruction
    Set oKnuckledPlane55 = oGCFactory.CreateEntity("KnuckledPlane", pPOM, "055-KnuckledPlane")
    oKnuckledPlane55.Inputs("OppositePoint").Add oPointAlongCurve54
    oKnuckledPlane55.Inputs("OriginPoint").Add oPointAlongCurve49
    oKnuckledPlane55.Inputs("YAxisPoint").Add oPointAlongCurve53
    oKnuckledPlane55.Inputs("XAxisPoint").Add oPointAlongCurve50
    oKnuckledPlane55.Parameter("Distance") = dKnuckleHeight
    oKnuckledPlane55.Parameter("Extension") = oParamDistConstant1
    oKnuckledPlane55.Evaluate

    Dim oLineByPoints56 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints56 = oGCFactory.CreateEntity("LineByPoints", pPOM, "056-LineByPoints")
    oLineByPoints56.Inputs("StartPoint").Add oKnuckledPlane55.Output("Knuckle", 1)
    oLineByPoints56.Inputs("EndPoint").Add oPointAlongCurve50
    oLineByPoints56.Evaluate

    Dim oPointAtCurveMiddle56 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle56 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "056-PointAtCurveMiddle")
    oPointAtCurveMiddle56.Inputs("Curve").Add oLineByPoints56
    oPointAtCurveMiddle56.Evaluate
    
    If pGeometricConstruction.Parameter(sKNUCKLE_TYPE) = 0 Then
        'Partial knuckled
        Dim oLineByPoints57 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints57 = oGCFactory.CreateEntity("LineByPoints", pPOM, "057-LineByPoints")
        oLineByPoints57.Inputs("StartPoint").Add oKnuckledPlane55.Output("Knuckle", 1)
        oLineByPoints57.Inputs("EndPoint").Add oPointAlongCurve54
        oLineByPoints57.Evaluate
    End If
    
    Dim oPointAlongCurve58 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve58 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "058-PointAlongCurve")
    oPointAlongCurve58.Inputs("Curve").Add oLineByPoints56
    oPointAlongCurve58.Inputs("Point").Add oPointAlongCurve50
    oPointAlongCurve58.Inputs("TrackPoint").Add oPointAtCurveMiddle56
    oPointAlongCurve58.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve58.Parameter("TrackFlag") = GCFar
    oPointAlongCurve58.Evaluate

    Dim oPointAlongCurve59 As SP3DGeometricConstruction.GeometricConstruction
    Dim oRibbonCurve60 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve59 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "059-PointAlongCurve")
    If pGeometricConstruction.Parameter(sKNUCKLE_TYPE) = 0 Then
        'Partial knuckled, create a linestring from the 2 extended lines as ribbon curve
        oPointAlongCurve59.Inputs("Curve").Add oLineByPoints57
        oPointAlongCurve59.Inputs("TrackPoint").Add oKnuckledPlane55.Output("Knuckle", 1)
        oPointAlongCurve59.Inputs("Point").Add oPointAlongCurve54
        oPointAlongCurve59.Parameter("Distance") = oParamDistConstant1
        oPointAlongCurve59.Parameter("TrackFlag") = GCFar
        oPointAlongCurve59.Evaluate
        
        Set oRibbonCurve60 = oGCFactory.CreateEntity("LineStringByPoints", pPOM, "060-LineStringForRibbon")
        oRibbonCurve60.Inputs("Points").Add oPointAlongCurve58
        oRibbonCurve60.Inputs("Points").Add oKnuckledPlane55.Output("Knuckle", 1)
        oRibbonCurve60.Inputs("Points").Add oPointAlongCurve59
        oRibbonCurve60.Parameter("ClosedFlag") = 0
        oRibbonCurve60.Evaluate
    Else
        'For full knuckled type, knuckle point is oPointAlongCurve54, use oPointAlongCurve50 as track point
        'just extend oLineByPoints56 as ribbon curve
        oPointAlongCurve59.Inputs("Curve").Add oLineByPoints56
        oPointAlongCurve59.Inputs("TrackPoint").Add oPointAtCurveMiddle56
        oPointAlongCurve59.Inputs("Point").Add oPointAlongCurve54
        oPointAlongCurve59.Parameter("Distance") = oParamDistConstant1
        oPointAlongCurve59.Parameter("TrackFlag") = GCFar
        oPointAlongCurve59.Evaluate
        
        Set oRibbonCurve60 = oGCFactory.CreateEntity("LineByPoints", pPOM, "060-LineByPointsForRibbon")
        oRibbonCurve60.Inputs("StartPoint").Add oPointAlongCurve58
        oRibbonCurve60.Inputs("EndPoint").Add oPointAlongCurve59
        oRibbonCurve60.Evaluate
    End If


    Dim oSurfAsRibbon61 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfAsRibbon61 = oGCFactory.CreateEntity("SurfAsRibbon", pPOM, "061-SurfAsRibbon")
    oSurfAsRibbon61.Inputs("Curve").Add oRibbonCurve60
    oSurfAsRibbon61.Inputs("Surface").Add oKnuckledPlane55.Output("Surface", 1)
    oSurfAsRibbon61.Parameter("Width") = 0.25
    Call GCEvaluate(oSurfAsRibbon61, lCANNOT_CREATE_RIBBON_SURFACE)

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Boundary", "AxisPort1") = oSurfaceBodyExtractor28a.Output
    oGCMacro.Output("Boundary", "AxisPort2") = oSurfaceBodyExtractor28b.Output
    oGCMacro.Output("Boundary", "FacePort") = oFacePortGenerator26.Output
    oGCMacro.Output("Boundary", 4) = oSurfAsRibbon61.Output
    oGCMacro.Output("Support", 1) = oKnuckledPlane55.Output("Surface", 1)
    
    Exit Sub
ErrorHandler:
    pGeometricConstruction.RaiseError Err.Number
End Sub
Private Function GetDesignParent(oDesignChild As Object) As Object
    Dim oDesignParent As Object
    If True Then
        Dim pDesignChild As IJDesignChild: Set pDesignChild = oDesignChild
        Set oDesignParent = pDesignChild.GetParent
    End If
    
    ' return result
    Set GetDesignParent = oDesignParent
End Function
Private Sub ShowName(pNamedItem As IJNamedItem)
    Debug.Print pNamedItem.Name
End Sub
Function GCEvaluate(pGC As GeometricConstruction, Optional lError As Long = lUNEXPECTED_INTERNAL_ERROR, Optional sName As String = "") As Object
    On Error Resume Next
    pGC.Evaluate
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise lError
    End If
    
    Dim pPOM As IJDPOM
    If True Then
        Dim pObject As IJDObject: Set pObject = pGC
        On Error Resume Next
        Set pPOM = pObject.ResourceManager
    On Error GoTo 0
    End If
    
    If Not pPOM Is Nothing And Not sName = "" Then Call GeometricConstruction_NameByItem(pGC, sName)
    
    On Error GoTo 0
    Set GCEvaluate = pGC
End Function
Sub GeometricConstruction_NameByItem(ByVal pGeometricConstruction As IJGeometricConstruction, ByVal sName As String)
    ' disconnect naming rule if exists
    If Not GeometricConstruction_GetNamingRuleAE(pGeometricConstruction) Is Nothing Then Call Entity_DisconnectRelatedEntity(pGeometricConstruction, "IJNamedItem", "entityAE")

    ' put the provided name
    Dim pNamedItem As IJNamedItem: Set pNamedItem = pGeometricConstruction
    pNamedItem.Name = sName
End Sub
Function GeometricConstruction_GetNamingRuleAE(ByVal pGeometricConstruction As IJGeometricConstruction) As Object
    ' prepare result
    Set GeometricConstruction_GetNamingRuleAE = Nothing

    ' retrieve collention of naming rules
    Dim pElementsOfNamingRuleAE As IMSCoreCollections.IJElements: Set pElementsOfNamingRuleAE = Entity_GetElementsOfRelatedEntities(pGeometricConstruction, "IJNamedItem", "entityAE")

    ' return result
    If pElementsOfNamingRuleAE.Count = 1 Then
        Set GeometricConstruction_GetNamingRuleAE = pElementsOfNamingRuleAE(1)
    End If
End Function
Public Function Entity_GetElementsOfRelatedEntities(oSource As Object, sNameOfRelation As String, sNameOfRole As String) As IJElements
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfRelation, sNameOfRole)
    
    Dim pElementsOfRelatedEntities As IJElements
    Set pElementsOfRelatedEntities = New JObjectCollection
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.Count
        Dim pRelationship As IJDRelationship
        Set pRelationship = pRelationshipCol.Item(lIndex)
        
        Dim oTarget As Object
        Set oTarget = pRelationship.Target
    
        Call pElementsOfRelatedEntities.Add(pRelationship.Target)
    Next
    Set Entity_GetElementsOfRelatedEntities = pElementsOfRelatedEntities
End Function
Private Sub Entity_DisconnectRelatedEntity(ByVal oSource As Object, ByVal sNameOfInterface As String, ByVal sNameOfRole As String)
    Dim pAssocRelations As IMSRelation.IJDAssocRelation
    Set pAssocRelations = oSource

    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfInterface, sNameOfRole)

    Dim pRevision As IJRevision: Set pRevision = New JRevision
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.Count
        Dim pRelationship As IJDRelationship: Set pRelationship = pRelationshipCol.Item(lIndex)
        Call pRevision.RemoveRelationship(pRelationship)
    Next
End Sub

Private Function GetGlobalShowDetails() As Long
    Dim oToolshelper As New PORTHELPERLib.ToolsHelper
    GetGlobalShowDetails = oToolshelper.GetGlobalShowDetails()
End Function

Private Function Object_GetPOM(ByVal pObject As IngrGeom3D.IJDObject) As IJDPOM
    On Error Resume Next
    Object_GetPOM = pObject.ResourceManager
    On Error GoTo 0
End Function

'Find which axis vector (ortho vectors) is colinear to the given vector, if none is found return the nearest axis
'each axis can be Null to only compare 2 axes
Private Function Orientation_FromCSVectors(pVectorToOrient As IJDVector, pGCVectorOfAxis1 As IJDVector, pGCVectorOfAxis2 As IJDVector, pGCVectorOfAxis3 As IJDVector, ByRef eOrientation As eGlobalTrackFlag) As Boolean
    Dim dDot1 As Double
    Dim dDot2 As Double
    Dim dDot3 As Double
    Dim dDotMax As Double
    Dim bIsMaxSet As Boolean
    
    bIsMaxSet = False
    Orientation_FromCSVectors = False
    
    Set pVectorToOrient = Vector_Normalize(pVectorToOrient)
    Call DebugValue("Orientation_FromCSVectors vector to orient ", pVectorToOrient)
    Call DebugValue("Orientation_FromCSVectors pGCVectorOfAxis1 ", pGCVectorOfAxis1)
    Call DebugValue("Orientation_FromCSVectors pGCVectorOfAxis2 ", pGCVectorOfAxis2)
    Call DebugValue("Orientation_FromCSVectors pGCVectorOfAxis3 ", pGCVectorOfAxis3)
    
    If Not pGCVectorOfAxis1 Is Nothing Then
        Set pGCVectorOfAxis1 = Vector_Normalize(pGCVectorOfAxis1)
        dDot1 = pGCVectorOfAxis1.Dot(pVectorToOrient)
        Call DebugValue("Orientation_FromCSVectors dDot1 ", dDot1)
        If (1 - Abs(dDot1)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusX
            If dDot1 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusX
            Orientation_FromCSVectors = True
            Call DebugMsg("found exact orientation on X")
            Exit Function
        Else
            dDotMax = dDot1
            bIsMaxSet = True
            eOrientation = eGlobalTrackFlag.GCPlusX
        End If
    End If
    If Not pGCVectorOfAxis2 Is Nothing Then
        Set pGCVectorOfAxis2 = Vector_Normalize(pGCVectorOfAxis2)
        dDot2 = pGCVectorOfAxis2.Dot(pVectorToOrient)
        Call DebugValue("Orientation_FromCSVectors dDot2 ", dDot2)
        If (1 - Abs(dDot2)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusY
            If dDot2 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusY
            Call DebugMsg("found exact orientation on Y")
            Orientation_FromCSVectors = True
            Exit Function
        Else
            If bIsMaxSet = False Or Abs(dDot2) - Abs(dDotMax) > EPSILON Then
                dDotMax = dDot2
                bIsMaxSet = True
                eOrientation = eGlobalTrackFlag.GCPlusY
            End If
        End If
    End If
    If Not pGCVectorOfAxis3 Is Nothing Then
        Set pGCVectorOfAxis3 = Vector_Normalize(pGCVectorOfAxis3)
        dDot3 = pGCVectorOfAxis3.Dot(pVectorToOrient)
        Call DebugValue("Orientation_FromCSVectors dDot3 ", dDot3)
        If (1 - Abs(dDot3)) < EPSILON Then
            'Found the colinear axis, exit
            eOrientation = eGlobalTrackFlag.GCPlusZ
            If dDot3 < -EPSILON Then eOrientation = eGlobalTrackFlag.GCMinusZ
            Call DebugMsg("found exact orientation on Z")
            Orientation_FromCSVectors = True
            Exit Function
        Else
            If bIsMaxSet = False Or Abs(dDot3) - Abs(dDotMax) > EPSILON Then
                dDotMax = dDot3
                bIsMaxSet = True
                eOrientation = eGlobalTrackFlag.GCPlusZ
            End If
        End If
    End If
  
    Call DebugValue("Orientation_FromCSVectors dDotMax ", dDotMax)
    Call DebugValue("Orientation_FromCSVectors eOrientation for max  ", eOrientation)
    If dDotMax < -EPSILON Then
        eOrientation = eOrientation + 1
    End If

End Function

'*******************************************************************************************************************
 'Adaptation of the GC for mirror copy or copy symmetry operations.
 '   This sub is called by the GC evaluate method (CGeometricConstruction::Evaluate) when a mirror is in progress
 '   (CTL_FLAG_MIRROR_IN_PROGRESS flag is set on the BO during mirror operations (CGeometricConstruction::Adapt))
 '   At this stage, the GC output (aggregated geometry) has been processed by the mirror (mirrored /adapted or transformed)
 '   and holds the mirrored geometry that we assume being the exact mirror. The inputs of the mirrored GC are also well defined:
 '   mirrored if they were in the copy set, identified, kept as original, or deleted (delete optional) by the user in the adapt form.
 '
 '   The purpose of this adapt method is to adapt the GC parameters so that the next recompute will match this exact mirrored geometry
 '   or be as close as possible.
 '
 '   FAILURE:
 '   We assume here that any failure or exception is caught by the caller which is the GC BO.
 '
 '   Adaptation of GCKnuckledBracket GC:
 '      Mirror pVectorOfOrientation and look for the closest from the axes of the new 'CoordinateSystem' input.
 '      The 'sORIENTATION' is adapted accordingly.
 '
Private Sub IJGCMirror_Adapt(ByVal pOriginalGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirroredGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMirrorPlane As IngrGeom3D.IJPlane, ByVal pMatrix As AutoMath.IJDT4x4, ByVal pGCInputsFromInitialCopySet As IMSCoreCollections.IJElements)
    Call DebugIn("Method GCKnuckledBracket IJGCMirror_Adapt")
    
    'Compute Mirrored vectorOfOrientation (normal to plateSystem )
    Dim oMirroredPlateSystem As Object: Set oMirroredPlateSystem = pMirroredGC.Input(sPLATE_SYSTEM)
    Dim oMirroredStiffenerSystem1 As Object: Set oMirroredStiffenerSystem1 = pMirroredGC.Input(sSTIFFENER_SYSTEM1)
    Dim oMirroredStiffenerSystem2 As Object: Set oMirroredStiffenerSystem2 = pMirroredGC.Input(sSTIFFENER_SYSTEM2)
    
    Dim pPOMForDebug As IJDPOM
    If GetGlobalShowDetails() Then Set pPOMForDebug = Object_GetPOM(oMirroredPlateSystem)
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oMirroredAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredAxisPortExtractor2 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "002-AxisPortExtractor")
    oMirroredAxisPortExtractor2.Inputs("Connectable").Add pMirroredGC.Inputs(sSTIFFENER_SYSTEM1).Item(1)
    oMirroredAxisPortExtractor2.Evaluate
    
    Dim oMirroredAxisPortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredAxisPortExtractor3 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "003-AxisPortExtractor")
    oMirroredAxisPortExtractor3.Inputs("Connectable").Add pMirroredGC.Inputs(sSTIFFENER_SYSTEM2).Item(1)
    oMirroredAxisPortExtractor3.Evaluate
    
    Dim oMirroredFacePortExtractor04 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredFacePortExtractor04 = oGCFactory.CreateEntity("FacePortExtractor0", pPOMForDebug, "004-FacePortExtractor0")
    oMirroredFacePortExtractor04.Inputs("Connectable").Add oMirroredPlateSystem
    oMirroredFacePortExtractor04.Parameter("Offset") = 0#
    oMirroredFacePortExtractor04.Evaluate
     
    Dim oMirroredPointByCurveAndSurface8 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredPointByCurveAndSurface8 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOMForDebug, "008-PointByCurveAndSurface")
    oMirroredPointByCurveAndSurface8.Inputs("Curve1").Add oMirroredAxisPortExtractor2
    oMirroredPointByCurveAndSurface8.Inputs("Surface2").Add oMirroredFacePortExtractor04
    oMirroredPointByCurveAndSurface8.Parameter("TrackFlag") = GCNear
    Call GCEvaluate(oMirroredPointByCurveAndSurface8, lSTIFFENER_AXIS1_NOT_INTERSECTING)

    Dim oMirroredPointByCurveAndSurface9 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredPointByCurveAndSurface9 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOMForDebug, "009-PointByCurveAndSurface")
    oMirroredPointByCurveAndSurface9.Inputs("Curve1").Add oMirroredAxisPortExtractor3
    oMirroredPointByCurveAndSurface9.Inputs("Surface2").Add oMirroredFacePortExtractor04
    oMirroredPointByCurveAndSurface9.Parameter("TrackFlag") = GCNear
    Call GCEvaluate(oMirroredPointByCurveAndSurface9, lSTIFFENER_AXIS2_NOT_INTERSECTING)
    
    Dim oMirroredLineByPoints11 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredLineByPoints11 = oGCFactory.CreateEntity("LineByPoints", pPOMForDebug, "011-LineByPoints")
    oMirroredLineByPoints11.Inputs("StartPoint").Add oMirroredPointByCurveAndSurface8
    oMirroredLineByPoints11.Inputs("EndPoint").Add oMirroredPointByCurveAndSurface9
    oMirroredLineByPoints11.Evaluate
    
    Dim oMirroredPointAtCurveMiddle12 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredPointAtCurveMiddle12 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOMForDebug, "012-PointAtCurveMiddle")
    oMirroredPointAtCurveMiddle12.Inputs("Curve").Add oMirroredLineByPoints11
    oMirroredPointAtCurveMiddle12.Evaluate
    
    Dim pGCGeomFactory As IJGCGeomFactory: Set pGCGeomFactory = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    Dim pGCGeomDotNetFactory As IJGCGeomDotNetFactory: Set pGCGeomDotNetFactory = CreateObject("GCCMNSTRDefinitions.GCGeomDotNetFactory")
    
    Dim oMirroredPointByProjectOnSurf13 As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredPointByProjectOnSurf13 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOMForDebug, "013-PointByProjectOnSurf")
    oMirroredPointByProjectOnSurf13.Inputs("Point").Add oMirroredPointAtCurveMiddle12
    oMirroredPointByProjectOnSurf13.Inputs("Surface").Add oMirroredFacePortExtractor04
    oMirroredPointByProjectOnSurf13.Evaluate
    
    ' axis z of oCSByLines15 (normal to Plate at midpoint of the line created by the 2 stiffeners)
    Dim oMirroredVectorNormalToSurface13a As SP3DGeometricConstruction.GeometricConstruction
    Set oMirroredVectorNormalToSurface13a = GCEvaluate(pGCGeomDotNetFactory.VectorNormalToSurface.Create(pPOMForDebug, _
                                                                                                 oMirroredFacePortExtractor04, _
                                                                                                 oMirroredPointByProjectOnSurf13, _
                                                                                                 Nothing, _
                                                                                                 1, GCNormal, GCNear), , "013a-VectorNormalToSurface")
   
    
    'Get Original orientation vector
    Dim pVectorOfOrientation As IJDVector
    Dim pLocalCoordinateSystem As IJLocalCoordinateSystem: Set pLocalCoordinateSystem = pOriginalGC.Inputs(sCOORDINATE_SYSTEM).Item(1)
    Select Case pOriginalGC.Parameter(sORIENTATION)
        Case 1: Set pVectorOfOrientation = pLocalCoordinateSystem.XAxis
        Case 2: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.XAxis, -1)
        Case 3: Set pVectorOfOrientation = pLocalCoordinateSystem.YAxis
        Case 4: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.YAxis, -1)
        Case 5: Set pVectorOfOrientation = pLocalCoordinateSystem.ZAxis
        Case 6: Set pVectorOfOrientation = Vector_Scale(pLocalCoordinateSystem.ZAxis, -1)
    End Select
    
    Call DebugValue("Original pVectorOfOrientation: ", pVectorOfOrientation)
    
    'mirror orientation vector
    Dim pMirroredVectorOfOrientation As IJDVector

    Set pMirroredVectorOfOrientation = pMatrix.TransformVector(pVectorOfOrientation)

    Call DebugValue("Mirrored pVectorOfOrientation: ", pMirroredVectorOfOrientation)
    
    'Orient normal of mirrored Platesystem according to pMirroredVectorOfOrientation
    Dim pNewOrientedVectorNormalToPlate As IJDVector
    Set pNewOrientedVectorNormalToPlate = Vector_FromLine(oMirroredVectorNormalToSurface13a)
    If pNewOrientedVectorNormalToPlate.Dot(pMirroredVectorOfOrientation) < -EPSILON Then
        Set pNewOrientedVectorNormalToPlate = Vector_Scale(pNewOrientedVectorNormalToPlate, -1)
    End If
    
    Call DebugValue("Original pNewOrientedVectorNormalToPlate: ", pNewOrientedVectorNormalToPlate)
    
    'write vector in new cs
    Dim pNewCoordinateSystem As IJLocalCoordinateSystem: Set pNewCoordinateSystem = pMirroredGC.Inputs(sCOORDINATE_SYSTEM).Item(1)
    Dim eOrientation As eGlobalTrackFlag
    If Orientation_FromCSVectors(pNewOrientedVectorNormalToPlate, pNewCoordinateSystem.XAxis, pNewCoordinateSystem.YAxis, pNewCoordinateSystem.ZAxis, eOrientation) = False Then
        DebugMsg ("Warning: Mirrored orientation has been approximated to match target cs axis")
    End If
    
    Call DebugValue("mirrored orientation set to ", eOrientation)
    pMirroredGC.Parameter(sORIENTATION) = eOrientation
    
    'mirror justification only if one is not 'Centered'
    If pOriginalGC.Parameter(sJUSTIFICATION1) <> 2 Or pOriginalGC.Parameter(sJUSTIFICATION2) <> 2 Then
    
        'if all inputs are mirrored, justification should be reversed
        'Xaxis is driven by the normal to the plate and Zaxis by the order of the stiffeners
        'the only case where justification should be kept is when both Xaxis and Zaxis are reversed
       
        'The new vector is oMirroredLineByPoints11
        Dim oNewStiffener1ToStiffener2Vector As IJDVector
        Set oNewStiffener1ToStiffener2Vector = Vector_FromLine(oMirroredLineByPoints11)
        
        Call DebugValue("oNewStiffener1ToStiffener2Vector", oNewStiffener1ToStiffener2Vector)
        
        'Get the original oLineByPoints11 vector
        Dim oStiffenerSystem1 As Object: Set oStiffenerSystem1 = pOriginalGC.Input(sSTIFFENER_SYSTEM1)
        Dim oStiffenerSystem2 As Object: Set oStiffenerSystem2 = pOriginalGC.Input(sSTIFFENER_SYSTEM2)
        Dim oPlateSystem As Object: Set oPlateSystem = pOriginalGC.Input(sPLATE_SYSTEM): Call ShowName(oPlateSystem)
        
        Dim oAxisPortExtractor2 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor2 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "002-AxisPortExtractor")
        oAxisPortExtractor2.Inputs("Connectable").Add pOriginalGC.Inputs(sSTIFFENER_SYSTEM1).Item(1)
        oAxisPortExtractor2.Evaluate
            
        Dim oAxisPortExtractor3 As SP3DGeometricConstruction.GeometricConstruction
        Set oAxisPortExtractor3 = oGCFactory.CreateEntity("AxisPortExtractor", pPOMForDebug, "003-AxisPortExtractor")
        oAxisPortExtractor3.Inputs("Connectable").Add pOriginalGC.Inputs(sSTIFFENER_SYSTEM2).Item(1)
        oAxisPortExtractor3.Evaluate
        
        Dim oFacePortExtractor04 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor04 = oGCFactory.CreateEntity("FacePortExtractor0", pPOMForDebug, "004-FacePortExtractor0")
        oFacePortExtractor04.Inputs("Connectable").Add oPlateSystem
        oFacePortExtractor04.Parameter("Offset") = 0#
        oFacePortExtractor04.Evaluate
        
        Dim oPointByCurveAndSurface8 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurveAndSurface8 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOMForDebug, "008-PointByCurveAndSurface")
        oPointByCurveAndSurface8.Inputs("Curve1").Add oAxisPortExtractor2
        oPointByCurveAndSurface8.Inputs("Surface2").Add oFacePortExtractor04
        oPointByCurveAndSurface8.Parameter("TrackFlag") = GCNear
        Call GCEvaluate(oPointByCurveAndSurface8, lSTIFFENER_AXIS1_NOT_INTERSECTING)
    
        Dim oPointByCurveAndSurface9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurveAndSurface9 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOMForDebug, "009-PointByCurveAndSurface")
        oPointByCurveAndSurface9.Inputs("Curve1").Add oAxisPortExtractor3
        oPointByCurveAndSurface9.Inputs("Surface2").Add oFacePortExtractor04
        oPointByCurveAndSurface9.Parameter("TrackFlag") = GCNear
        Call GCEvaluate(oPointByCurveAndSurface9, lSTIFFENER_AXIS2_NOT_INTERSECTING)
           
        Dim oLineByPoints11 As SP3DGeometricConstruction.GeometricConstruction
        Set oLineByPoints11 = oGCFactory.CreateEntity("LineByPoints", pPOMForDebug, "011-LineByPoints")
        oLineByPoints11.Inputs("StartPoint").Add oPointByCurveAndSurface8
        oLineByPoints11.Inputs("EndPoint").Add oPointByCurveAndSurface9
        oLineByPoints11.Evaluate
        
        Dim oOriginalStiffener1ToStiffener2Vector As IJDVector
        Set oOriginalStiffener1ToStiffener2Vector = Vector_FromLine(oLineByPoints11)
        
        Call DebugValue("oOriginalStiffener1ToStiffener2Vector", oNewStiffener1ToStiffener2Vector)
        
        Dim bZaxisReversed As Boolean:  bZaxisReversed = False
        Dim bXaxisReversed As Boolean:  bXaxisReversed = False
        If pVectorOfOrientation.Dot(pNewOrientedVectorNormalToPlate) < -EPSILON Then
            bXaxisReversed = True
        End If
        If oOriginalStiffener1ToStiffener2Vector.Dot(oNewStiffener1ToStiffener2Vector) < -EPSILON Then
            bZaxisReversed = True
        End If
        
        'reverse justification except if both x and z are reversed
        If Not (bXaxisReversed And bZaxisReversed) Then
            Select Case pOriginalGC.Parameter(sJUSTIFICATION1)
                Case 1:
                    pMirroredGC.Parameter(sJUSTIFICATION1) = 3
                Case 2:
                    pMirroredGC.Parameter(sJUSTIFICATION1) = 2
                Case 3:
                    pMirroredGC.Parameter(sJUSTIFICATION1) = 1
            End Select
            
            Select Case pOriginalGC.Parameter(sJUSTIFICATION2)
                Case 1:
                    pMirroredGC.Parameter(sJUSTIFICATION2) = 3
                Case 2:
                    pMirroredGC.Parameter(sJUSTIFICATION2) = 2
                Case 3:
                    pMirroredGC.Parameter(sJUSTIFICATION2) = 1
           End Select
        End If
            
    End If
    
    Call DebugOut
End Sub



Private Function IJKnuckledBracket_Create(ByVal pPOM As RESPOMLib.IJDPOM, _
                                ByVal oStiffenerSystem1 As Object, _
                                ByVal oStiffenerSystem2 As Object, _
                                ByVal oPlateSystem As Object, _
                                ByVal oCoordinateSystem As Object, _
                                ByVal dWidth As Double, _
                                ByVal dKnuckleHeight As Double, _
                                Optional ByVal lKnuckleType As Long = 0&, _
                                Optional ByVal lOrientation As Long = 1&, _
                                Optional ByVal lJustification1 As Long = 2&, _
                                Optional ByVal lJustification2 As Long = 2&, _
                                Optional ByVal dOffset1 As Double = 0#, _
                                Optional ByVal dOffset2 As Double = 0#) As Object
    ' create factory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory
    Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    ' create geometric construction
    Dim pGeometricConstruction As IJGeometricConstruction
    Set pGeometricConstruction = pGeometricConstructionFactory.CreateEntity("KnuckledBracket", pPOM)
    
    ' provide inputs and parameters
    Call pGeometricConstruction.Inputs(sSTIFFENER_SYSTEM1).Add(oStiffenerSystem1, "1")
    Call pGeometricConstruction.Inputs(sSTIFFENER_SYSTEM2).Add(oStiffenerSystem2, "1")
    Call pGeometricConstruction.Inputs(sPLATE_SYSTEM).Add(oPlateSystem, "1")
    Call pGeometricConstruction.Inputs(sCOORDINATE_SYSTEM).Add(oCoordinateSystem, "1")
    pGeometricConstruction.Parameter(sWIDTH) = dWidth
    pGeometricConstruction.Parameter(sKNUCKLE_HEIGHT) = dKnuckleHeight
    pGeometricConstruction.Parameter(sKNUCKLE_TYPE) = lKnuckleType
    pGeometricConstruction.Parameter(sORIENTATION) = lOrientation
    pGeometricConstruction.Parameter(sJUSTIFICATION1) = lJustification1
    pGeometricConstruction.Parameter(sJUSTIFICATION2) = lJustification2
    pGeometricConstruction.Parameter(sOFFSET1) = dOffset1
    pGeometricConstruction.Parameter(sOFFSET2) = dOffset2
    
    ' return geometric construction
    Set IJKnuckledBracket_Create = pGeometricConstruction

End Function
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    ' Delegate to the APS when exists
    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Set IJGCToDoDelegate_ToDoDelegate = pGC.ControlledInputs("AdvancedPlateSystem")(1)
    End If
End Property
