VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StraightEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJGeometricConstructionDefinitionService
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("Support1", "Select support", "IJPlane OR IJSurfaceBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJPlane]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Support2", "Select support", "IJPlane OR IJSurfaceBody AND [GCCMNSTRDefinitions.Filters,ImplementsIJPlane]", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("MemberAxis1", "Select member axis 1", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("MemberAxis2", "Select member axis 2", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("EdgePort1", "Select edge port 1", "IJCurve OR IJWireBody", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("EdgePort2", "Select edge port 2", "IJCurve OR IJWireBody", 1, 1)
    
    Call pGeometricConstructionDefinition.AddParameter("WeldToe", "WeldToe", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("CutBack", "CutBack", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("RibbonHeight", "RibbonHeight", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1#)
    
    Call pGeometricConstructionDefinition.AddParameter("Debug", "Debug", GCCodeList, 0, 0, 0, 0, 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "Off", 0)
    Call pGeometricConstructionDefinition.AddParameterValue("Debug", "On", 1)
    
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointX")
    Call pGeometricConstructionDefinition.AddOutput(1, "PointY")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "001-ParamDistConstant")
    oParamDistConstant1.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("WeldToe"))
    oParamDistConstant1.Evaluate
    oParamDistConstant1.Public = True
    
    Dim oParamDistConstant2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant2 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "002-ParamDistConstant")
    oParamDistConstant2.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("CutBack"))
    oParamDistConstant2.Evaluate
    oParamDistConstant2.Public = True
    
    Dim oParamDistConstant3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant3 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, "003-ParamDistConstant")
    oParamDistConstant3.Parameter("Value") = CDbl(pGeometricConstruction.Parameter("RibbonHeight"))
    oParamDistConstant3.Evaluate
    oParamDistConstant3.Public = True
    
    Dim oPointAtCurveStart3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart3 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "004-PointAtCurveStart")
    oPointAtCurveStart3.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis1").Item(1)
    oPointAtCurveStart3.Evaluate
    
    Dim oPointAtCurveMiddle3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle3 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "005-PointAtCurveMiddle")
    oPointAtCurveMiddle3.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis1").Item(1)
    oPointAtCurveMiddle3.Evaluate
 
    Dim oPointAtCurveEnd3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd3 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "006-PointAtCurveEnd")
    oPointAtCurveEnd3.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis1").Item(1)
    oPointAtCurveEnd3.Evaluate
    
    Dim oPointAtCurveStart4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveStart4 = oGCFactory.CreateEntity("PointAtCurveStart", pPOM, "007-PointAtCurveStart")
    oPointAtCurveStart4.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis2").Item(1)
    oPointAtCurveStart4.Evaluate
    
    Dim oPointAtCurveMiddle4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle4 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "008-PointAtCurveMiddle")
    oPointAtCurveMiddle4.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis2").Item(1)
    oPointAtCurveMiddle4.Evaluate
    
    Dim oPointAtCurveEnd4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd4 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "009-PointAtCurveEnd")
    oPointAtCurveEnd4.Inputs("Curve").Add pGeometricConstruction.Inputs("MemberAxis2").Item(1)
    oPointAtCurveEnd4.Evaluate

    Dim oPointAtCurveMiddle5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle5 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "010-PointAtCurveMiddle")
    oPointAtCurveMiddle5.Inputs("Curve").Add pGeometricConstruction.Inputs("EdgePort1").Item(1)
    oPointAtCurveMiddle5.Evaluate

    Dim oPointAtCurveMiddle6 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle6 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, "011-PointAtCurveMiddle")
    oPointAtCurveMiddle6.Inputs("Curve").Add pGeometricConstruction.Inputs("EdgePort2").Item(1)
    oPointAtCurveMiddle6.Evaluate

    Dim oCSByPoints1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints1 = oGCFactory.CreateEntity("CSByPoints", pPOM, "012-CSByPoints")
    oCSByPoints1.Inputs("OriginPoint").Add oPointAtCurveMiddle3
    oCSByPoints1.Inputs("AxisPoint1").Add oPointAtCurveMiddle5
    oCSByPoints1.Inputs("AxisPoint2").Add oPointAtCurveMiddle4
    oCSByPoints1.Parameter("AxesRoles") = 1
    oCSByPoints1.Parameter("CSOrientation") = 1
    oCSByPoints1.Evaluate
    
    Dim oCSByPoints2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints2 = oGCFactory.CreateEntity("CSByPoints", pPOM, "013-CSByPoints")
    oCSByPoints2.Inputs("OriginPoint").Add oPointAtCurveMiddle4
    oCSByPoints2.Inputs("AxisPoint1").Add oPointAtCurveMiddle6
    oCSByPoints2.Inputs("AxisPoint2").Add oPointAtCurveMiddle3
    oCSByPoints2.Parameter("AxesRoles") = 1
    oCSByPoints2.Parameter("CSOrientation") = 1
    oCSByPoints2.Evaluate
    oCSByPoints2.Public = True
    
    Dim oPointFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS1 = oGCFactory.CreateEntity("PointFromCS", pPOM, "014-PointFromCS")
    oPointFromCS1.Inputs("CoordinateSystem").Add oCSByPoints1
    oPointFromCS1.Parameter("X") = 0#
    oPointFromCS1.Parameter("Y") = 0#
    oPointFromCS1.Parameter("Z") = 0#
    oPointFromCS1.Evaluate
    
    Dim oPointFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS2 = oGCFactory.CreateEntity("PointFromCS", pPOM, "015-PointFromCS")
    oPointFromCS2.Inputs("CoordinateSystem").Add oCSByPoints2
    oPointFromCS2.Parameter("X") = 0#
    oPointFromCS2.Parameter("Y") = 0#
    oPointFromCS2.Parameter("Z") = 0#
    oPointFromCS2.Evaluate
    
    Dim oLineFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS1 = oGCFactory.CreateEntity("LineFromCS", pPOM, "016-LineFromCS")
    oLineFromCS1.Inputs("CoordinateSystem").Add oCSByPoints1
    oLineFromCS1.Parameter("LookingAxis") = 1
    oLineFromCS1.Parameter("Length") = 10
    oLineFromCS1.Evaluate

    Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, "017-LineFromCS")
    oLineFromCS2.Inputs("CoordinateSystem").Add oCSByPoints2
    oLineFromCS2.Parameter("LookingAxis") = 1
    oLineFromCS2.Parameter("Length") = 10
    oLineFromCS2.Evaluate

    Dim oPointAlongCurve1a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve1a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "018-PointAlongCurve")
    oPointAlongCurve1a.Inputs("Curve").Add oLineFromCS1
    oPointAlongCurve1a.Inputs("Point").Add oPointAtCurveMiddle5
    oPointAlongCurve1a.Parameter("Distance") = 0#
    oPointAlongCurve1a.Parameter("TrackFlag") = 1
    oPointAlongCurve1a.Evaluate
   
    Dim oPointAlongCurve1b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve1b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "019-PointAlongCurve")
    oPointAlongCurve1b.Inputs("Curve").Add oLineFromCS1
    oPointAlongCurve1b.Inputs("Point").Add oPointAtCurveMiddle6
    oPointAlongCurve1b.Parameter("Distance") = 0#
    oPointAlongCurve1b.Parameter("TrackFlag") = 1
    oPointAlongCurve1b.Evaluate
   
    Dim oPointAlongCurve2a As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve2a = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "020-PointAlongCurve")
    oPointAlongCurve2a.Inputs("Curve").Add oLineFromCS2
    oPointAlongCurve2a.Inputs("Point").Add oPointAtCurveMiddle5
    oPointAlongCurve2a.Parameter("Distance") = 0#
    oPointAlongCurve2a.Parameter("TrackFlag") = 1
    oPointAlongCurve2a.Evaluate
   
    Dim oPointAlongCurve2b As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve2b = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "021-PointAlongCurve")
    oPointAlongCurve2b.Inputs("Curve").Add oLineFromCS2
    oPointAlongCurve2b.Inputs("Point").Add oPointAtCurveMiddle6
    oPointAlongCurve2b.Parameter("Distance") = 0#
    oPointAlongCurve2b.Parameter("TrackFlag") = 1
    oPointAlongCurve2b.Evaluate
   
    Dim oParamDistMeasureParallel1a As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel1a = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "022-ParamDistMeasureParallel")
    oParamDistMeasureParallel1a.Inputs("Graphics").Add oPointFromCS1
    oParamDistMeasureParallel1a.Inputs("Graphics").Add oPointAlongCurve1a
    oParamDistMeasureParallel1a.Inputs("OrientationLine").Add oLineFromCS1
    oParamDistMeasureParallel1a.Parameter("Value") = 0
    oParamDistMeasureParallel1a.Evaluate

    Dim oParamDistMeasureParallel1b As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel1b = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "023-ParamDistMeasureParallel")
    oParamDistMeasureParallel1b.Inputs("Graphics").Add oPointFromCS1
    oParamDistMeasureParallel1b.Inputs("Graphics").Add oPointAlongCurve1b
    oParamDistMeasureParallel1b.Inputs("OrientationLine").Add oLineFromCS1
    oParamDistMeasureParallel1b.Parameter("Value") = 0
    oParamDistMeasureParallel1b.Evaluate
    
    Dim oParamDistMeasureParallel2a As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel2a = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "024-ParamDistMeasureParallel")
    oParamDistMeasureParallel2a.Inputs("Graphics").Add oPointFromCS2
    oParamDistMeasureParallel2a.Inputs("Graphics").Add oPointAlongCurve2a
    oParamDistMeasureParallel2a.Inputs("OrientationLine").Add oLineFromCS2
    oParamDistMeasureParallel2a.Parameter("Value") = 0
    oParamDistMeasureParallel2a.Evaluate

    Dim oParamDistMeasureParallel2b As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel2b = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, "025-ParamDistMeasureParallel")
    oParamDistMeasureParallel2b.Inputs("Graphics").Add oPointFromCS2
    oParamDistMeasureParallel2b.Inputs("Graphics").Add oPointAlongCurve2b
    oParamDistMeasureParallel2b.Inputs("OrientationLine").Add oLineFromCS2
    oParamDistMeasureParallel2b.Parameter("Value") = 0
    oParamDistMeasureParallel2b.Evaluate
    
    Dim oParamDistFunctionWithNArgs1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs1 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "026-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs1.Inputs("Args").Add oParamDistMeasureParallel1a
    oParamDistFunctionWithNArgs1.Inputs("Args").Add oParamDistMeasureParallel1b
    oParamDistFunctionWithNArgs1.Parameter("Value") = 0
    oParamDistFunctionWithNArgs1.Parameter("Operation") = 2
    oParamDistFunctionWithNArgs1.Evaluate
    oParamDistFunctionWithNArgs1.Public = True
    
    Dim oParamDistFunctionWithNArgs2 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistFunctionWithNArgs2 = oGCFactory.CreateEntity("ParamDistFunctionWithNArgs", pPOM, "027-ParamDistFunctionWithNArgs")
    oParamDistFunctionWithNArgs2.Inputs("Args").Add oParamDistMeasureParallel2a
    oParamDistFunctionWithNArgs2.Inputs("Args").Add oParamDistMeasureParallel2b
    oParamDistFunctionWithNArgs2.Parameter("Value") = 0
    oParamDistFunctionWithNArgs2.Parameter("Operation") = 2
    oParamDistFunctionWithNArgs2.Evaluate
    oParamDistFunctionWithNArgs2.Public = True
       
    Dim oPointAlongCurve1c As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve1c = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "028-PointAlongCurve")
    oPointAlongCurve1c.Inputs("Curve").Add oLineFromCS1
    oPointAlongCurve1c.Parameter("Distance") = CDbl(oParamDistFunctionWithNArgs1.Parameter("Value")) + CDbl(oParamDistConstant1.Parameter("Value"))
    oPointAlongCurve1c.Evaluate
        
    Dim oPointAlongCurve2c As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve2c = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "029-PointAlongCurve")
    oPointAlongCurve2c.Inputs("Curve").Add oLineFromCS2
    oPointAlongCurve2c.Parameter("Distance") = CDbl(oParamDistFunctionWithNArgs2.Parameter("Value")) + CDbl(oParamDistConstant1.Parameter("Value"))
    oPointAlongCurve2c.Evaluate
        
    Dim oLineAtAngle1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle1 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "030-LineAtAngle")
    oLineAtAngle1.Inputs("Line").Add pGeometricConstruction.Inputs("EdgePort1").Item(1)
    oLineAtAngle1.Inputs("Point").Add oPointAlongCurve1c
    oLineAtAngle1.Inputs("Plane").Add pGeometricConstruction.Inputs("Support1").Item(1)
    oLineAtAngle1.Parameter("Angle") = 0#
    oLineAtAngle1.Evaluate

    Dim oLineAtAngle2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineAtAngle2 = oGCFactory.CreateEntity("LineAtAngle", pPOM, "031-LineAtAngle")
    oLineAtAngle2.Inputs("Line").Add pGeometricConstruction.Inputs("EdgePort2").Item(1)
    oLineAtAngle2.Inputs("Point").Add oPointAlongCurve2c
    oLineAtAngle2.Inputs("Plane").Add pGeometricConstruction.Inputs("Support2").Item(1)
    oLineAtAngle2.Parameter("Angle") = 0#
    oLineAtAngle2.Evaluate

    Dim oPointAtCurveEnd1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd1 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "032-PointAtCurveEnd")
    oPointAtCurveEnd1.Inputs("Curve").Add oLineAtAngle1
    oPointAtCurveEnd1.Evaluate

    Dim oPointAtCurveEnd2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd2 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "033-PointAtCurveEnd")
    oPointAtCurveEnd2.Inputs("Curve").Add oLineAtAngle2
    oPointAtCurveEnd2.Evaluate

    Dim oLineByPoints7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineByPoints7 = oGCFactory.CreateEntity("LineByPoints", pPOM, "034-LineByPoints")
    oLineByPoints7.Inputs("StartPoint").Add oPointAtCurveEnd1
    oLineByPoints7.Inputs("EndPoint").Add oPointAtCurveEnd2
    oLineByPoints7.Evaluate
    oLineByPoints7.Public = True

    Dim oLineFromCS7 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS7 = oGCFactory.CreateEntity("LineFromCS", pPOM, "035-LineFromCS")
    oLineFromCS7.Inputs("CoordinateSystem").Add oCSByPoints1
    oLineFromCS7.Parameter("LookingAxis") = 3
    oLineFromCS7.Parameter("Length") = -CDbl(oParamDistConstant3.Parameter("Value")) / 2
    oLineFromCS7.Evaluate
    
    Dim oSurfByLinearExtrusion7 As SP3DGeometricConstruction.GeometricConstruction
    Set oSurfByLinearExtrusion7 = oGCFactory.CreateEntity("SurfByLinearExtrusion", pPOM, "036-SurfByLinearExtrusion")
    oSurfByLinearExtrusion7.Inputs("PlanarCrossSection").Add oLineByPoints7
    oSurfByLinearExtrusion7.Inputs("ExtrusionLine").Add oLineFromCS7
    oSurfByLinearExtrusion7.Evaluate
    oSurfByLinearExtrusion7.Public = True

    Dim oPointAlongCurve1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve1 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "037-PointAlongCurve")
    oPointAlongCurve1.Inputs("Curve").Add oLineByPoints7
    oPointAlongCurve1.Inputs("Point").Add oPointAtCurveStart3
    oPointAlongCurve1.Inputs("TrackPoint").Add oPointAtCurveEnd3
    oPointAlongCurve1.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve1.Parameter("TrackFlag") = 1
    oPointAlongCurve1.Evaluate
    oPointAlongCurve1.Public = True

    Dim oPointAlongCurve2 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve2 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, "038-PointAlongCurve")
    oPointAlongCurve2.Inputs("Curve").Add oLineByPoints7
    oPointAlongCurve2.Inputs("Point").Add oPointAtCurveStart4
    oPointAlongCurve2.Inputs("TrackPoint").Add oPointAtCurveEnd4
    oPointAlongCurve2.Parameter("Distance") = oParamDistConstant2
    oPointAlongCurve2.Parameter("TrackFlag") = 1
    oPointAlongCurve2.Evaluate
    oPointAlongCurve2.Public = True
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("PointY", 1) = oPointAlongCurve1.Output
    oGCMacro.Output("PointX", 1) = oPointAlongCurve2.Output
    oGCMacro.Output("Boundary", 1) = oSurfByLinearExtrusion7.Output
End Sub
