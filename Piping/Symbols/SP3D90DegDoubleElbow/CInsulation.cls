VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         dkl
'   Creation Date:  Wednesday 26, Sep 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ----    ------------------
'   25.Sep.2007     dkl     CR 127557: Created the symbol.
'   07.Jan.2007     dkl     CR 131834: made the symbol trimmable.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_outputcoll As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim depth1           As Double
    Dim flangeDiam1      As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    Dim flangeDiam2      As Double
    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim depth3           As Double
    Dim flangeDiam3      As Double

    Dim iOutput     As Integer

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFacetoCenter As Double
    Dim parSeattoCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parFace1toCenter = arrayOfInputs(3)
'    parFace2toCenter = arrayOfInputs(4)
'    parSeattoCenter = arrayOfInputs(5)
'    parSeat1toCenter = arrayOfInputs(6)
'    parSeat2toCenter = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
'    parAngle = arrayOfInputs(9)

    parAngle = PI / 2
    'Maintaining backward compatibilty with previous versions.
    If UBound(arrayOfInputs) >= 9 Then parAngle = arrayOfInputs(9)

    iOutput = 0
' The origin is assumed to be at intersection of Port 1 and Port 2/Port 3 axes.
    
    RetrieveParameters 1, oPartFclt, m_outputcoll, pipeDiam1, flangeThick1, flangeDiam1, _
                                                                    sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_outputcoll, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                    sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_outputcoll, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                    sptOffset3, depth3
    
    Dim dFace1toCenter As Double
    Dim dFace3toCenter As Double
    Dim dFace2toCenter As Double
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Long
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
' This symbol is based on the following Part data basis values that govern its geometry,
' (Part data Basis value -10): Face-to-center dimension basis.
' (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis.
' (Part data Basis value -35): Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
' (Part data Basis value -36): Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).

    Select Case lPartDataBasis
        Case Is <= 1, 10    'Face-to-center dimension basis.
            parFacetoCenter = arrayOfInputs(2)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
        Case 13    'Asymmetrical Face-to-Center dimension basis.
            parFace1toCenter = arrayOfInputs(3)
            parFace2toCenter = arrayOfInputs(4)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace1toCenter
            dFace3toCenter = parFace2toCenter
        Case 35 'Seat-to-center dimension (Face-to-center dimension excludes the socket depth
                'or thread depth).
            parSeattoCenter = arrayOfInputs(5)
            dFace1toCenter = parSeattoCenter + depth1 - sptOffset1 'towards port 1
            dFace2toCenter = parSeattoCenter + depth2 - sptOffset2 'towards port 2
            dFace3toCenter = parSeattoCenter + depth3 - sptOffset3 'towards port 3
        Case 36 'Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the
                'socket depth or thread depth)
            parSeat1toCenter = arrayOfInputs(6)
            parSeat2toCenter = arrayOfInputs(7)
            dFace1toCenter = parSeat1toCenter + depth1 - sptOffset1 'towards port 1
            dFace2toCenter = parSeat1toCenter + depth2 - sptOffset2 'towards port 2
            dFace3toCenter = parSeat2toCenter + depth3 - sptOffset3 'towards port 3
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    End Select
    
    Dim dTangentLength3 As Double   'Tangent length at Port 3.
    Dim dTangentLength12 As Double  'Tangent length at Port 1 and Port 2.
    
    'Tangent length 3 is minimum length computed based on OD at Port1 and Port2.
    If CmpDblGreaterthan(pipeDiam1, pipeDiam2) Then
        dTangentLength3 = (dFace3toCenter - pipeDiam1 / 2) * 0.8
    Else
        dTangentLength3 = (dFace3toCenter - pipeDiam2 / 2) * 0.8
    End If
    If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
        dTangentLength12 = (dFace1toCenter - pipeDiam3 / 2) * 0.125
    Else
        dTangentLength12 = (dFace2toCenter - pipeDiam3 / 2) * 0.125
    End If
    
    Dim dInsulationDiaPort1 As Double
    dInsulationDiaPort1 = pipeDiam1 + 2 * parInsulationThickness
    Dim dInsulationDiaPort2 As Double
    dInsulationDiaPort2 = pipeDiam2 + 2 * parInsulationThickness
    Dim dInsulationDiaPort3 As Double
    dInsulationDiaPort3 = pipeDiam3 + 2 * parInsulationThickness
    'Restricting the insulation radius to not exceed curvature dimensions.
    If CmpDblGreaterthan(dInsulationDiaPort1 / 2, dFace3toCenter) Then _
        dInsulationDiaPort1 = (dFace3toCenter - dTangentLength3 - NEGLIGIBLE_THICKNESS) * 2
    If CmpDblGreaterthan(dInsulationDiaPort2 / 2, dFace3toCenter) Then _
        dInsulationDiaPort2 = (dFace3toCenter - dTangentLength3 - NEGLIGIBLE_THICKNESS) * 2
    If CmpDblGreaterthan(dInsulationDiaPort3 / 2, dFace1toCenter) Or _
                    CmpDblGreaterthan(dInsulationDiaPort3 / 2, dFace2toCenter) Then _
        dInsulationDiaPort3 = (SmallerDim(dFace1toCenter, dFace2toCenter) - dTangentLength12 - NEGLIGIBLE_THICKNESS) * 2
            
' Insert your code for output 1(BendtowardsPort1)
    'Construct the cross sections for skinning.
    'Forming cross section at port 3 (splitting the circle into a complex string of four
    'quarter-circles to facilitate skinning).
    Dim objArcPort3_1 As IngrGeom3D.EllipticalArc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objArcPort3_1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                dFace3toCenter * Cos(parAngle), -dFace3toCenter * Sin(parAngle), 0, _
                -Cos(parAngle), Sin(parAngle), 0, _
                0, 0, dInsulationDiaPort3 / 2, _
                1, 0, PI / 2)
    Dim oCurvesColl As Collection
    Set oCurvesColl = New Collection
    oCurvesColl.Add objArcPort3_1
    Set objArcPort3_1 = Nothing

    Dim objArcPort3_2 As IngrGeom3D.EllipticalArc3d
    Set objArcPort3_2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            dFace3toCenter * Cos(parAngle), -dFace3toCenter * Sin(parAngle), 0, _
            -Cos(parAngle), Sin(parAngle), 0, _
            dInsulationDiaPort3 / 2 * Sin(parAngle), dInsulationDiaPort3 / 2 * Cos(parAngle), 0, _
            1, 0, PI / 2)
    oCurvesColl.Add objArcPort3_2
    Set objArcPort3_2 = Nothing

    Dim objArcPort3_3 As IngrGeom3D.EllipticalArc3d
    Set objArcPort3_3 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            dFace3toCenter * Cos(parAngle), -dFace3toCenter * Sin(parAngle), 0, _
            -Cos(parAngle), Sin(parAngle), 0, _
            0, 0, -dInsulationDiaPort3 / 2, _
            1, 0, PI / 2)
    oCurvesColl.Add objArcPort3_3
    Set objArcPort3_3 = Nothing

    Dim objArcPort3_4 As IngrGeom3D.EllipticalArc3d
    Set objArcPort3_4 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            dFace3toCenter * Cos(parAngle), -dFace3toCenter * Sin(parAngle), 0, _
            -Cos(parAngle), Sin(parAngle), 0, _
            -dInsulationDiaPort3 / 2 * Sin(parAngle), -dInsulationDiaPort3 / 2 * Cos(parAngle), 0, _
            1, 0, PI / 2)
    oCurvesColl.Add objArcPort3_4
    Set objArcPort3_4 = Nothing
        
    Dim oCrossSectionPort3 As IngrGeom3D.ComplexString3d
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set dFace3toCenter * Cos(parAngle), -dFace3toCenter * Sin(parAngle), dInsulationDiaPort3 / 2
    Set oCrossSectionPort3 = PlaceTrCString(oStPoint, oCurvesColl)
    
    Dim objCrossSections As IngrGeom3D.IJElements
    Set objCrossSections = New JObjectCollection
    objCrossSections.Add oCrossSectionPort3
    Set oCurvesColl = Nothing
    Set oCrossSectionPort3 = Nothing
    
    'Forming cross section at a distance of tangent length above port 3 (splitting the circle into a complex string of four
    'quarter-circles to facilitate skinning).
    Dim objArcTanlenAbovPort3_1 As IngrGeom3D.EllipticalArc3d
    Set objArcTanlenAbovPort3_1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                (dFace3toCenter - dTangentLength3) * Cos(parAngle), -(dFace3toCenter - dTangentLength3) * Sin(parAngle), 0, _
                -Cos(parAngle), Sin(parAngle), 0, _
                0, 0, dInsulationDiaPort3 / 2, _
                1, 0, PI / 2)
    Set oCurvesColl = New Collection
    oCurvesColl.Add objArcTanlenAbovPort3_1
    Set objArcTanlenAbovPort3_1 = Nothing
    
    Dim objArcTanlenAbovPort3_2 As IngrGeom3D.EllipticalArc3d
    Set objArcTanlenAbovPort3_2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            (dFace3toCenter - dTangentLength3) * Cos(parAngle), -(dFace3toCenter - dTangentLength3) * Sin(parAngle), 0, _
            -Cos(parAngle), Sin(parAngle), 0, _
            dInsulationDiaPort3 / 2 * Sin(parAngle), dInsulationDiaPort3 / 2 * Cos(parAngle), 0, _
            1, 0, PI / 2)
    oCurvesColl.Add objArcTanlenAbovPort3_2
    Set objArcTanlenAbovPort3_2 = Nothing
    
    Dim objArcTanlenAbovPort3_3 As IngrGeom3D.EllipticalArc3d
    Set objArcTanlenAbovPort3_3 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                (dFace3toCenter - dTangentLength3) * Cos(parAngle), -(dFace3toCenter - dTangentLength3) * Sin(parAngle), 0, _
                -Cos(parAngle), Sin(parAngle), 0, _
                0, 0, -dInsulationDiaPort3 / 2, _
                1, 0, PI / 2)
    oCurvesColl.Add objArcTanlenAbovPort3_3
    Set objArcTanlenAbovPort3_3 = Nothing
    
    Dim objArcTanlenAbovPort3_4 As IngrGeom3D.EllipticalArc3d
    Set objArcTanlenAbovPort3_4 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                (dFace3toCenter - dTangentLength3) * Cos(parAngle), -(dFace3toCenter - dTangentLength3) * Sin(parAngle), 0, _
                -Cos(parAngle), Sin(parAngle), 0, _
                -dInsulationDiaPort3 / 2 * Sin(parAngle), -dInsulationDiaPort3 / 2 * Cos(parAngle), 0, _
                1, 0, PI / 2)
    oCurvesColl.Add objArcTanlenAbovPort3_4
    Set objArcTanlenAbovPort3_4 = Nothing
       
    Dim oCrossSectionTanlenAbovPort3 As IngrGeom3D.ComplexString3d
    oStPoint.Set (dFace3toCenter - dTangentLength3) * Cos(parAngle), _
                    -(dFace3toCenter - dTangentLength3) * Sin(parAngle), dInsulationDiaPort3 / 2
    Set oCrossSectionTanlenAbovPort3 = PlaceTrCString(oStPoint, oCurvesColl)
    objCrossSections.Add oCrossSectionTanlenAbovPort3
    Set oCurvesColl = Nothing
    
    'Forming cross section at port 1 (splitting the circle into a complex string of two
    'semicircles to facilitate skinning).
    Dim objArcPort1_1 As IngrGeom3D.EllipticalArc3d
    Set objArcPort1_1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            -dFace1toCenter, 0, 0, _
            -1, 0, 0, _
            0, 0, dInsulationDiaPort1 / 2, _
            1, 0, PI)
    Set oCurvesColl = New Collection
    oCurvesColl.Add objArcPort1_1
    Set objArcPort1_1 = Nothing
    
    Dim objArcPort1_2 As IngrGeom3D.EllipticalArc3d
    Set objArcPort1_2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            -dFace1toCenter, 0, 0, _
            -1, 0, 0, _
            0, 0, -dInsulationDiaPort1 / 2, _
            1, 0, PI)
    oCurvesColl.Add objArcPort1_2
    Set objArcPort1_2 = Nothing
    
    Dim oCrossSectionPort1 As IngrGeom3D.ComplexString3d
    oStPoint.Set -dFace1toCenter, 0, dInsulationDiaPort1 / 2
    Set oCrossSectionPort1 = PlaceTrCString(oStPoint, oCurvesColl)
    objCrossSections.Add oCrossSectionPort1
    Set oCurvesColl = Nothing
    Set oCrossSectionPort1 = Nothing

    'Constructing Trace curves for skinning.
    'Define four points representing corners of rectanglar projection of tangent 3 on X-Y plane.
    Dim oPointPosXPort3 As AutoMath.DPosition   'Point at corner of port3 along postive X axis.
    Dim oPointNegXPort3 As AutoMath.DPosition   'Point at corner of port3 along negative X axis.
    Dim oPointPosXTanLenAbovPort3 As AutoMath.DPosition 'Point at tangent length above port3 along postive X axis.
    Dim oPointNegXTanLenAbovPort3 As AutoMath.DPosition 'Point at tangent length above port3 along negative X axis.
    Set oPointPosXPort3 = New AutoMath.DPosition
    Set oPointNegXPort3 = New AutoMath.DPosition
    Set oPointPosXTanLenAbovPort3 = New AutoMath.DPosition
    Set oPointNegXTanLenAbovPort3 = New AutoMath.DPosition
    oPointPosXPort3.Set dFace3toCenter * Cos(parAngle) + (dInsulationDiaPort3 / 2) * Sin(parAngle), _
                        -dFace3toCenter * Sin(parAngle) + (dInsulationDiaPort3 / 2) * Cos(parAngle), 0
    oPointNegXPort3.Set dFace3toCenter * Cos(parAngle) - (dInsulationDiaPort3 / 2) * Sin(parAngle), _
                        -dFace3toCenter * Sin(parAngle) - (dInsulationDiaPort3 / 2) * Cos(parAngle), 0
    oPointPosXTanLenAbovPort3.Set (dFace3toCenter - dTangentLength3) * Cos(parAngle) _
                                                        + (dInsulationDiaPort3 / 2) * Sin(parAngle), _
                                -(dFace3toCenter - dTangentLength3) * Sin(parAngle) _
                                                        + (dInsulationDiaPort3 / 2) * Cos(parAngle), 0
    oPointNegXTanLenAbovPort3.Set (dFace3toCenter - dTangentLength3) * Cos(parAngle) _
                                                        - (dInsulationDiaPort3 / 2) * Sin(parAngle), _
                                -(dFace3toCenter - dTangentLength3) * Sin(parAngle) _
                                                        - (dInsulationDiaPort3 / 2) * Cos(parAngle), 0
    'Outer Trace curve.
    Dim oLine_AlongOuterTangent3 As IngrGeom3D.Line3d
    Set oLine_AlongOuterTangent3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            oPointPosXPort3.x, oPointPosXPort3.y, oPointPosXPort3.z, _
            oPointPosXTanLenAbovPort3.x, oPointPosXTanLenAbovPort3.y, oPointPosXTanLenAbovPort3.z)
    Dim oCurveCollection As Collection
    Set oCurveCollection = New Collection
    oCurveCollection.Add oLine_AlongOuterTangent3

    Dim objEllipticalArcTowarsPort1_AlongOuterCurve As IngrGeom3D.EllipticalArc3d
    Set objEllipticalArcTowarsPort1_AlongOuterCurve = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
            Nothing, _
            -dFace1toCenter + dTangentLength12, oPointPosXTanLenAbovPort3.y, 0, _
            0, 0, 1, _
            dFace1toCenter - dTangentLength12 + oPointPosXTanLenAbovPort3.x, 0, 0, _
            (Abs(oPointPosXTanLenAbovPort3.y) + dInsulationDiaPort1 / 2) / _
                    (dFace1toCenter - dTangentLength12 + oPointPosXTanLenAbovPort3.x), _
            0, PI / 2)
    oCurveCollection.Add objEllipticalArcTowarsPort1_AlongOuterCurve
    Set objEllipticalArcTowarsPort1_AlongOuterCurve = Nothing

    Dim oLine_AlongOuterTangent1 As IngrGeom3D.Line3d
    Set oLine_AlongOuterTangent1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            -dFace1toCenter + dTangentLength12, dInsulationDiaPort1 / 2, 0, _
            -dFace1toCenter, dInsulationDiaPort1 / 2, 0)
    oCurveCollection.Add oLine_AlongOuterTangent1
    Set oLine_AlongOuterTangent1 = Nothing

    oStPoint.Set oPointPosXPort3.x, oPointPosXPort3.y, oPointPosXPort3.z
    Dim oTangentCurve As IngrGeom3D.ComplexString3d
    Set oTangentCurve = PlaceTrCString(oStPoint, oCurveCollection)
    Dim objTraceCurves As IngrGeom3D.IJElements
    Set objTraceCurves = New JObjectCollection
    objTraceCurves.Add oTangentCurve
    Set oTangentCurve = Nothing
    Set oCurveCollection = Nothing

    'Inner trace curve (approximated to facilitate larger insulation thickness)
    Set oCurveCollection = New Collection
    Dim objEllipticalArcTowarsPort1_AlongInnerCurve As IngrGeom3D.EllipticalArc3d
    Set objEllipticalArcTowarsPort1_AlongInnerCurve = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
            Nothing, _
            -dInsulationDiaPort3 / 2, -dInsulationDiaPort1 / 2, 0, _
            0, 0, 1, _
            dFace1toCenter - dInsulationDiaPort3 / 2, 0, 0, _
            (dFace3toCenter - dInsulationDiaPort1 / 2) / (dFace1toCenter - dInsulationDiaPort3 / 2), _
            0, PI / 2)
            
    oCurveCollection.Add objEllipticalArcTowarsPort1_AlongInnerCurve
    Set objEllipticalArcTowarsPort1_AlongInnerCurve = Nothing
    
    oStPoint.Set -dInsulationDiaPort3 / 2, -dFace3toCenter, 0
    Set oTangentCurve = PlaceTrCString(oStPoint, oCurveCollection)
    objTraceCurves.Add oTangentCurve
    Set oTangentCurve = Nothing
    Set oCurveCollection = Nothing
      
    Dim objInsBendfromPort3toPort1 As IngrGeom3D.IJElements
    Set objInsBendfromPort3toPort1 = oGeomFactory.GeometryServices.CreateBySkinning(m_outputcoll.ResourceManager, _
                                            objTraceCurves, objCrossSections, True)
    
'   Set the output
    Dim objOuput As Object
    For Each objOuput In objInsBendfromPort3toPort1
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objOuput
    Next
    Set objOuput = Nothing
    Set objInsBendfromPort3toPort1 = Nothing
    Set objCrossSections = Nothing
    Set objTraceCurves = Nothing
        
' Insert your code for output 2(BendtowardsPort2)
    'Construct the cross sections for skinning.
    Set objCrossSections = New JObjectCollection
    objCrossSections.Add oCrossSectionTanlenAbovPort3
    Set oCrossSectionTanlenAbovPort3 = Nothing

    'Forming cross section at port 2(splitting the circle into a complex string of two
    'semicircles to facilitate skinning).
    Dim objArcPort2_1 As IngrGeom3D.EllipticalArc3d
    Set objArcPort2_1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            dFace2toCenter, 0, 0, _
            1, 0, 0, _
            0, 0, dInsulationDiaPort2 / 2, _
            1, 0, PI)
    Set oCurvesColl = New Collection
    oCurvesColl.Add objArcPort2_1
    Set objArcPort2_1 = Nothing

    Dim objArcPort2_2 As IngrGeom3D.EllipticalArc3d
    Set objArcPort2_2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
            dFace2toCenter, 0, 0, _
            1, 0, 0, _
            0, 0, -dInsulationDiaPort2 / 2, _
            1, 0, PI)
    oCurvesColl.Add objArcPort2_2
    Set objArcPort2_2 = Nothing

    Dim oCrossSectionPort2 As IngrGeom3D.ComplexString3d
    oStPoint.Set dFace2toCenter, 0, dInsulationDiaPort2 / 2
    Set oCrossSectionPort2 = PlaceTrCString(oStPoint, oCurvesColl)
    objCrossSections.Add oCrossSectionPort2
    Set oCurvesColl = Nothing
    Set oCrossSectionPort2 = Nothing

    'Constructing Trace curves for skinning.
    Set oCurveCollection = New Collection
    Dim objEllipticalArcTowarsPort2_AlongOuterCurve As IngrGeom3D.EllipticalArc3d
    Set objEllipticalArcTowarsPort2_AlongOuterCurve = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
            Nothing, _
            dFace2toCenter - dTangentLength12, oPointNegXTanLenAbovPort3.y, 0, _
            0, 0, -1, _
            -(dFace2toCenter - dTangentLength12 - oPointNegXTanLenAbovPort3.x), 0, 0, _
            (Abs(oPointNegXTanLenAbovPort3.y) + dInsulationDiaPort2 / 2) / (dFace2toCenter - dTangentLength12 - oPointNegXTanLenAbovPort3.x), _
            0, PI / 2)
    Set oCurveCollection = New Collection
    oCurveCollection.Add objEllipticalArcTowarsPort2_AlongOuterCurve
    Set objEllipticalArcTowarsPort2_AlongOuterCurve = Nothing

    Dim oLine_AlongOuterTangent2 As IngrGeom3D.Line3d
    Set oLine_AlongOuterTangent2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            dFace2toCenter - dTangentLength12, dInsulationDiaPort2 / 2, 0, _
            dFace2toCenter, dInsulationDiaPort2 / 2, 0)
    oCurveCollection.Add oLine_AlongOuterTangent2
    Set oLine_AlongOuterTangent2 = Nothing
        
    oStPoint.Set oPointNegXTanLenAbovPort3.x, oPointNegXTanLenAbovPort3.y, oPointNegXTanLenAbovPort3.z
    Set oTangentCurve = PlaceTrCString(oStPoint, oCurveCollection)
    Set objTraceCurves = New JObjectCollection
    objTraceCurves.Add oTangentCurve
    Set oTangentCurve = Nothing
    Set oCurveCollection = Nothing

    'Inner trace curve (approximated to facilitate larger insulation thickness).
    Dim objEllipticalArcTowarsPort2_AlongInnerCurve As IngrGeom3D.EllipticalArc3d
    Set objEllipticalArcTowarsPort2_AlongInnerCurve = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
        Nothing, _
        dInsulationDiaPort3 / 2, -dInsulationDiaPort2 / 2, 0, _
        0, 0, -1, _
        -(dFace2toCenter - dInsulationDiaPort3 / 2), 0, 0, _
        (dFace3toCenter - dInsulationDiaPort2 / 2) / (dFace2toCenter - dInsulationDiaPort3 / 2), _
        0, PI / 2)

    Set oCurveCollection = New Collection
    oCurveCollection.Add objEllipticalArcTowarsPort2_AlongInnerCurve
    oStPoint.Set dInsulationDiaPort3 / 2, -dFace3toCenter, 0
    Dim oTangentCurve1 As IngrGeom3D.ComplexString3d
    Set oTangentCurve1 = PlaceTrCString(oStPoint, oCurveCollection)
    objTraceCurves.Add oTangentCurve1
    Set oTangentCurve = Nothing
    Set oCurveCollection = Nothing
    Set objEllipticalArcTowarsPort2_AlongInnerCurve = Nothing

    Dim objInsBendfromPort3toPort2 As IngrGeom3D.IJElements
    Set objInsBendfromPort3toPort2 = oGeomFactory.GeometryServices.CreateBySkinning(m_outputcoll.ResourceManager, _
                                            objTraceCurves, objCrossSections, True)
    
'   Set the output
    For Each objOuput In objInsBendfromPort3toPort2
        iOutput = iOutput + 1
        m_outputcoll.AddOutput arrayOfOutputs(iOutput), objOuput
    Next
    Set objOuput = Nothing
    Set objInsBendfromPort3toPort2 = Nothing
    Set objTraceCurves = Nothing
    Set objCrossSections = Nothing
    Set oGeomFactory = Nothing
    
'Ports are insulated only when Flanged/Female ends.
  ' Insert your code Insulated Port 1.
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    If CmpDblGreaterthan(flangeDiam1, LINEAR_TOLERANCE) Then
        oStPoint.Set -dFace1toCenter, 0, 0
        If CmpDblLessThanOrEqualTo(flangeThick1, 0) Then flangeThick1 = NEGLIGIBLE_THICKNESS
        oEnPoint.Set -dFace1toCenter + flangeThick1 + parInsulationThickness, 0, 0
        Dim dInsulationRadius As Double
        dInsulationRadius = flangeDiam1 / 2 + parInsulationThickness
        'Restricting the insulation to not exceed face-to-center dimensions.
        If CmpDblGreaterthan(dInsulationRadius, dFace3toCenter) Then _
            dInsulationRadius = (dFace3toCenter - NEGLIGIBLE_THICKNESS)
        Dim objInsulatedPort1 As Object
        Set objInsulatedPort1 = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, _
                                            2 * dInsulationRadius, True)
    
        'Set the output
        m_outputcoll.AddOutput "InsulatedPort1", objInsulatedPort1
        Set objInsulatedPort1 = Nothing
    End If

 ' Insert your code Insulated Port 2.
    If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
        oStPoint.Set dFace2toCenter, 0, 0
        If CmpDblLessThanOrEqualTo(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
        oEnPoint.Set dFace2toCenter - flangeThick2 - parInsulationThickness, 0, 0
        dInsulationRadius = flangeDiam2 / 2 + parInsulationThickness
        'Restricting the insulation to not exceed face-to-center dimensions.
        If CmpDblGreaterthan(dInsulationRadius, dFace3toCenter) Then _
            dInsulationRadius = dFace3toCenter - NEGLIGIBLE_THICKNESS
        Dim objInsulatedPort2 As Object
        Set objInsulatedPort2 = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, _
                                    2 * dInsulationRadius, True)
    
        'Set the output
        m_outputcoll.AddOutput "InsulatedPort2", objInsulatedPort2
        Set objInsulatedPort2 = Nothing
    End If
    
' Insert your code Insulated Port 3.
    If CmpDblGreaterthan(flangeDiam3, LINEAR_TOLERANCE) Then
        oStPoint.Set 0, -dFace3toCenter, 0
        If CmpDblLessThanOrEqualTo(flangeThick3, 0) Then flangeThick3 = NEGLIGIBLE_THICKNESS
        oEnPoint.Set 0, -dFace3toCenter + flangeThick3 + parInsulationThickness, 0
        dInsulationRadius = flangeDiam3 / 2 + parInsulationThickness
        'Restricting the insulation to not exceed face-to-center dimensions.
        If CmpDblGreaterthan(dInsulationRadius, dFace1toCenter) Then _
            dInsulationRadius = dFace1toCenter - NEGLIGIBLE_THICKNESS
        Dim objInsulatedPort3 As Object
        Set objInsulatedPort3 = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, _
                                        2 * dInsulationRadius, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_outputcoll.AddOutput "InsulatedPort3", objInsulatedPort3
        Set objInsulatedPort3 = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
