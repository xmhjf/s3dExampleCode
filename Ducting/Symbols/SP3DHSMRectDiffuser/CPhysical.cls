VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DHSMRectDiffuser.HSMRectDiffuser
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   26.04.2006      kkc     DI-97251 Removed hard coding for parWidth and parDepth.
'                           Added condition to check the baffle spacing.
'  08.SEP.2006      KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'  12.DEC.2007      MA      CR-123922 Implemented Part Data Basis for the value 119.
'  22.JAN.2008      MA      CR-135023 Implemented Part Data Basis for the values 120,121,122 and modified Part Data Basis Case 119 to use optional inputs.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private RAD As Double
Const NEGLIGIBLE_VALUE = 0.0001
Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim CptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim iOutput As Double
    Dim ObjDiffuserPlate As Object
    Dim ObjDiffuserBaffle1 As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle2 As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle3 As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle4 As IngrGeom3D.RuledSurface3d
    Dim ObjDiffuserBaffle5 As IngrGeom3D.RuledSurface3d

    Dim parWidth As Double
    Dim parDepth As Double
    Dim parNeckWidth As Double
    Dim parNeckDepth As Double
    Dim parDiffuserHeight As Double
    Dim parNeckHeight As Double
    Dim parDamperHeight As Double

    Dim dblExtendedWidth As Double
    Dim dblExtendedDepth As Double
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim dblX As Double
    Dim parlength As Double
    Dim dblDiffuserLength As Double
    Dim dblPnts1(0 To 11) As Double
    Dim dblPnts2(0 To 11) As Double
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New GeometryFactory
    Dim ObjOuterPlane As IngrGeom3D.Plane3d
    Dim oLine1 As IngrGeom3D.Line3d
    Dim oLine2 As IngrGeom3D.Line3d
    Dim oLine3 As IngrGeom3D.Line3d
    Dim oLine4 As IngrGeom3D.Line3d
    Dim oInnerBoundary As Collection
    Dim objHoleF As IngrGeom3D.ComplexString3d
    Dim Startcurve As New AutoMath.DPosition

    Dim dWidth5 As Double
    Dim dDepth5 As Double
    Dim dStripWidth As Double
    Dim dWidth4 As Double
    Dim dDepth4 As Double
    Dim dWidth3 As Double
    Dim dDepth3 As Double
    Dim dWidth2 As Double
    Dim dDepth2 As Double
    Dim dWidth1 As Double
    Dim dDepth1 As Double
    Dim dWidth0 As Double
    Dim dDepth0 As Double
    Dim ObjCyl As Object
    Dim oDamperVane As Object
    Dim dSpaceBtwVanes As Double
    Dim dVaneThickness As Double
    Dim lNumOfVanes As Long
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVec As AutoMath.DVector
    Dim oTransMatrix As DT4x4
    Dim oTransVector As AutoMath.DVector
    Dim dConst As Double
    Dim oRotVector As AutoMath.DVector
    Dim iCount As Integer
    Dim dYCoord As Double
    Dim dHeight As Double
    Dim dDiffuserHeight As Double
    Dim dDiffuserWidth As Double
    Dim dDiffuserDepth As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    '    parWidth = arrayOfInputs(2) 'A
    '    parDepth = arrayOfInputs(3) 'B
    '    parNeckWidth = arrayOfInputs(4)
    '    parNeckDepth = arrayOfInputs(5)
    '    parDiffuserHeight = arrayOfInputs(6)
    '    parNeckHeight = arrayOfInputs(7)
    '    parDamperHeight = arrayOfInputs(8)

    iOutput = 0

    ' This symbol is based on the following Part data basis values that govern its geometry
    ' (Part data Basis value -118): Rectangular Diffuser, Rectangular Neck without Damper
    ' (Part data Basis value -119): Rectangular Diffuser, Rectangular Neck with Damper
    ' (Part data Basis value -120): Rectangular Diffuser, Round Neck without Damper
    ' (Part data Basis value -121): Rectangular Diffuser, Round Neck with Damper
    ' (Part data Basis value -122): Rectangular Diffuser, Rectangular Neck without Damper,
    '                                            specified by Diffuser Height

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Dim tempStartPos As Double

    If lPartDataBasis = 119 Then
        parNeckWidth = arrayOfInputs(4)
        parNeckDepth = arrayOfInputs(5)
        parDiffuserHeight = arrayOfInputs(6)
        parDamperHeight = arrayOfInputs(8)
        dHeight = 0.25 * parDiffuserHeight
        tempStartPos = parDamperHeight + dHeight
    
    ElseIf lPartDataBasis = 120 Then
        parNeckWidth = arrayOfInputs(4)
        parDiffuserHeight = arrayOfInputs(6)
        parNeckHeight = arrayOfInputs(7)
        parNeckDepth = parNeckWidth    'Round Neck
        dHeight = 0.25 * parDiffuserHeight
        tempStartPos = parNeckHeight + dHeight
    
    ElseIf lPartDataBasis = 121 Then
        parNeckWidth = arrayOfInputs(4)
        parDiffuserHeight = arrayOfInputs(6)
        parNeckHeight = arrayOfInputs(7)
        parDamperHeight = arrayOfInputs(8)
        parNeckDepth = parNeckWidth    'Round Neck
        dHeight = 0.25 * parDiffuserHeight
        tempStartPos = parDamperHeight + parNeckHeight + dHeight
    
    ElseIf lPartDataBasis = 122 Then
        parNeckWidth = arrayOfInputs(4)
        parNeckDepth = arrayOfInputs(5)
        parDiffuserHeight = arrayOfInputs(6)
        dHeight = 0.25 * parDiffuserHeight
        tempStartPos = dHeight
    
    End If
    
    If (lPartDataBasis = 119) Or (lPartDataBasis = 120) Or (lPartDataBasis = 121) Or _
                                            (lPartDataBasis = 122) Then
    
        dDiffuserWidth = parNeckWidth + 0.12
        dDiffuserDepth = parNeckDepth + 0.12
        dHeight = 0.25 * parDiffuserHeight
        dDiffuserHeight = parDiffuserHeight - dHeight
        dStripWidth = 0.0508
        'Check and adjust StripWidth value to accomodate 5 baffles
        If CmpDblGreaterthan(dStripWidth * 5, parNeckWidth) Then
            dStripWidth = parNeckWidth / 5
        End If
        If CmpDblGreaterthan(dStripWidth * 5, parNeckDepth) Then
            dStripWidth = parNeckDepth / 5
        End If

        dWidth5 = dDiffuserWidth - dStripWidth
        dDepth5 = dDiffuserDepth - dStripWidth

        dWidth4 = dWidth5 - dStripWidth
        dDepth4 = dDepth5 - dStripWidth

        dWidth3 = dWidth4 - dStripWidth
        dDepth3 = dDepth4 - dStripWidth

        dWidth2 = dWidth3 - dStripWidth
        dDepth2 = dDepth3 - dStripWidth

        dWidth1 = dWidth2 - dStripWidth
        dDepth1 = dDepth2 - dStripWidth

        dWidth0 = dWidth1 - dStripWidth
        dDepth0 = dDepth1 - dStripWidth
        
        'Insert your code for output 3(Diffuser Stiffner -1)
        oPos.Set -dDiffuserWidth / 2, -dDiffuserDepth / 2, tempStartPos + dDiffuserHeight

        dblPnts1(0) = oPos.x
        dblPnts1(1) = oPos.y
        dblPnts1(2) = oPos.z

        dblPnts1(3) = oPos.x
        dblPnts1(4) = oPos.y + dDiffuserDepth
        dblPnts1(5) = oPos.z

        dblPnts1(6) = oPos.x + dDiffuserWidth
        dblPnts1(7) = oPos.y + dDiffuserDepth
        dblPnts1(8) = oPos.z

        dblPnts1(9) = oPos.x + dDiffuserWidth
        dblPnts1(10) = oPos.y
        dblPnts1(11) = oPos.z

        Set ObjOuterPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblPnts1)

        oPos.Set -dWidth5 / 2, -dDepth5 / 2, tempStartPos + dDiffuserHeight

        dblPnts2(0) = oPos.x
        dblPnts2(1) = oPos.y
        dblPnts2(2) = oPos.z

        dblPnts2(3) = oPos.x + dWidth5
        dblPnts2(4) = oPos.y
        dblPnts2(5) = oPos.z

        dblPnts2(6) = oPos.x + dWidth5
        dblPnts2(7) = oPos.y + dDepth5
        dblPnts2(8) = oPos.z

        dblPnts2(9) = oPos.x
        dblPnts2(10) = oPos.y + dDepth5
        dblPnts2(11) = oPos.z

        Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(0), dblPnts2(1), dblPnts2(2), dblPnts2(3), dblPnts2(4), dblPnts2(5))

        Set oLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(3), dblPnts2(4), dblPnts2(5), dblPnts2(6), dblPnts2(7), dblPnts2(8))

        Set oLine3 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(6), dblPnts2(7), dblPnts2(8), dblPnts2(9), dblPnts2(10), dblPnts2(11))

        Set oLine4 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(9), dblPnts2(10), dblPnts2(11), dblPnts2(0), dblPnts2(1), dblPnts2(2))

        Set oInnerBoundary = New Collection
        oInnerBoundary.Add oLine1
        oInnerBoundary.Add oLine2
        oInnerBoundary.Add oLine3
        oInnerBoundary.Add oLine4

        Startcurve.Set dblPnts2(0), dblPnts2(1), dblPnts2(2)

        Set objHoleF = PlaceTrCString(Startcurve, oInnerBoundary)
        ObjOuterPlane.AddBoundary objHoleF

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOuterPlane

        Set ObjOuterPlane = Nothing
        Set objHoleF = Nothing
        Set oPos = Nothing
        
        'Insert your code for output 4(Diffuser Baffle -2)
        stPoint.Set 0, 0, tempStartPos
        Set ObjDiffuserBaffle2 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth3, dDepth3, dWidth4, dDepth4, _
                                                dDiffuserHeight, False)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle2
        Set ObjDiffuserBaffle2 = Nothing
    
        'Insert your code for output 5(Diffuser Baffle -3)
        stPoint.Set 0, 0, tempStartPos
    
        Set ObjDiffuserBaffle3 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth2, dDepth2, dWidth3, dDepth3, _
                                                dDiffuserHeight, False)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle3
        Set ObjDiffuserBaffle3 = Nothing
    
        'Insert your code for output 6(Diffuser Baffle -4)
        stPoint.Set 0, 0, tempStartPos
    
        Set ObjDiffuserBaffle4 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth1, dDepth1, dWidth2, dDepth2, _
                                                dDiffuserHeight, False)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle4
        Set ObjDiffuserBaffle4 = Nothing
    
    
        'Insert your code for output 7(Diffuser Baffle -5)
        stPoint.Set 0, 0, tempStartPos
    
        Set ObjDiffuserBaffle5 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth0, dDepth0, dWidth1, dDepth1, _
                                                dDiffuserHeight, False)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle5
        Set ObjDiffuserBaffle5 = Nothing
    
    End If

    Select Case lPartDataBasis
    Case Is <= 1, 118

        parWidth = arrayOfInputs(2)    'A
        parDepth = arrayOfInputs(3)    'B
        'X = 4.25 Inches for the SESA-41 Spiral Catalogue Model.
        dblX = (25.4 / 1000) * 4.25    ' 4.25 inches
        dblDiffuserLength = 1.1875 * 25.4 / 1000
        dblExtendedWidth = parWidth + dblX    'A+X
        dblExtendedDepth = parDepth + dblX    ' B+X
        parlength = 25.4 / 1000

        Dim parWidth5 As Double
        Dim parDepth5 As Double
        Dim StripWidth As Double

        Dim parWidth4 As Double
        Dim parDepth4 As Double

        Dim parWidth3 As Double
        Dim parDepth3 As Double

        Dim parWidth2 As Double
        Dim parDepth2 As Double

        Dim parWidth1 As Double
        Dim parDepth1 As Double

        Dim parWidth0 As Double
        Dim parDepth0 As Double

        StripWidth = 0.0508
        'Check and adjust StripWidth value to accomodate 5 baffles
        If CmpDblGreaterthan(StripWidth * 5, parWidth) Then
            StripWidth = parWidth / 5
        End If
        If CmpDblGreaterthan(StripWidth * 5, parDepth) Then
            StripWidth = parDepth / 5
        End If

        parWidth5 = dblExtendedWidth - StripWidth
        parDepth5 = dblExtendedDepth - StripWidth

        parWidth4 = parWidth5 - StripWidth
        parDepth4 = parDepth5 - StripWidth

        parWidth3 = parWidth4 - StripWidth
        parDepth3 = parDepth4 - StripWidth

        parWidth2 = parWidth3 - StripWidth
        parDepth2 = parDepth3 - StripWidth

        parWidth1 = parWidth2 - StripWidth
        parDepth1 = parDepth2 - StripWidth

        parWidth0 = parWidth1 - StripWidth
        parDepth0 = parDepth1 - StripWidth

        oPos.Set 0, 0, 0

        'Insert your code for output 1(Diffuser Plate)
        stPoint.Set -parWidth / 2, -parDepth / 2, 0
        enPoint.Set parWidth / 2, parDepth / 2, parlength

        Set ObjDiffuserPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserPlate
        Set ObjDiffuserPlate = Nothing

        'Insert your code for output 2(Diffuser Baffle -1)
        stPoint.Set 0, 0, parlength

        Set ObjDiffuserBaffle1 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parWidth, parDepth, parWidth5, parDepth5, _
                                                dblDiffuserLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle1
        Set ObjDiffuserBaffle1 = Nothing

        'Insert your code for output 3(Diffuser Stiffner -1)
        oPos.Set -dblExtendedWidth / 2, -dblExtendedDepth / 2, parlength + dblDiffuserLength

        dblPnts1(0) = oPos.x
        dblPnts1(1) = oPos.y
        dblPnts1(2) = oPos.z

        dblPnts1(3) = oPos.x
        dblPnts1(4) = oPos.y + dblExtendedDepth
        dblPnts1(5) = oPos.z

        dblPnts1(6) = oPos.x + dblExtendedWidth
        dblPnts1(7) = oPos.y + dblExtendedDepth
        dblPnts1(8) = oPos.z

        dblPnts1(9) = oPos.x + dblExtendedWidth
        dblPnts1(10) = oPos.y
        dblPnts1(11) = oPos.z

        Set ObjOuterPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dblPnts1)

        oPos.Set -parWidth5 / 2, -parDepth5 / 2, parlength + dblDiffuserLength

        dblPnts2(0) = oPos.x
        dblPnts2(1) = oPos.y
        dblPnts2(2) = oPos.z

        dblPnts2(3) = oPos.x + parWidth5
        dblPnts2(4) = oPos.y
        dblPnts2(5) = oPos.z

        dblPnts2(6) = oPos.x + parWidth5
        dblPnts2(7) = oPos.y + parDepth5
        dblPnts2(8) = oPos.z

        dblPnts2(9) = oPos.x
        dblPnts2(10) = oPos.y + parDepth5
        dblPnts2(11) = oPos.z

        Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(0), dblPnts2(1), dblPnts2(2), dblPnts2(3), dblPnts2(4), dblPnts2(5))

        Set oLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(3), dblPnts2(4), dblPnts2(5), dblPnts2(6), dblPnts2(7), dblPnts2(8))

        Set oLine3 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(6), dblPnts2(7), dblPnts2(8), dblPnts2(9), dblPnts2(10), dblPnts2(11))

        Set oLine4 = geomFactory.Lines3d.CreateBy2Points(Nothing, dblPnts2(9), dblPnts2(10), dblPnts2(11), dblPnts2(0), dblPnts2(1), dblPnts2(2))

        Set oInnerBoundary = New Collection
        oInnerBoundary.Add oLine1
        oInnerBoundary.Add oLine2
        oInnerBoundary.Add oLine3
        oInnerBoundary.Add oLine4

        Startcurve.Set dblPnts2(0), dblPnts2(1), dblPnts2(2)

        Set objHoleF = PlaceTrCString(Startcurve, oInnerBoundary)
        ObjOuterPlane.AddBoundary objHoleF

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOuterPlane

        Set ObjOuterPlane = Nothing
        Set objHoleF = Nothing
        Set oPos = Nothing

        'Insert your code for output 4(Diffuser Baffle -2)
        stPoint.Set 0, 0, parlength

        Set ObjDiffuserBaffle2 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parWidth3, parDepth3, parWidth4, parDepth4, _
                                                dblDiffuserLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle2
        Set ObjDiffuserBaffle2 = Nothing

        'Insert your code for output 5(Diffuser Baffle -3)
        stPoint.Set 0, 0, parlength

        Set ObjDiffuserBaffle3 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parWidth2, parDepth2, parWidth3, parDepth3, _
                                                dblDiffuserLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle3
        Set ObjDiffuserBaffle3 = Nothing

        'Insert your code for output 6(Diffuser Baffle -4)
        stPoint.Set 0, 0, parlength

        Set ObjDiffuserBaffle4 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parWidth1, parDepth1, parWidth2, parDepth2, _
                                                dblDiffuserLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle4
        Set ObjDiffuserBaffle4 = Nothing

        'Insert your code for output 7(Diffuser Baffle -5)
        stPoint.Set 0, 0, parlength

        Set ObjDiffuserBaffle5 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parWidth0, parDepth0, parWidth1, parDepth1, _
                                                dblDiffuserLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle5
        Set ObjDiffuserBaffle5 = Nothing

    Case 119

        'Insert your code for output 1(Diffuser Plate)
        stPoint.Set -parNeckWidth / 2, -parNeckDepth / 2, parDamperHeight
        enPoint.Set parNeckWidth / 2, parNeckDepth / 2, parDamperHeight + dHeight

        Set ObjDiffuserPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserPlate
        Set ObjDiffuserPlate = Nothing

        'Insert your code for output 2(Diffuser Baffle -1)
        stPoint.Set 0, 0, parDamperHeight + dHeight

        Set ObjDiffuserBaffle1 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, parNeckWidth, parNeckDepth, dWidth5, dDepth5, _
                                                dDiffuserHeight, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle1
        Set ObjDiffuserBaffle1 = Nothing

        'Create Damper Plate
        stPoint.Set -parNeckWidth / 2, -parNeckDepth / 2, 0
        enPoint.Set parNeckWidth / 2, parNeckDepth / 2, parDamperHeight

        Dim ObjDamperPlate As Object
        Set ObjDamperPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        m_OutputColl.AddOutput "DamperPlate", ObjDamperPlate
        Set ObjDamperPlate = Nothing

        'Create Damper Vanes
        'Assume Space between Vanes is equal to 0.0125meter.
        dSpaceBtwVanes = 0.0125
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002
        Dim dDamperHeight As Double
        dDamperHeight = 0.035
        lNumOfVanes = (parNeckWidth + dSpaceBtwVanes) / _
                      (dDamperHeight - 0.002 * Cos(PI / 4) + _
                       0.002 * Sin(PI / 4) + dSpaceBtwVanes)
        If CmpDblGreaterthan(CInt(lNumOfVanes), lNumOfVanes) Then
            lNumOfVanes = CInt(lNumOfVanes) - 1
        Else
            lNumOfVanes = CInt(lNumOfVanes)
        End If

        dLineStrPoints(0) = -parNeckWidth / 2
        dLineStrPoints(1) = -0.5 * parNeckDepth
        dLineStrPoints(2) = ((parDamperHeight - dDamperHeight) / 2) + 0.002 * Cos(PI / 4)

        dLineStrPoints(3) = -(parNeckWidth / 2 - 0.002 * Sin(PI / 4))
        dLineStrPoints(4) = -0.5 * parNeckDepth
        dLineStrPoints(5) = ((parDamperHeight - dDamperHeight) / 2)

        dLineStrPoints(6) = -((parNeckWidth / 2) - 0.002 * Sin(PI / 4) - (dDamperHeight - (0.002 * Cos(PI / 4))))
        dLineStrPoints(7) = -0.5 * parNeckDepth
        dLineStrPoints(8) = ((parDamperHeight - dDamperHeight) / 2) + (dDamperHeight - 0.002 * Cos(PI / 4))

        dLineStrPoints(9) = -(parNeckWidth / 2) + (dDamperHeight - (0.002 * Cos(PI / 4)))
        dLineStrPoints(10) = -0.5 * parNeckDepth
        dLineStrPoints(11) = ((parDamperHeight - dDamperHeight) / 2) + dDamperHeight

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)

        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 1, 0

        'Translate the linestring to origin
        Set oTransMatrix = New DT4x4
        Set oTransVector = New AutoMath.DVector
        dConst = 0.5 * (dDamperHeight - 0.002 * (Cos(PI / 4) - Sin(PI / 4)))

        oTransVector.Set (parNeckWidth / 2 - dConst), 0, -0.5 * parDamperHeight
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oTransVector
        oLineString.Transform oTransMatrix

        Set oTransVector = Nothing
        Set oTransMatrix = Nothing

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        Set oRotVector = New AutoMath.DVector
        Set oTransVector = New AutoMath.DVector

        For iCount = 0 To lNumOfVanes - 1
            Set oDamperVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parNeckDepth, True)
            dYCoord = (parNeckWidth / 2) - ((2 * iCount + 1) * dConst + (iCount * dSpaceBtwVanes))
            If (iCount Mod 2) = 0 Then
                oTransVector.Set -dYCoord, 0, 0.5 * parDamperHeight
                oTransMatrix.LoadIdentity
                oTransMatrix.Translate oTransVector
                oDamperVane.Transform oTransMatrix

                'Set the output
                m_OutputColl.AddOutput "DamperVane", oDamperVane
                Set oDamperVane = Nothing
            Else
                oRotVector.Set 0, -1, 0
                oTransMatrix.LoadIdentity
                oTransMatrix.Rotate PI / 2, oRotVector
                oDamperVane.Transform oTransMatrix

                oTransVector.Set -dYCoord, 0, 0.5 * parDamperHeight
                oTransMatrix.LoadIdentity
                oTransMatrix.Translate oTransVector
                oDamperVane.Transform oTransMatrix

                'Set the output
                m_OutputColl.AddOutput "DamperVane", oDamperVane
                Set oDamperVane = Nothing
            End If
        Next

    Case 120

        'Insert your code for output 1(Diffuser Plate)
        stPoint.Set -dWidth4 / 2, -dDepth4 / 2, parNeckHeight
        enPoint.Set dWidth4 / 2, dDepth4 / 2, parNeckHeight + dHeight

        Set ObjDiffuserPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserPlate
        Set ObjDiffuserPlate = Nothing

        'Insert your code for output 2(Diffuser Baffle -1)
        stPoint.Set 0, 0, parNeckHeight + dHeight

        Set ObjDiffuserBaffle1 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth4, dDepth4, dWidth5, dDepth5, _
                                                dDiffuserHeight, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle1
        Set ObjDiffuserBaffle1 = Nothing

        'Place cylinder
        stPoint.Set 0, 0, 0
        enPoint.Set 0, 0, parNeckHeight

        Set ObjCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parNeckWidth, True)

        'Set the output
        m_OutputColl.AddOutput "Cylinder", ObjCyl
        Set ObjCyl = Nothing

    Case 121

        'Insert your code for output 1(Diffuser Plate)
        stPoint.Set -dWidth4 / 2, -dDepth4 / 2, parDamperHeight + parNeckHeight
        enPoint.Set dWidth4 / 2, dDepth4 / 2, parDamperHeight + parNeckHeight + dHeight

        Set ObjDiffuserPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserPlate
        Set ObjDiffuserPlate = Nothing

        'Insert your code for output 2(Diffuser Baffle -1)
        stPoint.Set 0, 0, parDamperHeight + parNeckHeight + dHeight

        Set ObjDiffuserBaffle1 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth4, dDepth4, dWidth5, dDepth5, _
                                                dDiffuserHeight, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle1
        Set ObjDiffuserBaffle1 = Nothing

        'Place cylinder 1
        stPoint.Set 0, 0, 0
        enPoint.Set 0, 0, parDamperHeight

        Set ObjCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parNeckWidth, True)

        'Set the output
        m_OutputColl.AddOutput "Cylinder1", ObjCyl
        Set ObjCyl = Nothing

        'Place cylinder 2
        If CmpDblEqual(parNeckHeight, 0) Then
            stPoint.Set 0, 0, parDamperHeight
            enPoint.Set 0, 0, parDamperHeight + NEGLIGIBLE_VALUE
        Else
            stPoint.Set 0, 0, parDamperHeight
            enPoint.Set 0, 0, parDamperHeight + parNeckHeight
        End If

        Set ObjCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parNeckWidth, True)

        'Set the output
        m_OutputColl.AddOutput "Cylinder2", ObjCyl
        Set ObjCyl = Nothing

        'Create Damper Vanes
        'Assume Space between Vanes is equal to 0.0125meter.
        dSpaceBtwVanes = 0.0125
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002

        dLineStrPoints(0) = -0.4 * parNeckWidth
        dLineStrPoints(1) = -0.4 * parNeckDepth
        dLineStrPoints(2) = 0.1 * parDamperHeight + 0.002 * Cos(PI / 4)

        dLineStrPoints(3) = -(0.4 * parNeckWidth - 0.002 * Sin(PI / 4))
        dLineStrPoints(4) = -0.4 * parNeckDepth
        dLineStrPoints(5) = 0.1 * parDamperHeight

        dLineStrPoints(6) = -((0.4 * parNeckWidth) - 0.002 * Sin(PI / 4) - (0.8 * parDamperHeight - (0.002 * Cos(PI / 4))))
        dLineStrPoints(7) = -0.4 * parNeckDepth
        dLineStrPoints(8) = 0.1 * parDamperHeight + (0.8 * parDamperHeight - 0.002 * Cos(PI / 4))

        dLineStrPoints(9) = -(0.4 * parNeckWidth) + (0.8 * parDamperHeight - (0.002 * Cos(PI / 4)))
        dLineStrPoints(10) = -0.4 * parNeckDepth
        dLineStrPoints(11) = 0.9 * parDamperHeight

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)

        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 1, 0

        'Translate the linestring to origin
        Set oTransMatrix = New DT4x4
        Set oTransVector = New AutoMath.DVector
        dConst = 0.4 * (0.8 * parDamperHeight - 0.002 * (Cos(PI / 4) - Sin(PI / 4)))

        oTransVector.Set (0.4 * parNeckWidth - dConst), 0, -0.5 * parDamperHeight
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oTransVector
        oLineString.Transform oTransMatrix

        Set oTransVector = Nothing
        Set oTransMatrix = Nothing
        Set oTransMatrix = New DT4x4
        Set oTransVector = New AutoMath.DVector
        Set oDamperVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, 0.8 * parNeckDepth, True)
        oTransVector.Set 0, 0, 0.5 * parDamperHeight
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oTransVector
        oDamperVane.Transform oTransMatrix

        'Set the output
        m_OutputColl.AddOutput "DamperVane", oDamperVane
        Set oDamperVane = Nothing

    Case 122

        'Insert your code for output 1(Diffuser Plate)
        stPoint.Set -dWidth4 / 2, -dDepth4 / 2, 0
        enPoint.Set dWidth4 / 2, dDepth4 / 2, dHeight

        Set ObjDiffuserPlate = PlaceBox(m_OutputColl, stPoint, enPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserPlate
        Set ObjDiffuserPlate = Nothing

        'Insert your code for output 2(Diffuser Baffle -1)
        stPoint.Set 0, 0, dHeight

        Set ObjDiffuserBaffle1 = PlaceTrapezoid(m_OutputColl, _
                                                stPoint, dWidth4, dDepth4, dWidth5, dDepth5, _
                                                dDiffuserHeight, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffuserBaffle1
        Set ObjDiffuserBaffle1 = Nothing

    Case Else
        GoTo ErrorLabel:

    End Select
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set geomFactory = Nothing
    Set oTransVector = Nothing
    Set oRotVector = Nothing
    Set oTransMatrix = Nothing
    Set oAxisVec = Nothing
    Set oLineString = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
