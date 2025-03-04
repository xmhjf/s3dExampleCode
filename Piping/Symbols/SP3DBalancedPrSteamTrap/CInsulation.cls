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
'   Author:         MA
'   Creation Date:  Wednesday, Sep 12 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
'''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim objInsulatedBody As Object
    Dim objInsBox As Object
    
    Dim parCenterlinetoCenterline As Double
    Dim parPipingSpecialtyHeight   As Double
    Dim parPipingSpecialtyDiameter As Double
    Dim parOutletCenterlineHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parFacetoCenter As Double
    Dim parCenterlineHeight As Double
    Dim parFacetoFace   As Double
    Dim parPipingSpecialtyDepth As Double
    Dim parPipingSpecialtyLength As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parInletCenterlineHeight As Double
    Dim parInletWidth As Double
    Dim parInlettoOutletCenterlineWidth As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double

    Dim LineStrPoints(0 To 26)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objInsCyl1 As Object
    Dim objInsCyl2 As Object
    Dim dInsCylDia As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oArcPoint As AutoMath.DPosition
    Dim oLine1   As IngrGeom3D.Line3d
    Dim oLine2   As IngrGeom3D.Line3d
    Dim oArc1  As IngrGeom3D.Arc3d
    Dim oCurve         As IngrGeom3D.ComplexString3d
    Dim oCurveCol      As Collection

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parCenterlinetoCenterline = arrayOfInputs(2)
'    parPipingSpecialtyHeight = arrayOfInputs(3)
'    parPipingSpecialtyDiameter = arrayOfInputs(4)
'    parOutletCenterlineHeight = arrayOfInputs(5)
'    parBodyOutsideDiameter = arrayOfInputs(6)
'    parFacetoCenter = arrayOfInputs(7)
'    parCenterlineHeight = arrayOfInputs(8)
'    parFacetoFace = arrayOfInputs(9)
'    parPipingSpecialtyDepth = arrayOfInputs(10)
'    parPipingSpecialtyLength = arrayOfInputs(11)
'    parPipingSpecialtyWidth = arrayOfInputs(12)
'    parInletCenterlineHeight = arrayOfInputs(13)
'    parInletWidth = arrayOfInputs(14)
'    parInlettoOutletCenterlineWidth = arrayOfInputs(15)
'    parOffset = arrayOfInputs(16)
    parInsulationThickness = arrayOfInputs(17)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                     sptOffset, depth

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2

 ' Insert your code for output 2(Insulated Body)

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Select Case lPartDataBasis

        Case Is <= 1, 405

            parCenterlinetoCenterline = arrayOfInputs(2)
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parOutletCenterlineHeight = arrayOfInputs(5)
            parBodyOutsideDiameter = arrayOfInputs(6)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then _
                        parBodyOutsideDiameter = 0.55 * parPipingSpecialtyDiameter

            'Point 1
            LineStrPoints(0) = -parOutletCenterlineHeight / 2
            LineStrPoints(1) = 0
            LineStrPoints(2) = 0

            'Point 2
            LineStrPoints(3) = LineStrPoints(0)
            LineStrPoints(4) = parBodyOutsideDiameter / 2 + parInsulationThickness
            LineStrPoints(5) = LineStrPoints(2)

            'Point 3
            LineStrPoints(6) = 0.75 * (parPipingSpecialtyHeight - parOutletCenterlineHeight) _
                                - parInsulationThickness
            LineStrPoints(7) = LineStrPoints(4)
            LineStrPoints(8) = LineStrPoints(2)

            'Point 4
            LineStrPoints(9) = LineStrPoints(6)
            LineStrPoints(10) = parPipingSpecialtyDiameter / 2 + parInsulationThickness
            LineStrPoints(11) = LineStrPoints(2)

            'Point 5
            LineStrPoints(12) = 0.9 * (parPipingSpecialtyHeight - parOutletCenterlineHeight) _
                                    + parInsulationThickness
            LineStrPoints(13) = LineStrPoints(10)
            LineStrPoints(14) = LineStrPoints(2)

            'Point 6
            LineStrPoints(15) = LineStrPoints(12)
            LineStrPoints(16) = parPipingSpecialtyDiameter / 4 + parInsulationThickness
            LineStrPoints(17) = LineStrPoints(2)

            'Point 7
            LineStrPoints(18) = (parPipingSpecialtyHeight - parOutletCenterlineHeight) _
                                    + parInsulationThickness
            LineStrPoints(19) = LineStrPoints(16)
            LineStrPoints(20) = LineStrPoints(2)
            
            'Point 8
            LineStrPoints(21) = LineStrPoints(18)
            LineStrPoints(22) = 0
            LineStrPoints(23) = LineStrPoints(2)
            
            'Point 9
            LineStrPoints(24) = LineStrPoints(0)
            LineStrPoints(25) = LineStrPoints(1)
            LineStrPoints(26) = LineStrPoints(2)

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    9, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
         
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parOutletCenterlineHeight, 0, 0
            oEnPoint.Set -parOutletCenterlineHeight, 0, 0
            
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
          
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -parCenterlinetoCenterline / 2, 0
            oEnPoint.Set 0, -parCenterlinetoCenterline, 0
           
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam2 + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam2 + 2 * parInsulationThickness
            End If
            
            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
          
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

        Case 406

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parBodyOutsideDiameter = arrayOfInputs(6)
            parFacetoCenter = arrayOfInputs(7)
            parCenterlineHeight = arrayOfInputs(8)

            If CmpDblEqual(parBodyOutsideDiameter, 0) Then _
                            parBodyOutsideDiameter = 0.6 * parPipingSpecialtyDiameter

            Dim LineStrPoints1(0 To 32)  As Double
            'Point 1
            LineStrPoints1(0) = 0
            LineStrPoints1(1) = (parPipingSpecialtyHeight - parCenterlineHeight) _
                                    + parInsulationThickness
            LineStrPoints1(2) = 0

            'Point 2
            LineStrPoints1(3) = -parPipingSpecialtyDiameter / 4 - parInsulationThickness
            LineStrPoints1(4) = LineStrPoints1(1)
            LineStrPoints1(5) = LineStrPoints1(2)

            'Point 3
            LineStrPoints1(6) = LineStrPoints1(3)
            LineStrPoints1(7) = 0.9 * (parPipingSpecialtyHeight - parCenterlineHeight) _
                                    + parInsulationThickness
            LineStrPoints1(8) = LineStrPoints1(2)

            'Point 4
            LineStrPoints1(9) = -parPipingSpecialtyDiameter / 2 - parInsulationThickness
            LineStrPoints1(10) = LineStrPoints1(7)
            LineStrPoints1(11) = LineStrPoints1(2)

            'Point 5
            LineStrPoints1(12) = LineStrPoints1(9)
            LineStrPoints1(13) = 0.75 * (parPipingSpecialtyHeight - parCenterlineHeight) _
                                    - parInsulationThickness
            LineStrPoints1(14) = LineStrPoints1(2)

            'Point 6
            LineStrPoints1(15) = -parBodyOutsideDiameter / 2 - parInsulationThickness
            LineStrPoints1(16) = LineStrPoints1(13)
            LineStrPoints1(17) = LineStrPoints1(2)

            'Point 7
            LineStrPoints1(18) = LineStrPoints1(15)
            LineStrPoints1(19) = -0.7 * parCenterlineHeight - parInsulationThickness
            LineStrPoints1(20) = LineStrPoints1(2)

            'Point 8
            LineStrPoints1(21) = -parPipingSpecialtyDiameter / 8 - parInsulationThickness
            LineStrPoints1(22) = LineStrPoints1(19)
            LineStrPoints1(23) = LineStrPoints1(2)

            'Point 9
            LineStrPoints1(24) = LineStrPoints1(21)
            LineStrPoints1(25) = -parCenterlineHeight - parInsulationThickness
            LineStrPoints1(26) = LineStrPoints1(2)

            'Point 10
            LineStrPoints1(27) = LineStrPoints1(0)
            LineStrPoints1(28) = LineStrPoints1(25)
            LineStrPoints1(29) = LineStrPoints1(2)

            'Point 11
            LineStrPoints1(30) = LineStrPoints1(0)
            LineStrPoints1(31) = LineStrPoints1(1)
            LineStrPoints1(32) = LineStrPoints1(2)
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    11, LineStrPoints1)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)

            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parBodyOutsideDiameter / 2, 0, 0
            oEnPoint.Set -parFacetoCenter, 0, 0
            
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
            
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set parBodyOutsideDiameter / 2, 0, 0
            oEnPoint.Set parFacetoCenter, 0, 0
            
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam2 + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam2 + 2 * parInsulationThickness
            End If
            
            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

        Case 407

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parPipingSpecialtyDepth = arrayOfInputs(10)
            parPipingSpecialtyLength = arrayOfInputs(11)

            'Place Cylinder
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dInsCylDia = dInsulationDia1
            Else
                dInsCylDia = dInsulationDia2
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Box1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.45 * parPipingSpecialtyLength - parInsulationThickness, _
                            0.5 * parPipingSpecialtyHeight + parInsulationThickness, _
                            0.5 * parPipingSpecialtyDiameter + parInsulationThickness
            oEnPoint.Set 0.45 * parPipingSpecialtyLength + parInsulationThickness, _
                            0.2 * parPipingSpecialtyHeight + parInsulationThickness, _
                            -0.5 * parPipingSpecialtyDiameter - parInsulationThickness
            
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
         
            'Set the output
            m_OutputColl.AddOutput "BoxIns", objInsBox
            Set objInsBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            Dim LineStrPoints2(0 To 14)  As Double
            'Point 1
            LineStrPoints2(0) = 0
            LineStrPoints2(1) = 0.5 * parPipingSpecialtyHeight
            LineStrPoints2(2) = 0

            'Point 2
            LineStrPoints2(3) = -0.5 * parPipingSpecialtyDiameter - parInsulationThickness
            LineStrPoints2(4) = LineStrPoints2(1)
            LineStrPoints2(5) = LineStrPoints2(2)

            'Point 3
            LineStrPoints2(6) = -0.3 * parPipingSpecialtyDiameter - parInsulationThickness
            LineStrPoints2(7) = parPipingSpecialtyHeight + parInsulationThickness
            LineStrPoints2(8) = LineStrPoints2(2)

            'Point 4
            LineStrPoints2(9) = 0
            LineStrPoints2(10) = LineStrPoints2(7)
            LineStrPoints2(11) = LineStrPoints2(2)

            'Point 5
            LineStrPoints2(12) = LineStrPoints2(0)
            LineStrPoints2(13) = LineStrPoints2(1)
            LineStrPoints2(14) = LineStrPoints2(2)

            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    5, LineStrPoints2)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.00001, 0, 0
            PI = 4 * Atn(1)

            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Box2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parPipingSpecialtyLength - parInsulationThickness, _
                            -0.2 * parPipingSpecialtyHeight - parInsulationThickness, _
                            0.1 * parPipingSpecialtyHeight + parInsulationThickness
            oEnPoint.Set 0.5 * parPipingSpecialtyLength + parInsulationThickness, _
                            -parPipingSpecialtyDepth - parInsulationThickness, _
                            -0.1 * parPipingSpecialtyHeight - parInsulationThickness
            
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       
            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", objInsBox
            Set objInsBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        Case 408

            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dInsCylDia = dInsulationDia1
            Else
                dInsCylDia = dInsulationDia2
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

        Case 409

            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parPipingSpecialtyWidth = arrayOfInputs(12)

            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
            
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dInsCylDia = dInsulationDia1
            Else
                dInsCylDia = dInsulationDia2
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0, _
                        -0.45 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            oEnPoint.Set 0, 0, _
                        -0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter) _
                             - parInsulationThickness
            dInsCylDia = 0.8 * parFacetoFace + 2 * parInsulationThickness
           
            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Cylinder 3
            Dim objInsCyl3 As Object
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0, _
                        -0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter)
            oEnPoint.Set 0, 0, _
                        -(parPipingSpecialtyWidth + 0.5 * parPipingSpecialtyDiameter) _
                            - parInsulationThickness
            dInsCylDia = 0.8 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter) + _
                                    2 * parInsulationThickness
            Set objInsCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl3
            Set objInsCyl3 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, _
                    0.4 * (parPipingSpecialtyWidth - 0.5 * parPipingSpecialtyDiameter), _
                            -parPipingSpecialtyWidth
            oEnPoint.Set 0, parPipingSpecialtyHeight + parInsulationThickness, _
                                -parPipingSpecialtyWidth
            dInsCylDia = parPipingSpecialtyDiameter + 2 * parInsulationThickness

            Set objInsCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl3
            Set objInsCyl3 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
        Case 410
            
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parInletCenterlineHeight = arrayOfInputs(13)
            parInletWidth = arrayOfInputs(14)
            parInlettoOutletCenterlineWidth = arrayOfInputs(15)
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth, 0, 0
            oEnPoint.Set -parInlettoOutletCenterlineWidth, 0, 0
            
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam + 2 * parInsulationThickness
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 2
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set 0, -parInletCenterlineHeight, 0
           
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                dInsCylDia = pipeDiam2 + 2 * parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                dInsCylDia = flangeDiam2 + 2 * parInsulationThickness
            End If
            
            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
    
            'Place Revolution
            'Line1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set -0.6 * parInlettoOutletCenterlineWidth - parInsulationThickness, _
                                        -0.6 * parInletCenterlineHeight, 0
            
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
            'Line2
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth - parInsulationThickness, _
                                        -0.6 * parInletCenterlineHeight, 0
            oEnPoint.Set -0.6 * parInlettoOutletCenterlineWidth - parInsulationThickness, _
                            (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            
            Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
          
            'Arc1
            oStPoint.Set -0.6 * parInlettoOutletCenterlineWidth - parInsulationThickness, _
                            (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            oEnPoint.Set 0, (parPipingSpecialtyHeight - parInletCenterlineHeight), 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, (parPipingSpecialtyHeight - parInletCenterlineHeight) / 2, 0
            
            Set oArc1 = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            Set oCurveCol = New Collection
            oCurveCol.Add oLine1
            oCurveCol.Add oLine2
            oCurveCol.Add oArc1
            oStPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            
            Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 0, 1, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.6 * parInletCenterlineHeight, 0
            PI = 4 * Atn(1)
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oCurve, oAxisVect, oCenPoint, 2 * PI, True)
         
            'Set the output
            m_OutputColl.AddOutput "InsulatedBody", objInsulatedBody
            Set objInsulatedBody = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLine1 = Nothing
            Set oLine2 = Nothing
            Set oArc1 = Nothing
            Set oCurve = Nothing
            Set oCurveCol = Nothing
    
        Case 411
            
            parPipingSpecialtyHeight = arrayOfInputs(3)
            parPipingSpecialtyDiameter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(9)
            parOffset = arrayOfInputs(16)
            
            'Place Cylinder 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
           
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDia1 = flangeDiam + parInsulationThickness * 2
            Else
                dInsulationDia1 = pipeDiam + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
            Else
                dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
            End If

            If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
                dInsCylDia = dInsulationDia1
            Else
                dInsCylDia = dInsulationDia2
            End If
            
            Set objInsCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl1
            Set objInsCyl1 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Box 1
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFacetoFace - parInsulationThickness, _
                                        0.2 * parPipingSpecialtyHeight, _
                                        0.4 * parPipingSpecialtyHeight + parInsulationThickness
            oEnPoint.Set 0.3 * parFacetoFace + parInsulationThickness, _
                                        0.5 * parPipingSpecialtyHeight + parInsulationThickness, _
                                        -0.4 * parPipingSpecialtyHeight - parInsulationThickness
           
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
          
            'Set the output
            m_OutputColl.AddOutput "BoxIns", objInsBox
            Set objInsBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 3
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0.5 * parPipingSpecialtyHeight, 0
            oEnPoint.Set 0, parPipingSpecialtyHeight + parInsulationThickness, 0
            dInsCylDia = 0.8 * parPipingSpecialtyDiameter + 2 * parInsulationThickness
           
            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)
           
            'Set the output
            m_OutputColl.AddOutput "CylIns", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
            'Place Cylinder 4
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.5 * parPipingSpecialtyDiameter * Cos(PI / 6) - parInsulationThickness, _
                                 -parOffset * Sin(PI / 6), 0
            oEnPoint.Set 0.4 * parPipingSpecialtyDiameter * Cos(PI / 6), _
                                -0.1 * parPipingSpecialtyHeight * Sin(PI / 6), 0
            dInsCylDia = 0.4 * parPipingSpecialtyHeight + 2 * parInsulationThickness

            Set objInsCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            m_OutputColl.AddOutput "InsCyl2", objInsCyl2
            Set objInsCyl2 = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            
    Case Else
        GoTo ErrorLabel:

    End Select

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
