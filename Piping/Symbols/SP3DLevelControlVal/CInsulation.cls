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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   Insulation.cls
'   Author:         MA
'   Creation Date:  Wednesday, Jan 30 2008
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----               ------------------
'  30.JAN.2008      MA   CR-135564  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim flangeDiam2      As Double

    Dim iOutput     As Double

    Dim parFacetoFace As Double
    Dim parConnectionHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parCoverDiameter As Double
    Dim parCoverHeight As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
'    parConnectionHeight = arrayOfInputs(3)
'    parBodyOutsideDiameter = arrayOfInputs(4)
'    parCoverDiameter = arrayOfInputs(5)
'    parCoverHeight = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, _
                                                            sptOffset, depth

    Dim dInsulationDiameter As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    Dim objInsCyl As Object
    Dim dInsCylDia As Double
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Computation of insulation diameter for insulating the horizontal cylinder.
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
        dInsulationDiameter = dInsulationDia1
    Else
        dInsulationDiameter = dInsulationDia2
    End If

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following three Part data basis values that govern its geometry,
    '   Level control valve, cylindrical type (Default case) - 740
    '   Level control valve, globe type - 741
    '   Level control valve, Y type - 742

    Select Case lPartDataBasis
        
        Case Is <= 1, 740
            parConnectionHeight = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)

            'Assumption: Body Outside Diameter is an optional Input.
            ' When not specified Body Outside Diameter = 1.6 x ConnectionHeight
            If CmpDblLessThanOrEqualTo(parBodyOutsideDiameter, 0) Then _
                                    parBodyOutsideDiameter = 1.6 * parConnectionHeight
                                    
            'Place Cylinder 1
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing

            'Insert your code for middle portion of the body of the valve
            Dim oPoints(1 To 8) As New AutoMath.DPosition
            Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
            Dim oAxisVec As AutoMath.DVector
            Dim oCenPoint As AutoMath.DPosition
            Dim objInsMiddlePortion As Object

            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set -0.3 * parFacetoFace, 0.5 * pipeDiam + parInsulationThickness, 0
            oPoints(2).Set -0.27 * parFacetoFace, 0.43 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(3).Set -0.25 * parFacetoFace, 0.45 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(4).Set -0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(5).Set 0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(6).Set 0.25 * parFacetoFace, 0.45 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(7).Set 0.27 * parFacetoFace, 0.43 * parBodyOutsideDiameter + parInsulationThickness, 0
            oPoints(8).Set 0.3 * parFacetoFace, 0.5 * pipeDiam2 + parInsulationThickness, 0

            Set objBsplineCurve = PlaceTrBspline(7, oPoints)

            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0

            Set objInsMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsMiddlePortion
            Set objInsMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing

            'Place Box
            oStPoint.Set -0.2 * parFacetoFace - parInsulationThickness, _
                                0.45 * parBodyOutsideDiameter, _
                                0.1 * parBodyOutsideDiameter + parInsulationThickness
            oEnPoint.Set -0.1 * parFacetoFace + parInsulationThickness, _
                            0.45 * parBodyOutsideDiameter + 0.15 * parConnectionHeight _
                                + parInsulationThickness, _
                                -0.1 * parBodyOutsideDiameter - parInsulationThickness

            Dim objInsBox As Object
            Set objInsBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBox
            Set objInsBox = Nothing

            'Place Cylinder 2
            oStPoint.Set -0.15 * parFacetoFace, 0.45 * parBodyOutsideDiameter _
                                            + 0.15 * parConnectionHeight, 0
            oEnPoint.Set -0.15 * parFacetoFace, 0.5 * parBodyOutsideDiameter _
                                + 0.2 * parConnectionHeight + parInsulationThickness, 0
            dInsCylDia = 0.1 * parBodyOutsideDiameter + 2 * parInsulationThickness

            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsCylDia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing

        Case 741
            parConnectionHeight = arrayOfInputs(3)
            parCoverDiameter = arrayOfInputs(5)
            
            'Place Cylinder 1
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.5 * parFacetoFace, 0, 0
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
                                   
            'Place Ellipsoid
            Dim dEllipseMajorX   As Double
            Dim dEllipseMajorY   As Double
            Dim dEllipseMajorZ   As Double
            Dim dEllipseMMRatio  As Double
            Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
            Dim objEllipsoid As Object
            
            dEllipseMMRatio = (0.3 * parConnectionHeight + parInsulationThickness) / (0.4 * parFacetoFace + parInsulationThickness)
            dEllipseMajorX = 0.4 * parFacetoFace + parInsulationThickness
            dEllipseMajorY = 0
            dEllipseMajorZ = 0
            
            Set objEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                    Nothing, 0, 0, 0, _
                                                    0, 0, 1, dEllipseMajorX, dEllipseMajorY, dEllipseMajorZ, _
                                                    dEllipseMMRatio, 0, PI)
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
          
            Set objEllipsoid = PlaceRevolution(m_OutputColl, objEllipticalArc, oAxisVec, _
                                            oCenPoint, 2 * PI, True)
          
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllipsoid
            Set objEllipsoid = Nothing
        
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Dim sp3dElem1 As IJDObject
            Set sp3dElem1 = objEllipticalArc
            Set objEllipticalArc = Nothing
            sp3dElem1.Remove
            Set sp3dElem1 = Nothing
            
            'Place Cylinder 2
            oStPoint.Set 0, 0.28 * parConnectionHeight, 0
            oEnPoint.Set 0, 0.4 * parConnectionHeight, 0
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                0.4 * parCoverDiameter + 2 * parInsulationThickness, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
            'Place Cylinder 3
            oStPoint.Set 0, 0.4 * parConnectionHeight, 0
            oEnPoint.Set 0, 0.6 * parConnectionHeight + parInsulationThickness, 0
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parCoverDiameter + 2 * parInsulationThickness, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing
            
        Case 742
            parCoverDiameter = arrayOfInputs(5)
            parCoverHeight = arrayOfInputs(6)
                       
            'Assumption: Cover Diameter is an optional Input.
            ' When not specified Cover Diameter = 0.75 x Face to Face, and
            If CmpDblLessThanOrEqualTo(parCoverDiameter, 0) Then _
                                    parCoverDiameter = 0.75 * parFacetoFace
                                    
            'These values are used in updating the valve Operator.
            parHandwheelAngle = PI / 2.4 'Assumption: the Operator is at inclined at 75Deg to x Axis.
            Dim dInclinationAngle As Double
            dInclinationAngle = parHandwheelAngle
            
            Dim dCylDia As Double
            'The Ellipsoid shoud not engulf the upper cover on which operator appears,
            'the diameter of the stem reaching upto the cover should be less than cover width,
            'the base ellipsoid (created below) should not exceed face to face dimension and
            'the section of horizontal cone that protrudes into the inclined ellipsoid is less than Pipe dia.
            dCylDia = ReturnMin(parFacetoFace / 4, pipeDiam, parCoverHeight, parCoverDiameter) _
                                                            + 2 * parInsulationThickness
            
            'Create the insualted inclined Ellipsoid
            Dim dHalfMajorAxisInclinedEllipArc As Double
            dHalfMajorAxisInclinedEllipArc = dCylDia * 1.1
            If CmpDblGreaterthan(dHalfMajorAxisInclinedEllipArc, parFacetoFace / 2) Then _
                dHalfMajorAxisInclinedEllipArc = parFacetoFace / 2

            Dim dHalfMinorAxisInclinedEllipArc As Double
            dHalfMinorAxisInclinedEllipArc = dCylDia * 0.85
            
            If CmpDblGreaterthan(dHalfMinorAxisInclinedEllipArc, parFacetoFace / 2) Then _
                dHalfMinorAxisInclinedEllipArc = parFacetoFace / 2
                            
            Dim objInclinedEllipticalArc As Object
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set objInclinedEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                0, 0, 0, _
                0, 0, 1, _
                dHalfMinorAxisInclinedEllipArc * Cos(dInclinationAngle), dHalfMinorAxisInclinedEllipArc * Sin(dInclinationAngle), 0, _
                dHalfMajorAxisInclinedEllipArc / dHalfMinorAxisInclinedEllipArc, _
                                    0, PI)

            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set Cos(dInclinationAngle), Sin(dInclinationAngle), 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0

            Dim objInclinedEllipsoid As Object
            Set objInclinedEllipsoid = PlaceRevolution(m_OutputColl, objInclinedEllipticalArc, _
                                oAxisVec, oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInclinedEllipsoid
            Set objInclinedEllipsoid = Nothing
            Set objInclinedEllipticalArc = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            Set oGeomFactory = Nothing

            'Create insulated horizontal Cylinder
            oStPoint.Set parFacetoFace / 2, 0, 0
            oEnPoint.Set -parFacetoFace / 2, 0, 0
            
            Set objInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dInsulationDiameter, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCyl
            Set objInsCyl = Nothing

            'Create insulated Cover
            'The height of the cover is measured from the center-line to the top-most edge of the inclined cylinder.
            'Effective height denotes the height of the center of the upper cross section of the cover.
            Dim dEffectiveHeightofCover As Double
            dEffectiveHeightofCover = (parCoverHeight - 0.5 * pipeDiam) - parCoverDiameter * 0.5 * Cos(dInclinationAngle)
            
            oStPoint.Set dEffectiveHeightofCover / Tan(dInclinationAngle), _
                    dEffectiveHeightofCover, 0
            oEnPoint.Set oStPoint.x - 0.15 * parCoverHeight / Tan(dInclinationAngle), _
                    oStPoint.y - 0.15 * parCoverHeight, 0
            
            Dim dInsulationDiaCover As Double
            Dim dInsulationDiaStem As Double
            'Limit the insualtion diameter if it extends to intersect the flange.
            If CmpDblGreaterthan(oEnPoint.y - parCoverDiameter * 0.5 * Cos(dInclinationAngle) _
                    - parInsulationThickness, ReturnMin(flangeDiam / 2, pipeDiam / 2)) Then
                dInsulationDiaCover = parCoverDiameter + 2 * parInsulationThickness
                dInsulationDiaStem = dCylDia    'dCylDia is inclusive of insulation Thickness.
            Else
                dInsulationDiaCover = parCoverDiameter + 0.001
                dInsulationDiaStem = dCylDia - 2 * parInsulationThickness + 0.001
            End If
            
            Dim objInsCover As Object
            Set objInsCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        dInsulationDiaCover, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCover
            Set objInsCover = Nothing
            
            'Create insulated stem reaching upto the cover
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set 0, 0, 0
            
            Dim objStem As Object
            Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                   dInsulationDiaStem, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
            Set objStem = Nothing

        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
    
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Function ReturnMin(A#, B#, Optional C# = -1, Optional D# = -1) As Double
    'Returns the minimum value among four input values.
    Const METHOD = "ReturnMin"
    On Error GoTo ErrorLabel
    
    Dim dMinValue As Double
    dMinValue = A
    If CmpDblLessThan(B, dMinValue) Then dMinValue = B
    
    If Not CmpDblLessThanOrEqualTo(C, 0) Then
        If CmpDblLessThan(C, dMinValue) Then dMinValue = C
    End If
    
    If Not CmpDblLessThanOrEqualTo(D, 0) Then
        If CmpDblLessThan(D, dMinValue) Then dMinValue = D
    End If
    
    ReturnMin = dMinValue
    
    Exit Function
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
End Function

