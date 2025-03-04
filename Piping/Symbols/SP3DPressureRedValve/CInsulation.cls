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
'   Insulation.cls
'   Author:         dkl
'   Creation Date:  Thursday, May 17 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol is designed to cater to requirements of JIS B8410 represented by part data basis
'    Pressure Reducing Valve with integral strainer, apart from considering other cases.
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----               ------------------
'  17.MAY.2007      dkl   CR-118766  Created the symbol.
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
    Dim parHeightCover As Double
    Dim parHeightLowerCover As Double
    Dim parHeightUpperCover As Double
    Dim parWidthCover As Double
    Dim parWidthLowerCover As Double
    Dim parWidthUpperCover As Double
    Dim parHeightofStrainer As Double
    Dim parDiaofStrainer As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
'    parHeightCover = arrayOfInputs(3)  'Height (optional parameter)
'    parHeightLowerCover = arrayOfInputs(4)  'Height1 (optional parameter)
'    parHeightUpperCover = arrayOfInputs(5)  'Height2 (optional parameter)
'    parWidthCover = arrayOfInputs(6)  'Width (optional parameter)
'    parWidthLowerCover = arrayOfInputs(7) 'Width1 (optional parameter)
'    parWidthUpperCover = arrayOfInputs(8)  'Width2 (optional parameter)
'    parHeightofStrainer = arrayOfInputs(9)
'    parDiaofStrainer = arrayOfInputs(10)    (optional parameter)
    parHandwheelAngle = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, _
                                                            sptOffset, depth
    
    Dim dInsulationDiameter As Double
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double
    
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
    
'The center is assumed at Face to Face / 2.
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

' This symbol is based on the following three Part data basis values that govern its geometry,
' Pressure reducing valve, Type 1 (Default case) - 315
' Pressure reducing valve, Type 2 - 320
' Pressure reducing valve, Type 3 - 325
' Pressure reducing valve, Y-type - 330
' Pressure reducing valve with integral strainer - 335

    Select Case lPartDataBasis
        
        Case Is <= 1, 315 'The default state, Pressure reducing valve, Type 1
            parHeightLowerCover = arrayOfInputs(4)  'Height1 (optional parameter)
            parHeightUpperCover = arrayOfInputs(5)  'Height2 (optional parameter)
            parWidthLowerCover = arrayOfInputs(7)  'Width1 (optional parameter)
            parWidthUpperCover = arrayOfInputs(8)  'Width2 (optional parameter)
                       
            'Assumption: Width of Lower Cover and Width of Upper Cover are optional Inputs.
            ' When not specified Width of Lower Cover = 0.4 x Face to Face, and
            '                    Width of Upper Cover = 0.4 x Face to Face.
            If CmpDblLessThanOrEqualTo(parWidthLowerCover, 0) Then _
                                            parWidthLowerCover = 0.4 * parFacetoFace
         
            If CmpDblLessThanOrEqualTo(parWidthUpperCover, 0) Then _
                                            parWidthUpperCover = 0.4 * parFacetoFace
                                            
            'Create Insulated Horizontal cylinder
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Dim objInsHorizontalCylinder As Object
            Set objInsHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            dInsulationDiameter, True)
                        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHorizontalCylinder
            Set objInsHorizontalCylinder = Nothing
            
            'Create insulated Vertical Cylinder
            ' Assumption: Diameter of Vertical Cylinder is 40% Face to Face.
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, -0.85 * parHeightLowerCover, 0
            
            Dim dDiaVerticalCylinder As Double
            dDiaVerticalCylinder = ReturnMin(parWidthUpperCover * 0.7 + 2 * parInsulationThickness, _
                            parWidthLowerCover * 0.7 + 2 * parInsulationThickness, _
                            parFacetoFace)
                       
            Dim objInsVerticalCylinder As Object
            Set objInsVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dDiaVerticalCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsVerticalCylinder
            Set objInsVerticalCylinder = Nothing
            
            'Create insulated Upper Cover
            'Assumption: The Thickness of Upper Cover is equal to 15% of its height.
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, parHeightUpperCover, 0
            
            Dim dInsulatedVerticalCylinderDia As Double
            dInsulatedVerticalCylinderDia = parWidthUpperCover + 2 * parInsulationThickness
            
            Dim objInsUpperCover As Object
            Set objInsUpperCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dInsulatedVerticalCylinderDia, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsUpperCover
            Set objInsUpperCover = Nothing
            
            'Create Lower Cover
            'Assumption: The Thickness of Lower Cover is equal to 15% of its height.
            oStPoint.Set 0, -0.85 * parHeightLowerCover, 0
            oEnPoint.Set 0, -parHeightLowerCover - parInsulationThickness, 0
            
            dInsulatedVerticalCylinderDia = parWidthLowerCover + 2 * parInsulationThickness
                
            Dim objInsLowerCover As Object
            Set objInsLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dInsulatedVerticalCylinderDia, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLowerCover
            Set objInsLowerCover = Nothing
            
        Case 320 'Pressure reducing valve, Type 2.
            parHeightLowerCover = arrayOfInputs(4)  'Height1 (optional parameter)
            parHeightUpperCover = arrayOfInputs(5)  'Height2 (optional parameter)
            parWidthLowerCover = arrayOfInputs(7)  'Width1 (optional parameter)
            parWidthUpperCover = arrayOfInputs(8)  'Width2 (optional parameter)
            
            'Assumption: Width of Lower Cover and Width of Upper Cover are optional Inputs.
            ' When not specified Width of Lower Cover = 0.15 x Face to Face, and
            '                    Width of Upper Cover = 0.4 x Face to Face.
            If CmpDblLessThanOrEqualTo(parWidthLowerCover, 0) Then _
                                            parWidthLowerCover = 0.15 * parFacetoFace
            If CmpDblLessThanOrEqualTo(parWidthUpperCover, 0) Then _
                                            parWidthUpperCover = 0.4 * parFacetoFace
                                            
            'Create Insulated Horizontal cylinder
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Set objInsHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            dInsulationDiameter, True)
                                            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHorizontalCylinder
            Set objInsHorizontalCylinder = Nothing
            
            'Create Insulated Upper Cover
            Dim dInsulatedCoverHalfWidth As Double
            dInsulatedCoverHalfWidth = parWidthUpperCover / 2 + parInsulationThickness
            
            oStPoint.Set -dInsulatedCoverHalfWidth, 0.85 * parHeightUpperCover, -dInsulatedCoverHalfWidth
            oEnPoint.Set dInsulatedCoverHalfWidth, parHeightUpperCover, dInsulatedCoverHalfWidth
            
            Set objInsUpperCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsUpperCover
            Set objInsUpperCover = Nothing
            
            'Create insulated Vertical Cylinder
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, -0.8 * parHeightLowerCover, 0  'The rest 20% height is the cover
            
            dDiaVerticalCylinder = ReturnMin(parWidthLowerCover * 1.2 + 2 * parInsulationThickness, _
                                parWidthUpperCover * 0.95 + 2 * parInsulationThickness, _
                                parFacetoFace)
            
            Dim objInsLowerVerticalCylinder As Object
            Set objInsLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                     dDiaVerticalCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLowerVerticalCylinder
            Set objInsLowerVerticalCylinder = Nothing
            
            'Insert your code for insulated Lower Cover
            'Assumption: The Thickness of Lower Cover is equal to 20% of its height.
            oStPoint.Set 0, -0.8 * parHeightLowerCover, 0
            oEnPoint.Set 0, -parHeightLowerCover - parInsulationThickness, 0
            
            dInsulatedVerticalCylinderDia = parWidthLowerCover + 2 * parInsulationThickness
            
            Set objInsLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                dInsulatedVerticalCylinderDia, True)
                                
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLowerCover
            Set objInsLowerCover = Nothing
            
        Case 325 'Pressure reducing valve, Type 3.
            parHeightLowerCover = arrayOfInputs(4)  'Height1 (optional parameter)
            parHeightUpperCover = arrayOfInputs(5)  'Height2 (optional parameter)
            parWidthLowerCover = arrayOfInputs(7)  'Width1 (optional parameter)
            parWidthUpperCover = arrayOfInputs(8)  'Width2 (optional parameter)
            
            'Assumption: Width of Lower Cover and Width of Upper Cover are optional Inputs.
            ' When not specified Width of Lower Cover = 0.4 x Face to Face, and
            '                    Width of Upper Cover = 0.4 x Face to Face.
            If CmpDblLessThanOrEqualTo(parWidthLowerCover, 0) Then _
                                            parWidthLowerCover = 0.4 * parFacetoFace
            If CmpDblLessThanOrEqualTo(parWidthUpperCover, 0) Then _
                                            parWidthUpperCover = 0.4 * parFacetoFace
    
            'Create Insulated Horizontal cylinder
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Set objInsHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            dInsulationDiameter, True)
                        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHorizontalCylinder
            Set objInsHorizontalCylinder = Nothing
            
            'Create Vertical Ellipsoid.
            ' Assumtion: From the origin, the Center of the Elliptical arc is at a distance of 40%
            ' of the lesser of the upper and lower cover widths. However it should not exceed
            '20% of Face to Face dimension.
            Dim dDistCenterofEllipArctoOrigin As Double
            dDistCenterofEllipArctoOrigin = ReturnMin(parWidthLowerCover * 0.4, _
                                parWidthUpperCover * 0.4, parFacetoFace * 0.2)
            
            ' Assumtion: From the origin, the Major axis of the Elliptical arc is at a distance
            ' of 80% of the lesser of the upper and lower cover Heights.
            Dim dHalfMajorAxisEllipticalArc As Double
            dHalfMajorAxisEllipticalArc = 0.8 * ReturnMin(parHeightLowerCover, _
                                                                    parHeightUpperCover)
            
            'Assumption: The Half-Width of the Ellipsoid is equal to 35% of Face to Face.
            Dim dHalfMinorAxisEllipticalArc As Double
            dHalfMinorAxisEllipticalArc = parFacetoFace * 0.35 - dDistCenterofEllipArctoOrigin _
                                                     + parInsulationThickness
            
            If CmpDblGreaterthan(dHalfMinorAxisEllipticalArc, _
                        parFacetoFace / 2 - flangeThick - dDistCenterofEllipArctoOrigin) Then _
                dHalfMinorAxisEllipticalArc = parFacetoFace / 2 - -flangeThick - dDistCenterofEllipArctoOrigin
            
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
            Dim objVerticalEllipticalArc As Object
            Set objVerticalEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                -dDistCenterofEllipArctoOrigin, 0, 0, _
                                0, 0, 1, _
                                0, dHalfMajorAxisEllipticalArc, 0, _
                                dHalfMinorAxisEllipticalArc / dHalfMajorAxisEllipticalArc, 0, PI)
            
            Dim oAxisVec As AutoMath.DVector
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 0, 1, 0
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0.0001, 0, 0
            
            Dim objEllipsoid As Object
            Set objEllipsoid = PlaceRevolution(m_OutputColl, objVerticalEllipticalArc, oAxisVec, _
                                            oCenPoint, 2 * PI, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllipsoid
            Set objEllipsoid = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            Set objVerticalEllipticalArc = Nothing
            Set oGeomFactory = Nothing
            
            'Create Upper Cover
            dInsulatedCoverHalfWidth = parWidthUpperCover / 2 + parInsulationThickness
            
            oStPoint.Set -dInsulatedCoverHalfWidth, dHalfMajorAxisEllipticalArc, -dInsulatedCoverHalfWidth
            oEnPoint.Set dInsulatedCoverHalfWidth, parHeightUpperCover, dInsulatedCoverHalfWidth
            
            Dim objUpperCover As Object
            Set objUpperCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperCover
            Set objUpperCover = Nothing
            
            'Create Lower Cover
            dInsulatedCoverHalfWidth = parWidthLowerCover / 2 + parInsulationThickness
                
            oStPoint.Set -dInsulatedCoverHalfWidth, -dHalfMajorAxisEllipticalArc, -dInsulatedCoverHalfWidth
            oEnPoint.Set dInsulatedCoverHalfWidth, -parHeightLowerCover - parInsulationThickness, dInsulatedCoverHalfWidth
            
            Set objInsLowerCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLowerCover
            Set objInsLowerCover = Nothing
            
        Case 330 'Pressure reducing valve, Y-type.
            parHeightCover = arrayOfInputs(3)  'Height
            parWidthCover = arrayOfInputs(6)  'Width
                        
            parHandwheelAngle = PI / 2.4 'Assumption: the Operator is at inclined at 75Deg to x Axis.
            Dim dInclinationAngle As Double
            dInclinationAngle = parHandwheelAngle
            
            'dCylDia is the diameter of the stem that reaches upto the cover and the diameter
            'of the section of horizontal cone that protrudes into the inclined ellipsoid.
            Dim dCylDia As Double
            
            'The Upper Ellipsoid shoud not engulf the upper cover on which operator appears,
            'the diameter of the stem reaching upto the cover should be less than cover width,
            'the base ellipsoid (created below) should not exceed face to face dimension and
            'the section of horizontal cone that protrudes into the inclined ellipsoid is less than Pipe dia.
            dCylDia = ReturnMin(parFacetoFace / 4, pipeDiam, parWidthCover, parHeightCover) _
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
            
            Dim objInsHorizontalSideCylinder As Object
            Set objInsHorizontalSideCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            dInsulationDiameter, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHorizontalSideCylinder
            Set objInsHorizontalSideCylinder = Nothing
                                                                                                                       
            'Create insulated Cover
            'The height of the cover is measured from the center-line to the top-most edge of the inclined cylinder.
            'Effective height denotes the height of the center of the upper cross section of the cover.
            Dim dEffectiveHeightofCover As Double
            dEffectiveHeightofCover = parHeightCover - parWidthCover * 0.5 * Cos(dInclinationAngle)
            
            oStPoint.Set dEffectiveHeightofCover / Tan(dInclinationAngle), _
                    dEffectiveHeightofCover, 0
            oEnPoint.Set oStPoint.x - 0.15 * parHeightCover / Tan(dInclinationAngle), _
                    oStPoint.y - 0.15 * parHeightCover, 0
            
            Dim dInsulationDiaCover As Double
            Dim dInsulationDiaStem As Double
            'Limit the insualtion diameter if it extends to intersect the flange.
            If CmpDblGreaterthan(oEnPoint.y - parWidthCover * 0.5 * Cos(dInclinationAngle) _
                    - parInsulationThickness, ReturnMin(flangeDiam / 2, pipeDiam / 2)) Then
                dInsulationDiaCover = parWidthCover + 2 * parInsulationThickness
                dInsulationDiaStem = dCylDia    'dCylDia is inclusive of insulation Thickness.
            Else
                dInsulationDiaCover = parWidthCover + 0.001
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

        Case 335 'Pressure reducing valve with integral strainer.
            parHeightLowerCover = arrayOfInputs(4)  'Height1 (optional Input)
            parHeightUpperCover = arrayOfInputs(5)  'Height2
            parWidthLowerCover = arrayOfInputs(7)  'Width1 (optional Input)
            parWidthUpperCover = arrayOfInputs(8)  'Width2
            parHeightofStrainer = arrayOfInputs(9)
            parDiaofStrainer = arrayOfInputs(10)    '(optional Input)

            'Assumption: Width of Lower Cover and Width of Upper Cover are optional Inputs.
            ' When not specified Width of Lower Cover = 0.15 x Face to Face,
            '                    Height of Lower Cover = 0.85 x Height of Strainer and
            '                    Diameter of Strainer = 0.2 x Face to Face
            If CmpDblLessThanOrEqualTo(parWidthLowerCover, 0) Then _
                                            parWidthLowerCover = 0.15 * parFacetoFace
         
            If CmpDblLessThanOrEqualTo(parHeightLowerCover, 0) Then _
                                            parHeightLowerCover = 0.85 * parHeightofStrainer
                                            
            If CmpDblLessThanOrEqualTo(parDiaofStrainer, 0) Then _
                                            parDiaofStrainer = 0.2 * parFacetoFace
            
            'Create Insulated Horizontal cylinder
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Set objInsHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            dInsulationDiameter, True)
                        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHorizontalCylinder
            Set objInsHorizontalCylinder = Nothing
            
            'Insert your code for output insulated Upper Cover and stem.
            'Asumption: The Center-axis of Upper Vertical Cylinder is at a distance 15% Face to Face
            '           from the center.
            
            Dim dDistVerticalCoverAxisfromOrigin As Double
            dDistVerticalCoverAxisfromOrigin = parFacetoFace * 0.15
            
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, 0, 0
            oEnPoint.Set oStPoint.x, parHeightUpperCover, 0
            
            dInsulatedVerticalCylinderDia = parWidthUpperCover + 2 * parInsulationThickness
            'Limit the insulation if it extends beyond the valve face.
            If CmpDblGreaterthan(dInsulatedVerticalCylinderDia, _
                        2 * (parFacetoFace / 2 - dDistVerticalCoverAxisfromOrigin)) Then _
                dInsulatedVerticalCylinderDia = 2 * (parFacetoFace / 2 - dDistVerticalCoverAxisfromOrigin)

            Set objInsUpperCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dInsulatedVerticalCylinderDia, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsUpperCover
            Set objInsUpperCover = Nothing
            
            'Insert your code for output insulated Lower Cover and stem
            'Assumption: The Thickness of Lower Cover is equal to 15% of its height.
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, 0, 0
            oEnPoint.Set oStPoint.x, -parHeightLowerCover - parInsulationThickness, 0
            
            dInsulatedVerticalCylinderDia = parWidthLowerCover + 2 * parInsulationThickness
            'Limit the insulation if it extends beyond the valve face.
            If CmpDblGreaterthan(dInsulatedVerticalCylinderDia, _
                         2 * (parFacetoFace / 2 - dDistVerticalCoverAxisfromOrigin)) Then _
                dInsulatedVerticalCylinderDia = 2 * (parFacetoFace / 2 - dDistVerticalCoverAxisfromOrigin)

            Set objInsLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dInsulatedVerticalCylinderDia, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsLowerCover
            Set objInsLowerCover = Nothing
            
            'Insert your code for output insulated Strainer
            ' Assumption: 1. Diameter of Strainer is 20% Face to Face.
            ' 2. The Strainer center-axis is at a distance of 20% of Face to Face from origin.
            oStPoint.Set -parFacetoFace * 0.2, 0, 0
            oEnPoint.Set oStPoint.x, -parHeightofStrainer - parInsulationThickness, 0
            
            dInsulatedVerticalCylinderDia = parDiaofStrainer + 2 * parInsulationThickness
            'Limit the insulation if it extends beyond the valve face.
            If CmpDblGreaterthan(dInsulatedVerticalCylinderDia, 0.6 * parFacetoFace) Then _
                                dInsulatedVerticalCylinderDia = 0.6 * parFacetoFace

            Dim objInsStrainer As Object
            Set objInsStrainer = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dInsulatedVerticalCylinderDia, True)
                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStrainer
            Set objInsStrainer = Nothing
        
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
