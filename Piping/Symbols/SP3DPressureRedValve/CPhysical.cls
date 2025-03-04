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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   Physical.cls
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

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
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
'The center is assumed at Face to Face / 2.

    Dim dOperatorHeight As Double 'used for updating the valve Operator height.
    Dim dOperatorDiameter As Double 'used for updating the valve Operator diameter.
    Dim dNozzleLength As Double 'specifies the length of the nozzle at port 2.
    dNozzleLength = flangeThick 'This is a default value overridden by the cases below if required.
    
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
                                            
            'These values are used in updating the valve Operator.
            dOperatorHeight = parHeightUpperCover
            dOperatorDiameter = parWidthUpperCover
                                            
            'Create left half Horizontal Cylinder.
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set 0, 0, 0
            
            Dim objLeftHorizontalCylinder As Object
            Set objLeftHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftHorizontalCylinder
            Set objLeftHorizontalCylinder = Nothing
            
            'Create right half Horizontal Cylinder.
            oStPoint.Set 0, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Dim objRightHorizontalCylinder As Object
            Set objRightHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightHorizontalCylinder
            Set objRightHorizontalCylinder = Nothing
            
            'Create upper Vertical Cylinder
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, 0, 0
            
            Dim dDiaVerticalCylinder As Double
            'Assumtion : The diameter of the upper vertical cylinder id 70% of width of upper cover.
            dDiaVerticalCylinder = ReturnMin(parWidthUpperCover * 0.7, parWidthLowerCover * 0.7, _
                                          parFacetoFace - 2 * flangeThick)
         
            Dim objUpperVerticalCylinder As Object
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dDiaVerticalCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
            'Create lower Vertical Cylinder
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0, -0.85 * parHeightLowerCover, 0   'The rest 20% height is the cover
            
            Dim objLowerVerticalCylinder As Object
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                     dDiaVerticalCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
            
            'Create Upper Cover
            'Assumption: The Thickness of Upper Cover is equal to 15% of its height.
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, parHeightUpperCover, 0
            
            Dim objUpperCover As Object
            Set objUpperCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidthUpperCover, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperCover
            Set objUpperCover = Nothing
            
            'Create Lower Cover
            'Assumption: The Thickness of Lower Cover is equal to 15% of its height.
            oStPoint.Set 0, -0.85 * parHeightLowerCover, 0
            oEnPoint.Set 0, -parHeightLowerCover, 0
            
            Dim objLowerCover As Object
            Set objLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidthLowerCover, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerCover
            Set objLowerCover = Nothing
            
        Case 320  'Pressure reducing valve, Type 2.
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
            
            'These values are used in updating the valve Operator.
            dOperatorHeight = parHeightUpperCover
            dOperatorDiameter = parWidthUpperCover
            
            'Create left half Horizontal Cylinder.
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set 0, 0, 0
            
            Set objLeftHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftHorizontalCylinder
            Set objLeftHorizontalCylinder = Nothing
            
            'Create right half Horizontal Cylinder.
            oStPoint.Set 0, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            
            Set objRightHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightHorizontalCylinder
            Set objRightHorizontalCylinder = Nothing
            
            'Create Upper Vertical Cylinder
            oStPoint.Set 0, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set 0, 0, 0
                                    
            'Assumtion : The diameter of the upper vertical cylinder id 120% of width of lower cover.
            dDiaVerticalCylinder = ReturnMin(parWidthUpperCover * 0.95, parWidthLowerCover * 1.2, _
                        parFacetoFace - 2 * flangeThick)
            
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                     dDiaVerticalCylinder, True)
                                                                        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
            'Create lower Vertical Cylinder
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0, -0.8 * parHeightLowerCover, 0   'The rest 20% height is the cover
                       
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                     dDiaVerticalCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
            
            'Create Upper Cover
            'Assumption: The Thickness of Upper Cover is equal to 15% of its height.
            oStPoint.Set -parWidthUpperCover / 2, 0.85 * parHeightUpperCover, -parWidthUpperCover / 2
            oEnPoint.Set parWidthUpperCover / 2, parHeightUpperCover, parWidthUpperCover / 2
            
            Set objUpperCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperCover
            Set objUpperCover = Nothing
            
            'Insert your code for output Lower Cover
            'Assumption: The Thickness of Lower Cover is equal to 20% of its height.
            oStPoint.Set 0, -0.8 * parHeightLowerCover, 0
            oEnPoint.Set 0, -parHeightLowerCover, 0
            
            Set objLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidthLowerCover, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerCover
            Set objLowerCover = Nothing
            
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
                                            
            'These values are used in updating the valve Operator.
            dOperatorHeight = parHeightUpperCover
            dOperatorDiameter = parWidthUpperCover
    
            'Create left half Horizontal Cylinder
            'Assumptions: The diameter of horizontal cylinder is equal to Pipe OD.
            
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set 0, oStPoint.y, oStPoint.z
            
            Set objLeftHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftHorizontalCylinder
            Set objLeftHorizontalCylinder = Nothing
            
            'Create right half Horizontal Cylinder.
            oStPoint.Set 0, 0, 0
            oEnPoint.Set parFacetoFace / 2, oStPoint.y, oStPoint.z
            
            Set objRightHorizontalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightHorizontalCylinder
            Set objRightHorizontalCylinder = Nothing
            
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
            dHalfMinorAxisEllipticalArc = parFacetoFace * 0.35 - dDistCenterofEllipArctoOrigin
                        
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
            
            'Create vertical cylinder
            oStPoint.Set 0, dHalfMajorAxisEllipticalArc, 0
            oEnPoint.Set 0, -dHalfMajorAxisEllipticalArc, oStPoint.z
            
            Dim objVerticalCylinder As Object
            Set objVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                dDistCenterofEllipArctoOrigin, True)
                                                
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVerticalCylinder
            Set objVerticalCylinder = Nothing
            
            'Create Upper Cover
            oStPoint.Set -parWidthUpperCover / 2, dHalfMajorAxisEllipticalArc, -parWidthUpperCover / 2
            oEnPoint.Set parWidthUpperCover / 2, parHeightUpperCover, parWidthUpperCover / 2
            
            Set objUpperCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperCover
            Set objUpperCover = Nothing
            
            'Create Lower Cover
            oStPoint.Set -parWidthLowerCover / 2, -dHalfMajorAxisEllipticalArc, -parWidthLowerCover / 2
            oEnPoint.Set parWidthLowerCover / 2, -parHeightLowerCover, parWidthLowerCover / 2
            
            Set objLowerCover = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerCover
            Set objLowerCover = Nothing
            
        Case 330 'Pressure reducing valve, Y-type.
            parHeightCover = arrayOfInputs(3)  'Height
            parWidthCover = arrayOfInputs(6)  'Width
                        
            'These values are used in updating the valve Operator.
            parHandwheelAngle = PI / 2.4 'Assumption: the Operator is at inclined at 75Deg to x Axis.
            Dim dInclinationAngle As Double
            dInclinationAngle = parHandwheelAngle
            
            'The height of the cover is measured from the center-line to the top-most edge of the inclined cylinder.
            'Effective height denotes the height of the center of the upper cross section of the cover.
            Dim dEffectiveHeightofCover As Double
            dEffectiveHeightofCover = parHeightCover - parWidthCover * 0.5 * Cos(dInclinationAngle)
            
            dOperatorHeight = dEffectiveHeightofCover * (2 - Sin(parHandwheelAngle))  'The operator places above the cover.
            dOperatorDiameter = parWidthCover
                        
            'dCylDia is the diameter of the stem that reaches upto the cover and the diameter
            'of the section of horizontal cone that protrudes into the inclined ellipsoid.
            Dim dCylDia As Double
            
            'The Ellipsoid shoud not engulf the upper cover on which operator appears,
            'the diameter of the stem reaching upto the cover should be less than cover width,
            'the base ellipsoid (created below) should not exceed face to face dimension and
            'the section of horizontal cone that protrudes into the inclined ellipsoid is less than Pipe dia.
            dCylDia = ReturnMin(parFacetoFace / 4, pipeDiam, parWidthCover, parHeightCover)
            
            'Create horizontal Cone adjacent to horizontal cylinder at port1 and protruding into inclined ellipsoid.
            oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oEnPoint.Set 0, 0, 0
                                        
            Dim objHorizontalLeftCone As Object
            Set objHorizontalLeftCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                                                pipeDiam / 2, dCylDia / 2)
            
                     
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorizontalLeftCone
            Set objEllipsoid = Nothing
            Set objHorizontalLeftCone = Nothing
            
            'Create the inclined Ellipsoid
            Dim dHalfMajorAxisInclinedEllipArc As Double
            dHalfMajorAxisInclinedEllipArc = dCylDia * 1.1
                                           
            Dim dHalfMinorAxisInclinedEllipArc As Double
            dHalfMinorAxisInclinedEllipArc = dCylDia * 0.85
                
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
            
            'Create horizontal Cylinder at port 2
            oStPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oEnPoint.Set (parFacetoFace / 2) * 0.8 - flangeThick, 0, 0
            
            Dim objHorizontalSideCylinder As Object
            Set objHorizontalSideCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorizontalSideCylinder
            Set objHorizontalSideCylinder = Nothing
                                                                                                           
            'Create horizontal Cone adjacent to horizontal cylinder at port2 and protruding into inclined ellipsoid.
            oStPoint.Set (parFacetoFace / 2) * 0.8 - flangeThick, 0, 0
            oEnPoint.Set 0, 0, 0
                                        
            Dim objHorizontalCone As Object
            Set objHorizontalCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                                                pipeDiam / 2, dCylDia / 2)
                                                                                               
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorizontalCone
            Set objHorizontalCone = Nothing
                                                                           
            'Create Cover
            oStPoint.Set dEffectiveHeightofCover / Tan(dInclinationAngle), _
                    dEffectiveHeightofCover, 0
            oEnPoint.Set oStPoint.x - 0.15 * parHeightCover / Tan(dInclinationAngle), _
                    oStPoint.y - 0.15 * parHeightCover, 0
            
            Dim objCover As Object
            Set objCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            parWidthCover, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCover
            Set objCover = Nothing
            
            'Create stem reaching upto the bottom of the cover
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z     'End point of the previous cylinder
            
            Dim objStem As Object
            Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    dCylDia, True)
            
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
                                                                                                       
            'These values are used in updating the valve Operator.
            dOperatorHeight = parHeightUpperCover
            dOperatorDiameter = parWidthUpperCover
            
            dNozzleLength = parFacetoFace
                                                                                                       
            'Insert your code for output Upper Vertical Cylinder
            ' Assumption: 1.Diameter of Upper Vertical Cylinder is 70% Width of upper cover.
            ' 2. The Center-axis of Upper Vertical Cylinder is at a distance 15% Face to Face
            '    from the center.
            
            Dim dDistVerticalCoverAxisfromOrigin As Double
            dDistVerticalCoverAxisfromOrigin = parFacetoFace * 0.15
            
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, 0, 0
            oEnPoint.Set oStPoint.x, 0.85 * parHeightUpperCover, 0
            
            Dim dDiameterofCylinder As Double
            dDiameterofCylinder = ReturnMin(0.7 * parWidthUpperCover, parFacetoFace * 0.4)
            
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dDiameterofCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
            'Insert your code for output Upper Cover
            'Assumption: The Thickness of Upper Cover is equal to 15% of its height.
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, 0.85 * parHeightUpperCover, 0
            oEnPoint.Set oStPoint.x, parHeightUpperCover, 0
            
            Set objUpperCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidthUpperCover, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperCover
            Set objUpperCover = Nothing
            
            'Insert your code for output Lower Vertical Cylinder
            ' Assumption: Diameter of Lower Vertical Cylinder is 70% width of lower cover.
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, 0, 0
            oEnPoint.Set oStPoint.x, -0.85 * parHeightLowerCover, 0
            
            dDiameterofCylinder = ReturnMin(0.7 * parWidthLowerCover, parFacetoFace * 0.4)
                        
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dDiameterofCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
            
            'Insert your code for output Lower Cover
            'Assumption: The Thickness of Lower Cover is equal to 15% of its height.
            oStPoint.Set dDistVerticalCoverAxisfromOrigin, -0.85 * parHeightLowerCover, 0
            oEnPoint.Set oStPoint.x, -parHeightLowerCover, 0
            
            Set objLowerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parWidthLowerCover, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerCover
            Set objLowerCover = Nothing
            
            'Insert your code for output Strainer
            ' Assumption: 1. Diameter of Strainer is 80% Diamter of Strainer cover.
            ' 2. The Strainer center-axis is at a distance of 20% of Face to Face from origin.
            oStPoint.Set -parFacetoFace * 0.2, 0, 0
            oEnPoint.Set oStPoint.x, -0.85 * parHeightofStrainer, 0 'The rest 15% height is the cover
            
            dDiameterofCylinder = ReturnMin(0.8 * parDiaofStrainer, parFacetoFace * 0.4, pipeDiam)
                        
            Dim objStrainer As Object
            Set objStrainer = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    dDiameterofCylinder, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrainer
            Set objStrainer = Nothing
            
            'Insert your code for output Strainer Cover at its bottom
            ' Assumption: The Thickness of Strainer Cover is equal to 15% of its height.
            oStPoint.Set -parFacetoFace * 0.2, -0.85 * parHeightofStrainer, 0
            oEnPoint.Set oStPoint.x, -parHeightofStrainer, 0
            
            Dim objStrainerCover As Object
            Set objStrainerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    parDiaofStrainer, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStrainerCover
            Set objStrainerCover = Nothing
        
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

' Insert your code for Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0#, 0#
    oDir.Set -1, 0#, 0#
    Dim objNozzle1 As GSCADNozzleEntities.IJDNozzle
    Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    
 ' Insert your code for Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0#, 0#
    oDir.Set 1, 0#, 0#
    
    Dim objNozzle2 As GSCADNozzleEntities.IJDNozzle
    Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

'   Insert your code for Valve Actuator
'-----------------------------------------------------------------------------------------------------
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    If lPartDataBasis = 330 Then    '' Pressure reducing valve, Y-type
        oDirX.Set Sin(parHandwheelAngle), -Cos(parHandwheelAngle), 0
        oDirY.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
        oDirZ.Set 0, 0, 1
    Else
        oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
        oDirY.Set 0, 1, 0
        oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    End If
    
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel

        Set oOperatorPart = oPipeComponent.GetValveOperatorPart

        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            
            'For the follwoing cases the operator is at an offset from the origin of the valve.
            If lPartDataBasis = 330 Then    'Pressure Reducing Valve, Y-Type.
                OpOrigin.Set (dEffectiveHeightofCover / Tan(parHandwheelAngle)) - dEffectiveHeightofCover * Cos(parHandwheelAngle), 0, 0
            ElseIf lPartDataBasis = 335 Then    'Pressure Reducing Valve having integral strainer
                OpOrigin.Set dDistVerticalCoverAxisfromOrigin, 0, 0
            Else
                OpOrigin.Set 0, 0, 0
            End If
            
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        
            '   Update Occurrence Attributes of Operator.
            Dim oIJDInputsArg As IJDInputsArg
            Dim oIJDEditJDArgument As IJDEditJDArgument
        
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
        
            UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorDiameter", dOperatorDiameter, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight1", dOperatorHeight, oIJDEditJDArgument
        
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
        End If
    End If
    
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
   
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                            vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)
    
    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel
    
    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
    Dim lInputIndex As Long
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long
    
    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1
    
    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub
        
    ' Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
    oIJDAttr.Value = vInputParameterValue
    
    ' Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
   
End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                    oIJDEditJDArgument As IJDEditJDArgument)
                    
    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel
    
    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument
                
    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC
    
    oIJDEditJDArgument.SetArg oArgument
    
    Set oArgument = Nothing
    Set oPC = Nothing
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
        
End Sub
Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg
    
    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument
    
    Exit Function
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Function

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
