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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          dkl
'   Creation Date:  Thursday, Jan 31 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'   This symbol is made to function as a Vertical adjustable splice plate for Part data basis it supports.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   31.Jan.2008     dkl    CR-130446 Created the symbol.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const PLATE_THICKNESS = 0.01
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_outputcoll As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim iOutput As Integer
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parArmLength As Double
    Dim parPorttoHinge As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parAngle As Double
    Dim dFlipUpwardorDownward As Integer

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parArmLength = arrayOfInputs(2)
    parPorttoHinge = arrayOfInputs(3)
'    parWidth = arrayOfInputs(4)
'    parDepth = arrayOfInputs(5)
'    parAngle = arrayOfInputs(6)
'    dFlipUpwardorDownward = arrayOfInputs(7)

    iOutput = 0
    
    Dim dHalfWidth As Double
    Dim dHalfDepth As Double
    Dim dBoltingPlatePart As Double  'The part of plate that is aligned with the tray for bolting.
    Dim dPorttoCenter As Double 'determines the location of ports with respect to the center of symbol.
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    Dim IPartDataBasis As Integer
    IPartDataBasis = oTrayPart.PartDataBasis
    
    'This symbol implements the part data bases honouring the depth to which the splice plate goes into the tray.
    'However in future we may have part data basis implementing insertion depth parameter.
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLoc As AutoMath.DPosition
    Set oPortLoc = New AutoMath.DPosition

'    Determine angle on the part
    Select Case oTrayPart.ComponentType
        Case 555, 557   '15 degrees
            parAngle = PI / 12
        Case 60, 65 '30 degrees
            parAngle = PI / 6
        Case 45, 50 '45 degrees
            parAngle = PI / 4
        Case 30, 35 '60 degrees
            parAngle = PI / 3
        Case 15, 20 '90 degrees
            parAngle = PI / 2
        Case Else   'variable angle
            parAngle = arrayOfInputs(6)
        End Select
    
    'The orientation of port (Inside/Outside) 2 is determined from value of Cable Tray Component Type.
    Select Case oTrayPart.ComponentType
        Case 15, 30, 45, 60, 555, 79
    '        90 Degree vertical inside bend (15)
    '        60 Degree vertical inside bend (30)
    '        45 Degree vertical inside bend (45)
    '        30 Degree vertical inside bend (60)
    '        15 Degree vertical inside bend (555)
    '        Vertical adjustable bend (0-90 Deg), inside (79)
            dFlipUpwardorDownward = 1
        Case 20, 35, 50, 65, 557, 80
    '        90 Degree vertical outside bend (20)
    '        60 Degree vertical outside bend (35)
    '        45 Degree vertical outside bend (50)
    '        30 Degree vertical outside bend (65)
    '        15 Degree vertical outside bend (557)
    '        Vertical adjustable bend (0-90 Deg), outside (80)
            dFlipUpwardorDownward = -1
        Case Else
            GoTo ErrorLabel
    End Select
       
    'Port 2 is flipped for Inside or Outside routing based on value of dFlipUpwardorDownward.
    parAngle = dFlipUpwardorDownward * parAngle
    
    ' Retrieve Port 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    'Port to Hinge is not optional parameter and has to have non-zero value.
    If CmpDblLessThanOrEqualTo(parPorttoHinge, 0) Then GoTo ErrorLabel

'This symbol supports the following Part data basis that govern its geometry,
'Splice plate, Vertical adjustable, attached at sides of tray, specified by Arm length, Hinge to port and Depth (134).
'Splice plate, Vertical adjustable, attached at bottom of tray, specified by Length, Hinge to port and Width (135).
'Splice plate, Vertical adjustable, attached at center of tray, specified by Arm length, Hinge to port, Width and Depth (136).
    
    Select Case IPartDataBasis
        Case Is <= 1, 134   'Splice plate, Vertical adjustable, attached at sides of tray,
                            'specified by Arm length, Hinge to port, Width and Depth
            dPorttoCenter = parPorttoHinge
            parDepth = arrayOfInputs(5)
            dHalfDepth = parDepth / 2
            dHalfWidth = parActualWidth / 2 + NEGLIGIBLE_THICKNESS
            If CmpDblLessThanOrEqualTo(dHalfDepth, 0) Then dHalfDepth = parActualDepth / 2.5
            'If Arm length is not specified, it is assumed 1.6 times port to hinge.
            If CmpDblLessThanOrEqualTo(parArmLength, 0) Then parArmLength = parPorttoHinge * 1.6
            dBoltingPlatePart = 0.8 * (parArmLength - parPorttoHinge)
            
            'construct horizontal arm 1 on right side.
            Dim oLowerHorLine As IngrGeom3D.Line3d
            Set oLowerHorLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        0, -dHalfWidth, -dHalfDepth, _
                                        -1, 0, 0, parPorttoHinge + dBoltingPlatePart)
            
            Dim oCurvesColl As Collection
            Set oCurvesColl = New Collection
            oCurvesColl.Add oLowerHorLine
            Set oLowerHorLine = Nothing
            
            Dim oVerticalLine As IngrGeom3D.Line3d
            Set oVerticalLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        -(parPorttoHinge + dBoltingPlatePart), -dHalfWidth, -dHalfDepth, _
                                        0, 0, 1, dHalfDepth * 2)
            
            oCurvesColl.Add oVerticalLine
            Set oVerticalLine = Nothing
            
            Dim oUpperHorLine As IngrGeom3D.Line3d
            Set oUpperHorLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                        -(parPorttoHinge + dBoltingPlatePart), -dHalfWidth, dHalfDepth, _
                                        1, 0, 0, parPorttoHinge + dBoltingPlatePart)
            
            oCurvesColl.Add oUpperHorLine
            Set oUpperHorLine = Nothing
            
            Dim oArcatHinge As IngrGeom3D.EllipticalArc3d
            Set oArcatHinge = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, -dHalfWidth, 0, _
                    0, 1, 0, _
                    0, 0, dHalfDepth, _
                    (parArmLength - parPorttoHinge - dBoltingPlatePart) / dHalfDepth, 0, PI)
            oCurvesColl.Add oArcatHinge
            Set oArcatHinge = Nothing
            
            Dim oArm1Boundary As IngrGeom3D.ComplexString3d
            Dim oStPoint As AutoMath.DPosition
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set 0, -dHalfWidth, -dHalfDepth
            Set oArm1Boundary = PlaceTrCString(oStPoint, oCurvesColl)
            Set oCurvesColl = Nothing
            Set oStPoint = Nothing

            Dim oArm1onRightSide As IngrGeom3D.Plane3d
            Set oArm1onRightSide = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputcoll.ResourceManager, oArm1Boundary)

            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), oArm1onRightSide
            Set oArm1onRightSide = Nothing

            'constructing Arm 2 on right side.
            Dim oArm2onRightSide As IngrGeom3D.Plane3d
            Set oArm2onRightSide = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputcoll.ResourceManager, oArm1Boundary)
            Dim oTransformationMatrix As IJDT4x4
            Set oTransformationMatrix = New DT4x4
            oTransformationMatrix.LoadIdentity
            Dim oVector As AutoMath.DVector
            Set oVector = New AutoMath.DVector
            oVector.Set 0, 1, 0
            oTransformationMatrix.Rotate (PI - parAngle), oVector
            oArm2onRightSide.Transform oTransformationMatrix

            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), oArm2onRightSide
            Set oArm2onRightSide = Nothing

'           constructing Arm 1 on left side.
            Dim oArm1onLeftSide As IngrGeom3D.Plane3d
            Set oArm1onLeftSide = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputcoll.ResourceManager, oArm1Boundary)
            oTransformationMatrix.LoadIdentity
            oVector.Set 0, 2 * dHalfWidth, 0
            oTransformationMatrix.Translate oVector
            oArm1onLeftSide.Transform oTransformationMatrix

            m_outputcoll.AddOutput "Arm1onLeftSide", oArm1onLeftSide
            Set oArm1onLeftSide = Nothing

            'constructing Arm 2 on left side.
            Dim oArm2onLeftSide As IngrGeom3D.Plane3d
            Set oArm2onLeftSide = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputcoll.ResourceManager, oArm1Boundary)
            oVector.Set 0, 1, 0
            oTransformationMatrix.Rotate (PI - parAngle), oVector
            oArm2onLeftSide.Transform oTransformationMatrix

            m_outputcoll.AddOutput "Arm2onLeftSide", oArm2onLeftSide
            Set oArm2onLeftSide = Nothing
            Set oVector = Nothing
            Set oTransformationMatrix = Nothing
            Set oArm1Boundary = Nothing
            
        Case 135 'Splice plate, Vertical adjustable, attached at bottom of tray,
                 'specified by Length, Hinge to port and Width
            parWidth = arrayOfInputs(4)
            dHalfWidth = parWidth / 2
            dHalfDepth = parActualDepth / 2
            If CmpDblLessThanOrEqualTo(dHalfWidth, 0) Then dHalfWidth = parActualWidth / 2
            Dim parPlateOverallLength As Double
            parPlateOverallLength = oTrayPart.Length
            'If Plate Overall Length is not specified, it is computed as 3.2 times Port to Hinge,
            If CmpDblLessThanOrEqualTo(parPlateOverallLength, 0) Then parPlateOverallLength = parPorttoHinge * 3.2
            parArmLength = parPlateOverallLength / 2
            Dim dHingetoCenterAlongX As Double
            dHingetoCenterAlongX = dHalfDepth * Tan(parAngle / 2)
            dPorttoCenter = parPorttoHinge - dHingetoCenterAlongX
            Dim dArmLength As Double    'dyanmic arm length computed based on center of symbol.
            dArmLength = parArmLength - dHingetoCenterAlongX
            'Construct plate at the bottom of tray 1.
            Dim oUpperLineofPlate As IngrGeom3D.Line3d
            Set oUpperLineofPlate = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -dArmLength, -dHalfWidth, -dHalfDepth, _
                dHingetoCenterAlongX, -dHalfWidth, -dHalfDepth)
            Set oCurvesColl = New Collection
            oCurvesColl.Add oUpperLineofPlate
            Set oUpperLineofPlate = Nothing

            Dim oVerArcatHinge As IngrGeom3D.EllipticalArc3d
            Set oVerArcatHinge = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                dHingetoCenterAlongX, -dHalfWidth, -(dHalfDepth + PLATE_THICKNESS / 2), _
                0, 1, 0, _
                0, 0, PLATE_THICKNESS / 2, _
                1, 0, PI)
            oCurvesColl.Add oVerArcatHinge
            Set oVerArcatHinge = Nothing
                
            Dim oBottomLineofPlate As IngrGeom3D.Line3d
            Set oBottomLineofPlate = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dHingetoCenterAlongX, -dHalfWidth, -(dHalfDepth + PLATE_THICKNESS), _
                -dArmLength, -dHalfWidth, -(dHalfDepth + PLATE_THICKNESS))
            oCurvesColl.Add oBottomLineofPlate
            Set oBottomLineofPlate = Nothing
            
            Dim oVerLineofPlate As IngrGeom3D.Line3d
            Set oVerLineofPlate = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -dArmLength, -dHalfWidth, -(dHalfDepth + PLATE_THICKNESS), _
                -dArmLength, -dHalfWidth, -dHalfDepth)
            oCurvesColl.Add oVerLineofPlate
            Set oVerLineofPlate = Nothing
            
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set -dArmLength, -dHalfWidth, -dHalfDepth
            Dim oBoundary As IngrGeom3D.ComplexString3d
            Set oBoundary = PlaceTrCString(oStPoint, oCurvesColl)
            Set oCurvesColl = Nothing
                
            Dim oBasePlate1 As Object
            Set oBasePlate1 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, oBoundary, _
                                0, 1, 0, dHalfWidth * 2, True)
                                                    
            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), oBasePlate1
            Set oBasePlate1 = Nothing
            
            'construct plate at the bottom of tray 2.
            Dim oBasePlate2 As Object
            Set oBasePlate2 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, oBoundary, _
                                0, 1, 0, dHalfWidth * 2, True)

            Set oTransformationMatrix = New DT4x4
            oTransformationMatrix.LoadIdentity
            Set oVector = New AutoMath.DVector
            oVector.Set 0, 1, 0
            oTransformationMatrix.Rotate (PI - parAngle), oVector
            oVector.Set 0, 0, (dHalfDepth * 2 + PLATE_THICKNESS)
            oTransformationMatrix.Translate oVector
            oBasePlate2.Transform oTransformationMatrix

            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), oBasePlate2
            Set oBoundary = Nothing
            Set oBasePlate2 = Nothing
            
        Case 136 'Splice plate, Vertical adjustable, attached at center of tray,
                 'specified by Arm length, Hinge to port, Width and Depth.
            parWidth = arrayOfInputs(4)
            parDepth = arrayOfInputs(5)
            dHalfWidth = parWidth / 2
            dHalfDepth = parDepth / 2
            If CmpDblLessThanOrEqualTo(dHalfWidth, 0) Then dHalfWidth = parActualWidth / 12
            If CmpDblLessThanOrEqualTo(dHalfDepth, 0) Then dHalfDepth = parActualDepth / 2.5
            dHingetoCenterAlongX = (parActualDepth / 2 - dHalfDepth) * Tan(parAngle / 2)
            dPorttoCenter = parPorttoHinge - dHingetoCenterAlongX
            'If Arm length is not specified, it is assumed 1.6 times port to hinge.
            If CmpDblLessThanOrEqualTo(parArmLength, 0) Then parArmLength = parPorttoHinge * 1.6
            dBoltingPlatePart = 0.1 * (parArmLength - parPorttoHinge)
            Dim dHalfArmwidth As Double 'Arm width is assumed to be 50% of width specified.
            dHalfArmwidth = dHalfWidth / 2
            
            'construct box on arm 1.
            Dim oLineVerticalNegY As IngrGeom3D.Line3d
            Set oLineVerticalNegY = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                            -(parPorttoHinge - dHingetoCenterAlongX), -dHalfWidth, -parActualDepth / 2, _
                            0, 0, 1, 2 * dHalfDepth)
            Set oCurvesColl = New Collection
            oCurvesColl.Add oLineVerticalNegY
            Set oLineVerticalNegY = Nothing

            Dim oLineHorTop As IngrGeom3D.Line3d
            Set oLineHorTop = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                            -(parPorttoHinge - dHingetoCenterAlongX), -dHalfWidth, 2 * dHalfDepth - parActualDepth / 2, _
                            0, 1, 0, 2 * dHalfWidth)
            oCurvesColl.Add oLineHorTop
            Set oLineHorTop = Nothing
            
            Dim oLineVerticalPosY As IngrGeom3D.Line3d
            Set oLineVerticalPosY = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                            -(parPorttoHinge - dHingetoCenterAlongX), dHalfWidth, 2 * dHalfDepth - parActualDepth / 2, _
                            0, 0, -1, 2 * dHalfDepth)
            oCurvesColl.Add oLineVerticalPosY
            Set oLineVerticalPosY = Nothing
            
            Dim oLineHorBottom As IngrGeom3D.Line3d
            Set oLineHorBottom = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                            -(parPorttoHinge - dHingetoCenterAlongX), dHalfWidth, -parActualDepth / 2, _
                            0, -1, 0, 2 * dHalfWidth)
            oCurvesColl.Add oLineHorBottom
            Set oLineHorBottom = Nothing

            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set -(parPorttoHinge - dHingetoCenterAlongX), -dHalfWidth, -parActualDepth / 2
            Dim oBoxBoundary As IngrGeom3D.ComplexString3d
            Set oBoxBoundary = PlaceTrCString(oStPoint, oCurvesColl)
            Set oCurvesColl = Nothing
            
            Dim oBoxatArm1 As IngrGeom3D.Projection3d
            Set oBoxatArm1 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                            oBoxBoundary, -1, 0, 0, dBoltingPlatePart, True)

            iOutput = iOutput + 1
            m_outputcoll.AddOutput arrayOfOutputs(iOutput), oBoxatArm1
            Set oBoxatArm1 = Nothing

            'construct arm 1.
            Set oLowerHorLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                                dHingetoCenterAlongX, -dHalfArmwidth, -parActualDepth / 2, _
                                -1, 0, 0, parPorttoHinge)
            
            Set oCurvesColl = New Collection
            oCurvesColl.Add oLowerHorLine
            Set oLowerHorLine = Nothing
            
            Set oVerticalLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                    -(parPorttoHinge - dHingetoCenterAlongX), -dHalfArmwidth, -parActualDepth / 2, _
                    0, 0, 1, 2 * dHalfDepth)
            oCurvesColl.Add oVerticalLine
            Set oVerticalLine = Nothing
            
            Set oUpperHorLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, _
                -(parPorttoHinge - dHingetoCenterAlongX), -dHalfArmwidth, 2 * dHalfDepth - parActualDepth / 2, _
                1, 0, 0, parPorttoHinge)
            oCurvesColl.Add oUpperHorLine
            Set oUpperHorLine = Nothing
            
            Set oArcatHinge = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    dHingetoCenterAlongX, -dHalfArmwidth, -(parActualDepth / 2 - dHalfDepth), _
                    0, 1, 0, _
                    0, 0, dHalfDepth, _
                    (parArmLength - parPorttoHinge - dBoltingPlatePart) / dHalfDepth, 0, PI)
            oCurvesColl.Add oArcatHinge
            Set oArcatHinge = Nothing
            
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set dHingetoCenterAlongX, -dHalfArmwidth, -parActualDepth / 2
            Set oArm1Boundary = PlaceTrCString(oStPoint, oCurvesColl)
            Set oCurvesColl = Nothing
            Set oStPoint = Nothing
            
            Dim oArm1 As IngrGeom3D.Projection3d
            Set oArm1 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, oArm1Boundary, _
                                                0, 1, 0, _
                                                2 * dHalfArmwidth, True)

            m_outputcoll.AddOutput "Arm1", oArm1
            Set oArm1 = Nothing
            
            'Construct arm 2.
            Dim oArm2 As IngrGeom3D.Projection3d
            Set oArm2 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, oArm1Boundary, _
                                                0, 1, 0, _
                                                2 * dHalfArmwidth, True)

            Set oTransformationMatrix = New DT4x4
            oTransformationMatrix.LoadIdentity
            Set oVector = New AutoMath.DVector
            oVector.Set 0, 1, 0
            oTransformationMatrix.Rotate (PI - parAngle), oVector
            oVector.Set 0, 0, 2 * (parActualDepth / 2 - dHalfDepth)
            oTransformationMatrix.Translate oVector
            oArm2.Transform oTransformationMatrix

            m_outputcoll.AddOutput "Arm2", oArm2
            Set oArm2 = Nothing
            
            'construct box on arm 2.
            Dim oBoxatArm2 As IngrGeom3D.Projection3d
            Set oBoxatArm2 = oGeomFactory.Projections3d.CreateByCurve(m_outputcoll.ResourceManager, _
                            oBoxBoundary, -1, 0, 0, dBoltingPlatePart, True)
            oBoxatArm2.Transform oTransformationMatrix

            m_outputcoll.AddOutput "BoxatArm2", oBoxatArm2
            Set oBoxBoundary = Nothing
            Set oBoxatArm2 = Nothing
            Set oVector = Nothing
            Set oTransformationMatrix = Nothing
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    Set oTrayPart = Nothing
' Place Port 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim oPortLocation As AutoMath.DPosition
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set -dPorttoCenter, 0, 0

    oPortLoc.Set oPortLocation.x - dInsertionDepth * oDir.x, oPortLocation.y - dInsertionDepth * oDir.y, oPortLocation.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLoc, oDir, oRadialOrient, m_outputcoll)

' Set the output
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    Set oPortLoc = Nothing
    
' Place Port 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    Set oPortLoc = New AutoMath.DPosition
    oDir.Set Cos(parAngle), 0, Sin(parAngle)
    oRadialOrient.Set -Sin(parAngle), 0, Cos(parAngle)
    oPortLocation.Set dPorttoCenter * Cos(parAngle), 0, dPorttoCenter * Sin(parAngle)
    
    oPortLoc.Set oPortLocation.x - dInsertionDepth * oDir.x, oPortLocation.y - dInsertionDepth * oDir.y, oPortLocation.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLoc, oDir, oRadialOrient, m_outputcoll)

' Set the output
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPortLocation = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLoc = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
