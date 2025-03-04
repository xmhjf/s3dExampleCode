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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Thursday, 14th Oct 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   18.FEB.2005     MS     TR 71938 The present intent is to only support
'                                           tangential non-radials which means that the offset will always be
'                                           1/2 the OD of the header minus 1/2 the OD of the branch pipe. Hence removed
'                                           ComponentOffset as input to symbol.
'   23.FEB.2005     MS     The branch-port is now located at the center of the header with the required offset to be non-radial
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private RAD          As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjPad As Object
    Dim ObjInsulatedPad As Object
    
    Dim parRPadWidth As Double
    Dim parRPadThickness As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parRPadWidth = arrayOfInputs(2)
    parRPadThickness = arrayOfInputs(3)
    parAngle = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 ' Insert your code for output 1(Reinforcing Pad)
    Dim a, b, c, d As Double
    Dim p(84) As Double
    Dim pr, sr, w, t As Double
    Dim ang, angle, xoff, prt, srp, prtl, prl, csd As Double
    Dim sina As Double
    Dim cosa As Double
    
'( value=pr, prompt="keyin major radius")
'( value=sr, prompt="keyin minor radius")
'( value=w, prompt="keyin width")
'( value=t, prompt="keyin thickness")
'( value=angle, prompt="keyin angle of branch 90 is straight out")

    angle = Abs(parAngle)
    ang = angle
    sina = Sin(ang)
    cosa = Cos(ang)
    w = parRPadWidth
    t = parRPadThickness
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    pr = pipeDiam / 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    sr = pipeDiam / 2
    'The offset will always be 1/2 the OD of the header minus 1/2 the OD of the branch pipe
    Dim dComponentOffset As Double
    dComponentOffset = pr - sr
        
    prt = pr + t
    xoff = prt / Tan(ang)
    srp = sr / Sin(ang)
''angle of wrapped width
    a = (360 * w) / (2 * PI * (pr + t) * RAD)
''angle of the width of branch at main radius plus thickness
    Dim parBB As Double
    parBB = sr / prt
    b = Atn(parBB / Sqr(-parBB * parBB + 1))
    c = a + b
''this is the angle to calculate the pole length from center of main pipe
    d = c / 3
    csd = Cos(d)
''this is the length of the exterior poles from the center
    prtl = prt / csd
''this is the length of the interior poles from the center
    prl = pr / csd
    
    ''/* curve 1 */
    
    p(0) = xoff
    p(1) = Cos(c) * pr
    p(2) = Sin(c) * pr
    
    p(3) = xoff + (0.25 * (w + srp))
    p(4) = p(1)
    p(5) = p(2)
    
    p(6) = xoff + (0.75 * (w + srp))
    p(7) = Cos(c * 0.6666) * prl
    p(8) = Sin(c * 0.6666) * prl
    
    p(9) = xoff + (1.077 * (w + srp))
    p(10) = prl
    p(11) = 0
    
    p(12) = p(6)
    p(13) = p(7)
    p(14) = -p(8)
    
    p(15) = p(3)
    p(16) = p(4)
    p(17) = -p(5)
    
    p(18) = p(0)
    p(19) = p(1)
    p(20) = -p(2)
    
    ''/* curve 2 */
    
    p(21) = xoff
    p(22) = Cos(c) * (pr + t)
    p(23) = Sin(c) * (pr + t)
    
    p(24) = xoff + (0.25 * (w + srp))
    p(25) = p(22)
    p(26) = p(23)
    
    p(27) = xoff + (0.75 * (w + srp))
    p(28) = Cos(c * 0.6666) * prtl
    p(29) = Sin(c * 0.6666) * prtl
    
    p(30) = xoff + (1.077 * (w + srp))
    p(31) = prtl
    p(32) = 0
    
    p(33) = p(27)
    p(34) = p(28)
    p(35) = -p(29)
    
    p(36) = p(24)
    p(37) = p(25)
    p(38) = -p(26)
    
    p(39) = p(21)
    p(40) = p(22)
    p(41) = -p(23)
    
    ''/* curve 3 */
    
    p(42) = p(21)
    p(43) = p(22)
    p(44) = p(23)
    
    p(45) = xoff - (0.25 * (w + srp))
    p(46) = p(25)
    p(47) = p(26)
    
    p(48) = xoff - (0.75 * (w + srp))
    p(49) = p(28)
    p(50) = p(29)
    
    p(51) = xoff - (1.077 * (w + srp))
    p(52) = p(31)
    p(53) = 0
    
    p(54) = p(48)
    p(55) = p(49)
    p(56) = -p(50)
    
    p(57) = p(45)
    p(58) = p(46)
    p(59) = -p(47)
    
    p(60) = p(42)
    p(61) = p(43)
    p(62) = -p(44)
    
    ''/* curve 4 */
    
    p(63) = xoff
    p(64) = p(1)
    p(65) = p(2)
    
    p(66) = xoff - (0.25 * (w + srp))
    p(67) = p(4)
    p(68) = p(5)
    
    p(69) = xoff - (0.75 * (w + srp))
    p(70) = p(7)
    p(71) = p(8)
    
    p(72) = xoff - (1.077 * (w + srp))
    p(73) = p(10)
    p(74) = p(11)
    
    p(75) = p(69)
    p(76) = p(70)
    p(77) = -p(71)
    
    p(78) = p(66)
    p(79) = p(67)
    p(80) = -p(68)
    
    p(81) = p(63)
    p(82) = p(64)
    p(83) = -p(65)
    Dim parNumPolesV As Long
    Dim parNumPolesU As Long
    Dim parUorder As Long
    Dim parVorder As Long
    Dim fWeights(0) As Double, uKnots(0) As Double, vKnots(0) As Double
    parNumPolesV = 4
    parNumPolesU = 7
    parUorder = 3
    parVorder = 2
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set ObjPad = geomFactory.BSplineSurfaces3d.CreateByParameters(m_OutputColl.ResourceManager, _
        parNumPolesU, parNumPolesV, _
        p, fWeights, _
        parUorder, parVorder, _
        uKnots, vKnots, False, False, False)

'Transform the output  (Rotation about header axis)
    Dim dProjLength As Double
    Dim dOffsetangle As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    dProjLength = Sqr((pipeDiam / 2) ^ 2 - (dComponentOffset ^ 2))
    dOffsetangle = Atn(dComponentOffset / dProjLength)
    
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As New AutoMath.DVector
    oAboutVec.Set 1, 0, 0
    
    oTransMat.Rotate dOffsetangle, oAboutVec
    ObjPad.Transform oTransMat
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPad
    Set ObjPad = Nothing
    
    
' Place Nozzle 1
    
'    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0

    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
' Case Equal diameters branching
    Dim MiniDistance As Double
    If pr = sr Then
        MiniDistance = 0
    Else
        If parAngle = PI / 2 Then
            MiniDistance = Sqr(pr * pr - sr * sr)
        Else
' Define closest distance between plane normal to branch and intersion between branch and Header curve
    
    Dim objPlane As IngrGeom3D.Plane3d
        
'   plane normal to branch
    Set objPlane = geomFactory.Planes3d.CreateByPointNormal(Nothing, 0, 0, 0, cosa, sina, 0)

' Alternate curve = Intersection between Header and Branch cylinders
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim objCylinderHeader  As Object
    Dim objCylinderBranch  As Object
    stPoint.Set -2, 0, 0
    enPoint.Set 2, 0, 0
    Set objCylinderHeader = PlaceCylinderTrans(stPoint, enPoint, pr * 2)
    stPoint.Set 0, 0, 0
    enPoint.Set 2 * cosa, 2 * sina, 0
    Set objCylinderBranch = PlaceCylinderTrans(stPoint, enPoint, sr * 2)

    Dim oIJSurfaceHeader As IJSurface
    Dim oIntersectionCurve As IJElements
    Dim IntersectCode As Geom3dIntersectConstants
    Set oIJSurfaceHeader = objCylinderHeader
    oIJSurfaceHeader.Intersect objCylinderBranch, oIntersectionCurve, IntersectCode
  
' Closest distance between Plane and Curve
    Dim oIJcurve As IJCurve
    Dim srcx As Double
    Dim srcy As Double
    Dim srcz As Double
    Dim inx As Double
    Dim iny As Double
    Dim inz As Double
    Set oIJcurve = oIntersectionCurve.Item(1)
    oIJcurve.DistanceBetween objPlane, MiniDistance, srcx, srcy, srcz, inx, iny, inz
        End If
    End If
' RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    'oPlacePoint.Set cosa * MiniDistance, sina * MiniDistance, dComponentOffset
    oPlacePoint.Set 0, 0, dComponentOffset
    
    oDir.Set cosa, sina, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set geomFactory = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub

Private Function PlaceCylinderTrans(lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double) As Object

''' This function creates persistent projetion of circle
''' based on two points (axis of cylinder) and diameter
''' Example of call:
''' Dim stPoint   As new AutoMath.DPosition
''' Dim enPoint   As new AutoMath.DPosition
''' Dim ldiam     as long
''' Dim objCylinder  As object
''' stPoint.set 0, 0, 0
''' enPoint.set 0, 0, 1
''' lDiam = 1.5
''' set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, lDiam, True)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
''' Set objCylinder = Nothing

    Const METHOD = "PlaceCylinderTrans:"
    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.x - lStartPoint.x, _
                     lEndPoint.y - lStartPoint.y, _
                     lEndPoint.z - lStartPoint.z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
    
' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, False)
    
    Set objCircle = Nothing
    
    Set PlaceCylinderTrans = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
        
End Function
'''<{(Cylinder end)}>
