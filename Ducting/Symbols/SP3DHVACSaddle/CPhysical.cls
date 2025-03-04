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
'   ProgID:         SP3DHVACSaddle.HVACSaddle
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'   Physical aspect of HVAC Saddle
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   06.Sep.06       RRK                 TR:104417: Added length parameter (and removed approximation.)
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   20.SEP.2006     dkl/svsmylav        TR-104419: Changed port direction from -z to +z and removed Radial orientation.
'   30.Jan.2008     VRK                 TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private RAD As Double
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
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput As Double
    Dim ObjPad As Object
    Dim parRPadWidth As Double
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parBLength As Double
    Dim parRPadThickness As Double
    Dim parAngle As Double


    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parBLength = arrayOfInputs(4)
    parRPadWidth = parBWidth * 0.2
    parRPadThickness = 0.015
    parAngle = PI / 2

    iOutput = 0

    'Check for pad thickness
    If parBLength < (parRPadThickness + (parWidth / 2)) Then
        parRPadThickness = 0.9 * (parBLength - (parWidth / 2))
    End If

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

    pr = parWidth / 2
    sr = parBWidth / 2

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
    Dim ProjVect As AutoMath.DVector
    Set ProjVect = New AutoMath.DVector

    Set ObjPad = geomFactory.BSplineSurfaces3d.CreateByParameters(m_OutputColl.ResourceManager, _
                                                                  parNumPolesU, parNumPolesV, _
                                                                  p, fWeights, _
                                                                  parUorder, parVorder, _
                                                                  uKnots, vKnots, False, False, False)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPad


    'Transform the Object and get other backing plates
    'Backing Plate 2
    Dim TransMatrix As IJDT4x4
    Dim pVector As AutoMath.DVector
    Dim TransVector As AutoMath.DVector

    Set pVector = New AutoMath.DVector
    Set TransVector = New AutoMath.DVector
    Set TransMatrix = New DT4x4

    TransVector.Set 0, -pr, 0
    TransMatrix.Translate TransVector
    ObjPad.Transform TransMatrix


    pVector.Set 1, 0, 0
    TransMatrix.LoadIdentity
    TransMatrix.Rotate PI / 2, pVector
    ObjPad.Transform TransMatrix

    Set ObjPad = Nothing


    ' Place Nozzle 1
    '=====================================
    'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
    '=====================================
    Dim lEndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim bDimBaseOuter As Boolean
    Dim dPortDepth As Double
    Dim dCptOffset As Double
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double
    Dim dNozzleLength As Double
    Dim dFlangeWidth As Double
    Dim dFlangeThickness As Double
    Dim parPlaneOfBranch As Double
    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    'Set HVAC nozzle parameters
    iPortIndex = 1
    lEndPrep = 301
    dCornerRadius = 0#
    parPlaneOfBranch = 0#

    '    dNozzleLength = parHLength
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    'To construct nozzle as crosssection only, use FlangeWidth of 0, and
    'some non-zero value for flange thickness
    'Flange thickness assumed to be negigible thickness
    dFlangeWidth = 0#
    dFlangeThickness = 0
    dPortDepth = 0#
    dCptOffset = 0#
    FlowDir = DistribFlow_BOTH
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector

    ' Insert your code for output 3(HvacPort2 of Saddle)
    iPortIndex = 1
    FlowDir = DistribFlow_BOTH
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, _
                                                      lEndPrep, dFlangeThickness, dFlangeWidth, FlowDir, parBWidth, _
                                                      parBWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                                      "HvacNozzle1", dPortDepth, dCptOffset, False, m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    Dim dZposition As Double
    dZposition = Sqr((parWidth / 2) ^ 2 - (parBWidth / 2) ^ 2)
    oPos.Set 0, 0, parBLength - parWidth / 2
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos
    dNozzleLength = parBLength - dZposition

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    oDir.Set 0, 0, 1
    oDistribPort.SetDirectionVector oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzleLength

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oRadialDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
Private Function PlaceCylinderTrans(lStartPoint As AutoMath.DPosition, _
                                    lEndPoint As AutoMath.DPosition, _
                                    lDiameter As Double) As Object

    Const METHOD = "PlaceCylinderTrans:"
    On Error GoTo ErrorHandler

    Dim circleCenter As AutoMath.DPosition
    Dim circleNormal As AutoMath.DVector
    Dim objCircle As IngrGeom3D.Circle3d
    Dim dblCylWidth As Double
    Dim objProjection As IngrGeom3D.Projection3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

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
