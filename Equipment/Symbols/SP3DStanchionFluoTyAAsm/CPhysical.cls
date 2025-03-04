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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008-10 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:    VRK
'   Creation Date:  FriDay,May 16 2008
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.May.2008     VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010     VKC     TR-CP-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    PI = Atn(1) * 4

    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Integer

    Dim parPoleHeight As Double
    Dim parFixtureLength As Double
    Dim parPoleDiameter As Double
    Dim parAngle As Double
    Dim parwattage As Double
    Dim parFixtureWidth As Double
    Dim parFixtureDepth As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double

    Dim Surfset As IngrGeom3D.IJElements
    Dim ObjLightFix As Object
    Dim dPoleRadius As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oCollection As Collection
    Dim oCircle As Object
    Dim objPole As Object
    Dim objLight As Object
    Dim objcylinder As Object
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim oNormal As AutoMath.DVector
    Dim oTraceStr As IngrGeom3D.ComplexString3d
    Dim oAxisvec As New AutoMath.DVector
    Dim oTransMatrix As IJDT4x4
    Dim dMMRatio As Double
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oTraceStr = New ComplexString3d
    Set oCollection = New Collection
    Set oLine = New Line3d
    Set oArc = New Arc3d
    Set oNormal = New DVector
    Set oTransMatrix = New DT4x4
    Set oAxisvec = New AutoMath.DVector
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight = arrayOfInputs(2)    'A
    parPoleDiameter = arrayOfInputs(3)  'PD
    parAngle = arrayOfInputs(4)         'C
    parwattage = arrayOfInputs(5)
    parFixtureLength = arrayOfInputs(6) 'L
    parFixtureWidth = arrayOfInputs(7)  'W
    parFixtureDepth = arrayOfInputs(8)  'D

    iOutput = 0
    dPoleRadius = parPoleDiameter / 2
    
    'check for Angle
    If CmpDblGreaterthanOrEqualTo(parAngle, PI / 2) Then GoTo ErrorLabel
    
    'Create the Default Surface at the origin
    'Create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                      0, 0, 0, _
                                                                      0, 0, -1, dPoleRadius)

    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                   (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing
    
    'Vertical Pole
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, 0.8 * parPoleHeight
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPoleDiameter, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Create Output 2(Pole Body)
    'Vertical Line
    oStPoint.Set dPoleRadius, 0, 0.8 * parPoleHeight
    oEnPoint.Set dPoleRadius, 0, parPoleHeight
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     oStPoint.x, oStPoint.y, oStPoint.z, _
                                                     oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Assuming ArcRadius as 1/10 times PoleHeight
    Dim dArcRadius As Double
    dArcRadius = parPoleHeight / 10
    
    oStPoint.Set dPoleRadius, 0, parPoleHeight
    oEnPoint.Set oStPoint.x + dArcRadius * (1 - Cos(PI / 2 - parAngle)), 0, parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle)
    oCenter.Set dPoleRadius + dArcRadius, 0, oStPoint.z
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                          oCenter.x, oCenter.y, oCenter.z, _
                                                          oStPoint.x, oStPoint.y, oStPoint.z, _
                                                          oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line at light
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x + parFixtureLength * Cos(parAngle), 0, oStPoint.z + parFixtureLength * Sin(parAngle)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     oStPoint.x, oStPoint.y, oStPoint.z, _
                                                     oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Circle
    oCenter.Set 0, 0, 0.8 * parPoleHeight
    oNormal.Set 0, 0, 1
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                    oCenter.x, oCenter.y, oCenter.z, _
                                                                    oNormal.x, oNormal.y, oNormal.z, dPoleRadius)

    oStPoint.Set dPoleRadius, 0, 0.8 * parPoleHeight
    Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                  m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                  CircularCorner, 0, stnorm, ednorm, False)
    iOutput = iOutput + 1
    For Each ObjLightFix In Surfset
        m_OutputColl.AddOutput "PoleCurveBody", ObjLightFix
    Next ObjLightFix
    Dim iCount As Integer
    'Remove References
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    For iCount = 1 To oCollection.Count
        oTraceStr.RemoveCurve True
    Next iCount
    Set oTraceStr = Nothing
    
    'Cylinder 1
    oStPoint.Set 0, 0, -0.5 * parPoleDiameter
    oEnPoint.Set oStPoint.x, 0, oStPoint.z + 1.5 * parPoleDiameter
    Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * parPoleDiameter, True)
    oTransMatrix.LoadIdentity
    oAxisvec.Set 0, -1, 0
    oTransMatrix.Rotate parAngle, oAxisvec
    oTransMatrix.IndexValue(12) = dPoleRadius + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + 0.15 * parFixtureLength * Cos(parAngle)
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + 0.15 * parFixtureLength * Sin(parAngle)
    objcylinder.Transform oTransMatrix
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcylinder
    Set objcylinder = Nothing

    'Cylinder 2
    oStPoint.Set 0, 0, -0.5 * parPoleDiameter
    oEnPoint.Set oStPoint.x, 0, oStPoint.z + 1.5 * parPoleDiameter
    Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * parPoleDiameter, True)
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oAxisvec.Set 0, -1, 0
    oTransMatrix.Rotate parAngle, oAxisvec
    oTransMatrix.IndexValue(12) = dPoleRadius + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + 0.85 * parFixtureLength * Cos(parAngle)
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + 0.85 * parFixtureLength * Sin(parAngle)
    objcylinder.Transform oTransMatrix
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcylinder
    Set objcylinder = Nothing
    
    'Create Output for Light
    Dim dBottomPoints(0 To 17) As Double

    dBottomPoints(0) = -0.495 * parFixtureLength
    dBottomPoints(1) = -0.2 * parFixtureWidth
    dBottomPoints(2) = -0.25 * parFixtureDepth
    dBottomPoints(3) = dBottomPoints(0)
    dBottomPoints(4) = -0.45 * parFixtureWidth
    dBottomPoints(5) = -0.05 * parFixtureDepth
    dBottomPoints(6) = dBottomPoints(0)
    dBottomPoints(7) = -0.45 * parFixtureWidth
    dBottomPoints(8) = 0.25 * parFixtureDepth
    dBottomPoints(9) = dBottomPoints(0)
    dBottomPoints(10) = 0.45 * parFixtureWidth
    dBottomPoints(11) = 0.25 * parFixtureDepth
    dBottomPoints(12) = dBottomPoints(0)
    dBottomPoints(13) = 0.45 * parFixtureWidth
    dBottomPoints(14) = -0.05 * parFixtureDepth
    dBottomPoints(15) = dBottomPoints(0)
    dBottomPoints(16) = 0.2 * parFixtureWidth
    dBottomPoints(17) = -0.25 * parFixtureDepth
    Set oCollection = New Collection
    
    'Line 1
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dBottomPoints(9), dBottomPoints(10), dBottomPoints(11), _
                                                     dBottomPoints(12), dBottomPoints(13), dBottomPoints(14))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Arc 1
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                          dBottomPoints(12), dBottomPoints(13) - 0.2 * parFixtureDepth, dBottomPoints(14), _
                                                          dBottomPoints(12), dBottomPoints(13), dBottomPoints(14), _
                                                          dBottomPoints(15), dBottomPoints(13) - 0.2 * parFixtureDepth, dBottomPoints(17))
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line 2
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dBottomPoints(15), dBottomPoints(13) - 0.2 * parFixtureDepth, dBottomPoints(17), _
                                                     dBottomPoints(0), dBottomPoints(4) + 0.2 * parFixtureDepth, dBottomPoints(2))

    oCollection.Add oLine
    Set oLine = Nothing
   
    'Arc 2
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                          dBottomPoints(0), dBottomPoints(4) + 0.2 * parFixtureDepth, dBottomPoints(5), _
                                                          dBottomPoints(0), dBottomPoints(4) + 0.2 * parFixtureDepth, dBottomPoints(2), _
                                                          dBottomPoints(3), dBottomPoints(4), dBottomPoints(5))
    oCollection.Add oArc
    Set oArc = Nothing
    
    'Line 3
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dBottomPoints(3), dBottomPoints(4), dBottomPoints(5), _
                                                     dBottomPoints(6), dBottomPoints(7), dBottomPoints(8))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 4
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dBottomPoints(6), dBottomPoints(7), dBottomPoints(8), _
                                                     dBottomPoints(9), dBottomPoints(10), dBottomPoints(11))
    oCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dBottomPoints(9), dBottomPoints(10), dBottomPoints(11)
    Dim oBottomTraceStr As IngrGeom3D.ComplexString3d
    Set oBottomTraceStr = New ComplexString3d
    Set oBottomTraceStr = PlaceTrCString(oStPoint, oCollection)
    oAxisvec.Set 1, 0, 0
    Set objLight = PlaceProjection(m_OutputColl, oBottomTraceStr, oAxisvec, 0.99 * parFixtureLength, True)
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oAxisvec.Set 0, -1, 0
    oTransMatrix.Rotate parAngle, oAxisvec
    oTransMatrix.IndexValue(12) = dPoleRadius + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + 0.5 * parFixtureLength * Cos(parAngle) + 0.5 * parPoleDiameter * Sin(parAngle) + 0.75 * parFixtureDepth * Sin(parAngle)
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + 0.5 * parFixtureLength * Sin(parAngle) - 0.75 * parFixtureDepth * Cos(parAngle) - 0.5 * parPoleDiameter * Cos(parAngle)
    objLight.Transform oTransMatrix
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLight
    Set objLight = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
        oBottomTraceStr.RemoveCurve True
    Next iCount
    Set oBottomTraceStr = Nothing
    Set oCollection = Nothing
    
    Dim dPoints(0 To 11) As Double
    Dim oEllipArc As New IngrGeom3D.EllipticalArc3d
    dPoints(0) = -0.495 * parFixtureLength
    dPoints(1) = -0.5 * parFixtureWidth
    dPoints(2) = -0.05 * parFixtureDepth
    dPoints(3) = 0.495 * parFixtureLength
    dPoints(4) = dPoints(1)
    dPoints(5) = dPoints(2)
    dPoints(6) = dPoints(3)
    dPoints(7) = 0.5 * parFixtureWidth
    dPoints(8) = dPoints(2)
    dPoints(9) = dPoints(0)
    dPoints(10) = dPoints(7)
    dPoints(11) = dPoints(2)
    
    Set oCollection = New Collection
    
    'Line 1
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dPoints(0), dPoints(1), dPoints(2), _
                                                     dPoints(3), dPoints(4), dPoints(5))

    oCollection.Add oLine
    Set oLine = Nothing
   
    'Arc 1
    dMMRatio = (0.005 * parFixtureLength) / (0.5 * parFixtureWidth)
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                        dPoints(3), 0, dPoints(5), _
                                                                                        0, 0, -1, 0, 0.5 * parFixtureWidth, 0, dMMRatio, 0, PI)
    oCollection.Add oEllipArc
    Set oEllipArc = Nothing
    
    'Line 2
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dPoints(6), dPoints(7), dPoints(8), _
                                                     dPoints(9), dPoints(10), dPoints(11))
    oCollection.Add oLine
    Set oLine = Nothing

    'Arc 2
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                        dPoints(9), 0, dPoints(11), _
                                                                                        0, 0, -1, 0, -0.5 * parFixtureWidth, 0, dMMRatio, 0, PI)
    oCollection.Add oEllipArc
    Set oEllipArc = Nothing
    
    oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
    Dim oLineStr As IngrGeom3D.ComplexString3d
    Set oLineStr = New ComplexString3d
    Set oLineStr = PlaceTrCString(oStPoint, oCollection)
    oAxisvec.Set 0, 0, 1
    Set objLight = PlaceProjection(m_OutputColl, oLineStr, oAxisvec, 0.1 * parFixtureDepth, True)
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oAxisvec.Set 0, -1, 0
    oTransMatrix.Rotate parAngle, oAxisvec
    oTransMatrix.IndexValue(12) = dPoleRadius + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + 0.5 * parFixtureLength * Cos(parAngle) + 0.5 * parPoleDiameter * Sin(parAngle) + 0.45 * parFixtureDepth * Sin(parAngle)
    oTransMatrix.IndexValue(13) = parFixtureWidth
    oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + 0.5 * parFixtureLength * Sin(parAngle) - 0.45 * parFixtureDepth * Cos(parAngle) - 0.5 * parPoleDiameter * Cos(parAngle)
    objLight.Transform oTransMatrix
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLight
    Set objLight = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
        oLineStr.RemoveCurve True
    Next iCount
    Set oLineStr = Nothing
    Set oCollection = Nothing
    
    'Top Points
    Dim dTopPoints(0 To 23) As Double
    Dim oTopStr As New IngrGeom3D.ComplexString3d
    dTopPoints(0) = -0.495 * parFixtureLength
    dTopPoints(1) = 0.46 * parFixtureWidth
    dTopPoints(2) = -0.2 * parFixtureDepth
    
    dTopPoints(3) = dTopPoints(0)
    dTopPoints(4) = -0.46 * parFixtureWidth
    dTopPoints(5) = dTopPoints(2)
    
    dTopPoints(6) = dTopPoints(0)
    dTopPoints(7) = dTopPoints(4)
    dTopPoints(8) = -0.1 * parFixtureDepth
    
    dTopPoints(9) = dTopPoints(0)
    dTopPoints(10) = -0.25 * parFixtureWidth
    dTopPoints(11) = dTopPoints(8)
    
    dTopPoints(12) = dTopPoints(0)
    dTopPoints(13) = -0.15 * parFixtureWidth
    dTopPoints(14) = 0.2 * parFixtureDepth
    
    dTopPoints(15) = dTopPoints(0)
    dTopPoints(16) = 0.15 * parFixtureWidth
    dTopPoints(17) = dTopPoints(14)
    
    dTopPoints(18) = dTopPoints(0)
    dTopPoints(19) = 0.25 * parFixtureWidth
    dTopPoints(20) = dTopPoints(8)
    
    dTopPoints(21) = dTopPoints(0)
    dTopPoints(22) = dTopPoints(1)
    dTopPoints(23) = dTopPoints(8)
    Set oCollection = New Collection
    
    'Line 1
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(0), dTopPoints(1), dTopPoints(2), _
                                                     dTopPoints(3), dTopPoints(4), dTopPoints(5))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 2
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(3), dTopPoints(4), dTopPoints(5), _
                                                     dTopPoints(6), dTopPoints(7), dTopPoints(8))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 3
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(6), dTopPoints(7), dTopPoints(8), _
                                                     dTopPoints(9), dTopPoints(10), dTopPoints(11))
    oCollection.Add oLine
    Set oLine = Nothing
   
    'Arc 1
    dMMRatio = parFixtureWidth / (3 * parFixtureDepth)
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                        dTopPoints(9), dTopPoints(13), dTopPoints(11), _
                                                                                        -1, 0, 0, 0, 0, 0.3 * parFixtureDepth, dMMRatio, -PI / 2, PI / 2)
    oCollection.Add oEllipArc
    Set oEllipArc = Nothing

    'Line 4
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(12), dTopPoints(13), dTopPoints(14), _
                                                     dTopPoints(15), dTopPoints(16), dTopPoints(17))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Arc 2
    Set oEllipArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                        dTopPoints(15), dTopPoints(16), dTopPoints(20), _
                                                                                        -1, 0, 0, 0, 0, 0.3 * parFixtureDepth, dMMRatio, 0, PI / 2)
    oCollection.Add oEllipArc
    Set oEllipArc = Nothing

    'Line 5
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(18), dTopPoints(19), dTopPoints(20), _
                                                     dTopPoints(21), dTopPoints(22), dTopPoints(23))
    oCollection.Add oLine
    Set oLine = Nothing
    
    'Line 6
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                     dTopPoints(21), dTopPoints(22), dTopPoints(23), _
                                                     dTopPoints(0), dTopPoints(1), dTopPoints(2))
    oCollection.Add oLine
    Set oLine = Nothing
    
    oStPoint.Set dTopPoints(0), dTopPoints(1), dTopPoints(2)
    Set oTopStr = PlaceTrCString(oStPoint, oCollection)
    oAxisvec.Set 1, 0, 0
    Set objLight = PlaceProjection(m_OutputColl, oTopStr, oAxisvec, 0.99 * parFixtureLength, True)
    Set oTransMatrix = New DT4x4
    oTransMatrix.LoadIdentity
    oAxisvec.Set 0, -1, 0
    oTransMatrix.Rotate parAngle, oAxisvec
    oTransMatrix.IndexValue(12) = dPoleRadius + dArcRadius * (1 - Cos(PI / 2 - parAngle)) + 0.5 * parFixtureLength * Cos(parAngle) + 0.5 * parPoleDiameter * Sin(parAngle) + 0.2 * parFixtureDepth * Sin(parAngle)
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight + dArcRadius * Sin(PI / 2 - parAngle) + 0.5 * parFixtureLength * Sin(parAngle) - 0.2 * parFixtureDepth * Cos(parAngle) - 0.5 * parPoleDiameter * Cos(parAngle)
    objLight.Transform oTransMatrix
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLight
    Set objLight = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
        oTopStr.RemoveCurve True
    Next iCount
    Set oTopStr = Nothing
    Set oCollection = Nothing
    Set oAxisvec = Nothing
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing
    Set oTransMatrix = Nothing
    Set oCircle = Nothing

    ' Insert your code for output 9(Conduit Port 1)
    Dim oConduitPortPoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oConduitPortPoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oDir.Set 0, 0, -1

    oConduitPortPoint.Set 0, 0, 0
    Dim ObjConduitPort1 As IJConduitPortOcc
    Set ObjConduitPort1 = CreateConduitNozzlePH(oConduitPortPoint, oDir, m_OutputColl, oPartFclt, 1)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing

    ' Insert your code for output 10(Cable Port 2)
    Dim oCableTrayPortPoint As AutoMath.DPosition
    Set oCableTrayPortPoint = New AutoMath.DPosition

    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oRadialOrient.Set 0, -1, 0
    oCableTrayPortPoint.Set 0, 0, 0

    Dim objCableNozzle As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory

    Set objCableNozzle = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                                      m_OutputColl.ResourceManager)

    Set iLogicalDistPort = objCableNozzle
    iLogicalDistPort.SetCenterLocation oCableTrayPortPoint
    Set iDistribPort = objCableNozzle
    iDistribPort.SetDirectionVector oDir

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableNozzle
    Set objCableNozzle = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub
