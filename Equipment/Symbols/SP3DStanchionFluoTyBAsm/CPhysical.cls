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
'   Creation Date:  MonDay,May 26 2008
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   26.May.2008     VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010     KKC     TR-157988  TO Do List entries are generated whern placing lighting fixtures
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

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parPoleHeight1 As Double
    Dim parPoleHeight2 As Double
    Dim parPoleHeight3 As Double
    Dim parOffset As Double
    Dim parPoleDiameter As Double
    Dim parFixtureLength As Double
    Dim parFixtureWidth As Double
    Dim parFixtureDepth As Double
    Dim parwattage As Double
    Dim parCPx As Double
    Dim parCPy As Double
    Dim parCPz As Double

    Dim dPoleRadius As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim objPole As Object
    Dim objcylinder As Object
    Dim dLinePoints(0 To 32) As Double
    Dim oAxisvec As New AutoMath.DVector
    Dim oTransMatrix As IJDT4x4
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenter = New AutoMath.DPosition
    Set oAxisvec = New AutoMath.DVector
    Set oTransMatrix = New DT4x4

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPoleHeight1 = arrayOfInputs(2)   'P1
    parPoleHeight2 = arrayOfInputs(3)   'P2
    parPoleHeight3 = arrayOfInputs(4)   'P3
    parOffset = arrayOfInputs(5)        'P4
    parPoleDiameter = arrayOfInputs(6)  'PD
    parFixtureLength = arrayOfInputs(7) 'L
    parFixtureWidth = arrayOfInputs(8)  'W
    parFixtureDepth = arrayOfInputs(9)  'D
    parwattage = arrayOfInputs(10)
    
    If UBound(arrayOfInputs) > 10 Then
        parCPx = arrayOfInputs(11)
        parCPy = arrayOfInputs(12)
        parCPz = arrayOfInputs(13)
    End If

    iOutput = 0
    dPoleRadius = parPoleDiameter / 2
    
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
     
    'Vertical Pole 1
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, parPoleHeight1
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPoleDiameter, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Vertical Pole 2
    Dim dDistbetweenCenLines As Double
    dDistbetweenCenLines = parOffset - parPoleDiameter
    oStPoint.Set -dDistbetweenCenLines, 0, parPoleHeight1 + parPoleHeight2
    oEnPoint.Set -dDistbetweenCenLines, 0, parPoleHeight1 + parPoleHeight2 + parPoleHeight3
    Set objPole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parPoleDiameter, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPole
    Set objPole = Nothing
    
    'Create Output 3(Pole Curve Body)
    Dim Surfset As IngrGeom3D.IJElements
    Dim oTraceStr As IngrGeom3D.ComplexString3d
    Dim ObjSurface As Object
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim objCollection As Collection
    Dim oNormal As AutoMath.DVector
    Set oNormal = New AutoMath.DVector
    Set oTraceStr = New ComplexString3d
    Set oLine = New Line3d
    Set oArc = New Arc3d
    
    oCenter.Set 0, 0, parPoleHeight1
    oNormal.Set 0, 0, 1
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dPoleRadius)
                                                 
    Dim oTrace As IngrGeom3D.BSplineCurve3d
    Dim points(1 To 8) As New AutoMath.DPosition
    points(1).Set 0, 0, parPoleHeight1
    points(2).Set 0, 0, points(1).z + 0.22 * parPoleHeight2
    points(3).Set -dDistbetweenCenLines / 5, 0, points(1).z + 0.36 * parPoleHeight2
    points(4).Set -2 * dDistbetweenCenLines / 5, 0, points(1).z + 0.43 * parPoleHeight2
    points(5).Set -3 * dDistbetweenCenLines / 5, 0, points(1).z + 0.57 * parPoleHeight2
    points(6).Set -4 * dDistbetweenCenLines / 5, 0, points(1).z + 0.78 * parPoleHeight2
    points(7).Set -dDistbetweenCenLines, 0, points(1).z + 0.95 * parPoleHeight2
    points(8).Set -dDistbetweenCenLines, 0, points(1).z + parPoleHeight2
    Set oTrace = PlaceTrBspline(8, points)
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTrace, oCircle, _
                                  CircularCorner, 0, stnorm, ednorm, False)
    iOutput = iOutput + 1
    For Each ObjSurface In Surfset
         m_OutputColl.AddOutput "PoleCurveBody", ObjSurface
    Next ObjSurface
    Set Surfset = Nothing
    Set ObjSurface = Nothing
    
    'Cylinder 1
    oStPoint.Set -(dDistbetweenCenLines + dPoleRadius), 0, parPoleHeight1 + parPoleHeight2 + 0.15 * parPoleHeight3
    oEnPoint.Set oStPoint.x + 2 * parPoleDiameter, 0, oStPoint.z
    Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * parPoleDiameter, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcylinder
    Set objcylinder = Nothing

    'Cylinder 2
    oStPoint.Set -(dDistbetweenCenLines + dPoleRadius), 0, parPoleHeight1 + parPoleHeight2 + 0.85 * parPoleHeight3
    oEnPoint.Set oStPoint.x + 2 * parPoleDiameter, 0, oStPoint.z
    Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.1 * parPoleDiameter, True)
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
    oTransMatrix.Rotate PI / 2, oAxisvec
    oTransMatrix.IndexValue(12) = -dDistbetweenCenLines + dPoleRadius + parPoleDiameter + 0.75 * parFixtureDepth
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight1 + parPoleHeight2 + 0.5 * parPoleHeight3
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
    dPoints(1) = 0.5 * parFixtureWidth
    dPoints(2) = -0.05 * parFixtureDepth
    dPoints(3) = 0.495 * parFixtureLength
    dPoints(4) = dPoints(1)
    dPoints(5) = dPoints(2)
    dPoints(6) = dPoints(3)
    dPoints(7) = -0.5 * parFixtureWidth
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
    oTransMatrix.Rotate PI / 2, oAxisvec
    oTransMatrix.IndexValue(12) = -dDistbetweenCenLines + dPoleRadius + parPoleDiameter + 0.45 * parFixtureDepth
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight1 + parPoleHeight2 + 0.5 * parPoleHeight3
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
    oTransMatrix.Rotate PI / 2, oAxisvec
    oTransMatrix.IndexValue(12) = -dDistbetweenCenLines + dPoleRadius + parPoleDiameter + 0.2 * parFixtureDepth
    oTransMatrix.IndexValue(13) = 0
    oTransMatrix.IndexValue(14) = parPoleHeight1 + parPoleHeight2 + 0.5 * parPoleHeight3
    objLight.Transform oTransMatrix
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLight
    Set objLight = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
        oTopStr.RemoveCurve True
    Next iCount
        
    ' Insert your code for output 10(Conduit Port 1)
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
     
    ' Insert your code for output 11(Cable Port 1)
    Dim oCableTrayPortPoint As AutoMath.DPosition
    Set oCableTrayPortPoint = New AutoMath.DPosition
    oCableTrayPortPoint.Set 0, 0, 0
    
    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oRadialOrient.Set 0, -1, 0
    
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
    
    Set oTopStr = Nothing
    Set oCollection = Nothing
    Set oAxisvec = Nothing
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oTransMatrix = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
