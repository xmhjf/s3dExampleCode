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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:03 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: RUK
'   Creation Date:  Friday, Feb 19 2007
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
     
    Set m_oSymGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
    
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt  As PartFacelets.IJDPart
    
    Dim iOutput As Double
    Dim ObjEquipment As Object
    Dim ObjDriver As Object
    
    Dim parBaseLength As Double                     'P1
    Dim parBaseWidth1 As Double                     'P2
    Dim parBaseWidth2 As Double                     'P3
    Dim parBaseThickness As Double                  'P4
    Dim parDistBaseLeftToDischarge As Double        'P5
    Dim parDistEqpLeftToDischarge As Double         'P6
    Dim parDistEqpRightToDischarge As Double        'P7
    Dim parEqpWidth1 As Double                      'P8
    Dim parEqpWidth2 As Double                      'P9
    Dim parEqpDepth1 As Double                      'P10
    Dim parEqpDepth2 As Double                      'P11
    Dim parRotEqpLength As Double                   'P12
    Dim parRotEqpDiameter As Double                 'P13
    Dim parDriverLength As Double                   'P14
    Dim parDriverWidth1 As Double                   'P15
    Dim parDriverWidth2 As Double                   'P16
    Dim parDriverHeight As Double                   'P17
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseLength = arrayOfInputs(2)
    parBaseWidth1 = arrayOfInputs(3)
    parBaseWidth2 = arrayOfInputs(4)
    parBaseThickness = arrayOfInputs(5)
    parDistBaseLeftToDischarge = arrayOfInputs(6)
    parDistEqpLeftToDischarge = arrayOfInputs(7)
    parDistEqpRightToDischarge = arrayOfInputs(8)
    parEqpWidth1 = arrayOfInputs(9)
    parEqpWidth2 = arrayOfInputs(10)
    parEqpDepth1 = arrayOfInputs(11)
    parEqpDepth2 = arrayOfInputs(12)
    parRotEqpLength = arrayOfInputs(13)
    parRotEqpDiameter = arrayOfInputs(14)
    parDriverLength = arrayOfInputs(15)
    parDriverWidth1 = arrayOfInputs(16)
    parDriverWidth2 = arrayOfInputs(17)
    parDriverHeight = arrayOfInputs(18)
    parXBoltHole = arrayOfInputs(19)
    parYBoltHole = arrayOfInputs(20)
    parInsulationThickness = arrayOfInputs(21)
        
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Origin is at DP/PP 1(Inside the Equipment Body)
    'Insert your code for output 2(Equipment)
    stPoint.Set -parDistEqpLeftToDischarge, parEqpWidth2, parEqpDepth1
    enPoint.Set parDistEqpRightToDischarge, -parEqpWidth1, -parEqpDepth2
    
    Set ObjEquipment = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEquipment
    Set ObjEquipment = Nothing
    
    'Insert your code for output 3(Driver)
    stPoint.Set parDistEqpRightToDischarge + parRotEqpLength, _
    parDriverWidth2, -parEqpDepth2
    enPoint.Set stPoint.x + parDriverLength, -parDriverWidth1, parDriverHeight - parEqpDepth2
    
    Set ObjDriver = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDriver
    Set ObjDriver = Nothing
    
    'Insert your code for output 4(Rotating Equipment)
    stPoint.Set parDistEqpRightToDischarge, 0, 0
    enPoint.Set stPoint.x + parRotEqpLength, 0, 0
    
    'Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parRotEqpDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    'Insert your code for output 5 to 10(Base plate as planes)
    Dim objColl As New Collection
    Dim dBaseCenter As New AutoMath.DPosition
    
    dBaseCenter.Set parBaseLength / 2, 0, -parEqpDepth2 - parBaseThickness
    Dim oTopPts(0 To 3) As IJDPosition
    Dim oBotPts(0 To 3) As IJDPosition
    
    Dim iCount As Integer
    For iCount = 0 To 3
        Set oTopPts(iCount) = New DPosition
        Set oBotPts(iCount) = New DPosition
    Next iCount
    
    oTopPts(0).Set -parDistBaseLeftToDischarge, dBaseCenter.y - parBaseWidth1, dBaseCenter.z + parBaseThickness
    oTopPts(1).Set oTopPts(0).x + parBaseLength, oTopPts(0).y, oTopPts(0).z
    oTopPts(2).Set oTopPts(1).x, oTopPts(1).y + parBaseWidth1 + parBaseWidth2, oTopPts(0).z
    oTopPts(3).Set oTopPts(0).x, oTopPts(2).y, oTopPts(0).z
    
    oBotPts(0).Set oTopPts(0).x, oTopPts(0).y, dBaseCenter.z
    oBotPts(1).Set oTopPts(1).x, oTopPts(1).y, dBaseCenter.z
    oBotPts(2).Set oTopPts(2).x, oTopPts(2).y, dBaseCenter.z
    oBotPts(3).Set oTopPts(3).x, oTopPts(3).y, dBaseCenter.z
    
    Set objColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopPts, oBotPts)
    
    'Set the output
    For iCount = 1 To objColl.Count
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objColl(iCount)
    Next iCount
    Set objColl = Nothing
    Set dBaseCenter = Nothing
    
    For iCount = 0 To 3
        Set oTopPts(iCount) = Nothing
        Set oBotPts(iCount) = Nothing
    Next iCount
    
    'Adding LinePoints and edges on the Baseplate
    'Create line by two pints:
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim LinePoints(0 To 23) As Double
    Dim objPoint As Object
    Dim line As Object

    LinePoints(0) = -parDistBaseLeftToDischarge
    LinePoints(1) = -parBaseWidth1
    LinePoints(2) = -parEqpDepth2

    LinePoints(3) = -parDistBaseLeftToDischarge + parBaseLength
    LinePoints(4) = -parBaseWidth1
    LinePoints(5) = -parEqpDepth2

    LinePoints(6) = -parDistBaseLeftToDischarge + parBaseLength
    LinePoints(7) = parBaseWidth2
    LinePoints(8) = -parEqpDepth2

    LinePoints(9) = -parDistBaseLeftToDischarge
    LinePoints(10) = parBaseWidth2
    LinePoints(11) = -parEqpDepth2

    LinePoints(12) = -parDistBaseLeftToDischarge
    LinePoints(13) = -parBaseWidth1
    LinePoints(14) = -parEqpDepth2 - parBaseThickness

    LinePoints(15) = -parDistBaseLeftToDischarge + parBaseLength
    LinePoints(16) = -parBaseWidth1
    LinePoints(17) = -parEqpDepth2 - parBaseThickness

    LinePoints(18) = -parDistBaseLeftToDischarge + parBaseLength
    LinePoints(19) = parBaseWidth2
    LinePoints(20) = -parEqpDepth2 - parBaseThickness

    LinePoints(21) = -parDistBaseLeftToDischarge
    LinePoints(22) = parBaseWidth2
    LinePoints(23) = -parEqpDepth2 - parBaseThickness

    'Create salient points on each face of the body
    Dim x, y, z As Double
    'Bottom Point
    
    x = (LinePoints(0) + LinePoints(6)) / 2
    y = (LinePoints(1) + LinePoints(7)) / 2
    z = (LinePoints(2) + LinePoints(8)) / 2

    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    'Top Point
    x = (LinePoints(12) + LinePoints(18)) / 2
    y = (LinePoints(13) + LinePoints(19)) / 2
    z = (LinePoints(14) + LinePoints(20)) / 2

    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Left Face Point
    x = (LinePoints(0) + LinePoints(15)) / 2
    y = (LinePoints(1) + LinePoints(16)) / 2
    z = (LinePoints(2) + LinePoints(17)) / 2

    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Right Face Point
    x = (LinePoints(6) + LinePoints(21)) / 2
    y = (LinePoints(7) + LinePoints(22)) / 2
    z = (LinePoints(8) + LinePoints(23)) / 2
    
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Front Face Point
    x = (LinePoints(0) + LinePoints(21)) / 2
    y = (LinePoints(1) + LinePoints(22)) / 2
    z = (LinePoints(2) + LinePoints(23)) / 2

    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    'Back Face Point
    x = (LinePoints(3) + LinePoints(18)) / 2
    y = (LinePoints(4) + LinePoints(19)) / 2
    z = (LinePoints(5) + LinePoints(20)) / 2

    Set objPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(3), LinePoints(4), LinePoints(5))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(6), LinePoints(7), LinePoints(8), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(9), LinePoints(10), LinePoints(11), LinePoints(0), LinePoints(1), LinePoints(2))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(12), LinePoints(13), LinePoints(14), LinePoints(21), LinePoints(22), LinePoints(23))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(15), LinePoints(16), LinePoints(17), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(18), LinePoints(19), LinePoints(20))
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    Set geomFactory = Nothing
    
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Coordinate system is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |       |
    '   |                  |       |
    '   |                  |       |
    '   -----> X           |       |
    '   Symbol CS          |       |
    '                      |       |
    '                      |   X   |
    '                      |   ^   |
    '                      |   |   |
    '                      |   |   |
    '                      |---|---|----> Y
    '                   Port CS
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parEqpDepth2 - parBaseThickness
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

'''<{(PlaceTrapezoidWithPlanes begin)}>
Private Function PlaceTrapezoidWithPlanes(ByVal objOutputColl As Object, _
                        ByRef topSurfacePoints() As IJDPosition, _
                        ByRef bottomSurfacePoints() As IJDPosition _
                        ) As Collection
''This function returns six Planes of Trapezoid as "Collection". The order of planes in the
''collection is Top,Bottom and four sides. The top and bottom surface points are taken in
''anti-clockwise direction when viewing from top.
    Const METHOD = "PlaceTrapezoidWithPlanes:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    Dim arrayTopPoints(0 To 11) As Double
    Dim arrayBottomPoints(0 To 11) As Double
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim oTmpCollection As New Collection
    
    'These are to obtain the coordinates of various vertices.
    For iCount = 0 To 3
        arrayTopPoints(3 * iCount) = topSurfacePoints(iCount).x
        arrayTopPoints(3 * iCount + 1) = topSurfacePoints(iCount).y
        arrayTopPoints(3 * iCount + 2) = topSurfacePoints(iCount).z
        
        arrayBottomPoints(3 * iCount) = bottomSurfacePoints(iCount).x
        arrayBottomPoints(3 * iCount + 1) = bottomSurfacePoints(iCount).y
        arrayBottomPoints(3 * iCount + 2) = bottomSurfacePoints(iCount).z
    Next iCount
    Set objGeomFactory = New IngrGeom3D.GeometryFactory

    Dim arrayPlanePoints(0 To 11) As Double
    Dim objPlane As IngrGeom3D.Plane3d
    
    'Top plane
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                        4, arrayTopPoints)
    oTmpCollection.Add objPlane
    
    'Bottom plane
    'To orient normal of the bottom plane outside the body, the plane points are taken in reverse order.
    Dim botPtCount As Integer
    For iCount = 1 To 4
        botPtCount = 5 - iCount
        arrayPlanePoints(3 * iCount - 3) = arrayBottomPoints(3 * botPtCount - 3)
        arrayPlanePoints(3 * iCount - 2) = arrayBottomPoints(3 * botPtCount - 2)
        arrayPlanePoints(3 * iCount - 1) = arrayBottomPoints(3 * botPtCount - 1)
    Next iCount
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
    'Front plane
    arrayPlanePoints(0) = arrayBottomPoints(0)
    arrayPlanePoints(1) = arrayBottomPoints(1)
    arrayPlanePoints(2) = arrayBottomPoints(2)
    
    arrayPlanePoints(3) = arrayBottomPoints(3)
    arrayPlanePoints(4) = arrayBottomPoints(4)
    arrayPlanePoints(5) = arrayBottomPoints(5)
    
    arrayPlanePoints(6) = arrayTopPoints(3)
    arrayPlanePoints(7) = arrayTopPoints(4)
    arrayPlanePoints(8) = arrayTopPoints(5)
    
    arrayPlanePoints(9) = arrayTopPoints(0)
    arrayPlanePoints(10) = arrayTopPoints(1)
    arrayPlanePoints(11) = arrayTopPoints(2)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
    'Right hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(3)
    arrayPlanePoints(1) = arrayBottomPoints(4)
    arrayPlanePoints(2) = arrayBottomPoints(5)
    
    arrayPlanePoints(3) = arrayBottomPoints(6)
    arrayPlanePoints(4) = arrayBottomPoints(7)
    arrayPlanePoints(5) = arrayBottomPoints(8)
    
    arrayPlanePoints(6) = arrayTopPoints(6)
    arrayPlanePoints(7) = arrayTopPoints(7)
    arrayPlanePoints(8) = arrayTopPoints(8)
    
    arrayPlanePoints(9) = arrayTopPoints(3)
    arrayPlanePoints(10) = arrayTopPoints(4)
    arrayPlanePoints(11) = arrayTopPoints(5)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
    'Rear plane
    arrayPlanePoints(0) = arrayBottomPoints(6)
    arrayPlanePoints(1) = arrayBottomPoints(7)
    arrayPlanePoints(2) = arrayBottomPoints(8)
    
    arrayPlanePoints(3) = arrayBottomPoints(9)
    arrayPlanePoints(4) = arrayBottomPoints(10)
    arrayPlanePoints(5) = arrayBottomPoints(11)
    
    arrayPlanePoints(6) = arrayTopPoints(9)
    arrayPlanePoints(7) = arrayTopPoints(10)
    arrayPlanePoints(8) = arrayTopPoints(11)
    
    arrayPlanePoints(9) = arrayTopPoints(6)
    arrayPlanePoints(10) = arrayTopPoints(7)
    arrayPlanePoints(11) = arrayTopPoints(8)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
    'Left hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(9)
    arrayPlanePoints(1) = arrayBottomPoints(10)
    arrayPlanePoints(2) = arrayBottomPoints(11)
    
    arrayPlanePoints(3) = arrayBottomPoints(0)
    arrayPlanePoints(4) = arrayBottomPoints(1)
    arrayPlanePoints(5) = arrayBottomPoints(2)
    
    arrayPlanePoints(6) = arrayTopPoints(0)
    arrayPlanePoints(7) = arrayTopPoints(1)
    arrayPlanePoints(8) = arrayTopPoints(2)
    
    arrayPlanePoints(9) = arrayTopPoints(9)
    arrayPlanePoints(10) = arrayTopPoints(10)
    arrayPlanePoints(11) = arrayTopPoints(11)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
    Set PlaceTrapezoidWithPlanes = oTmpCollection

    Set oTmpCollection = Nothing
    Set objGeomFactory = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Function
'''<{(PlaceTrapezoidWithPlanes end)}>



