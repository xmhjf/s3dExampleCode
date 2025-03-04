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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:    VRK
'   Creation Date: WFriday, oct 6 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   23.May.2008     VRK     CR-142762: Provide instrument transmitter and pressure transmitter symbols
'******************************************************************************

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Public Sub run(ByVal m_outputcoll As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim ObjBasePlate As Object
    Dim ObjMountFixture As Object
    Dim ObjInstrumentTube As Object
    Dim ObjTubeCylinder As Object
    
    Dim parBasePlateHeight As Double
    Dim parBasePlateWidth As Double
    Dim parBasePlateLength As Double
    Dim parMountFixtureHeight As Double
    Dim parMountFixtureDiameter As Double
    Dim parInstrumentTubeLength  As Double
    Dim parInstrumentTubeDiameter As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim pBasePlate1  As IJDPosition
    Dim pBasePlate2  As IJDPosition
    Dim oStPoint  As New AutoMath.DPosition
    Dim oEnPoint  As New AutoMath.DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set pBasePlate1 = New DPosition
    Set pBasePlate2 = New DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasePlateHeight = arrayOfInputs(2)
    parBasePlateWidth = arrayOfInputs(3)
    parBasePlateLength = arrayOfInputs(4)
    parMountFixtureHeight = arrayOfInputs(5)
    parMountFixtureDiameter = arrayOfInputs(6)
    parInstrumentTubeLength = arrayOfInputs(7)
    parInstrumentTubeDiameter = arrayOfInputs(8)
       
    iOutput = 0
    
    'Bottom centerof the Baseplate taken as origin
    'Create BasePlate--  Rectangular Solid
    pBasePlate1.Set -parBasePlateLength / 2, -parBasePlateWidth / 2, 0
    pBasePlate2.Set pBasePlate1.x + parBasePlateLength, pBasePlate1.y + parBasePlateWidth, pBasePlate1.z + parBasePlateHeight
    Set ObjBasePlate = PlaceBox(m_outputcoll, pBasePlate1, pBasePlate2)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjBasePlate
    Set ObjBasePlate = Nothing

    'MountFixture ------Cylinder
    oStPoint.Set 0, 0, parBasePlateHeight
    oEnPoint.Set 0, 0, parBasePlateHeight + parMountFixtureHeight
    Set ObjMountFixture = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, parMountFixtureDiameter, False)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjMountFixture
    Set ObjMountFixture = Nothing
        
    'Instrument Tube
    oStPoint.Set -parInstrumentTubeLength / 2, 0, parBasePlateHeight + parMountFixtureHeight
    oEnPoint.Set parInstrumentTubeLength / 2, 0, parBasePlateHeight + parMountFixtureHeight
    Set ObjInstrumentTube = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, parInstrumentTubeDiameter, True)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjInstrumentTube
    Set ObjInstrumentTube = Nothing

    'PipingTubeCylinderDiameter
    Dim Pipedia As Double
    RetrievePipeOD_1 1, oPartFclt, m_outputcoll, Pipedia
  
    Dim dInstrumentTubeRadius As Double
    dInstrumentTubeRadius = parInstrumentTubeDiameter / 2
    oStPoint.Set 0, -dInstrumentTubeRadius, parBasePlateHeight + parMountFixtureHeight + dInstrumentTubeRadius - Pipedia
    oEnPoint.Set 0, dInstrumentTubeRadius, parBasePlateHeight + parMountFixtureHeight + dInstrumentTubeRadius - Pipedia
    Set ObjTubeCylinder = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, Pipedia, False)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjTubeCylinder
    Set ObjTubeCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set pBasePlate1 = Nothing
    Set pBasePlate2 = Nothing
    
    'Create the Edges and Points for Base Plate
    Dim iCount As Integer
    Dim ObjEdgeColl As New Collection
    Dim Points(0 To 23) As Double
    
    Points(0) = -parBasePlateLength / 2
    Points(1) = -parBasePlateWidth / 2
    Points(2) = 0
    
    Points(3) = -parBasePlateLength / 2
    Points(4) = parBasePlateWidth / 2
    Points(5) = 0
    
    Points(6) = -parBasePlateLength / 2
    Points(7) = parBasePlateWidth / 2
    Points(8) = parBasePlateHeight
    
    Points(9) = -parBasePlateLength / 2
    Points(10) = -parBasePlateWidth / 2
    Points(11) = parBasePlateHeight
    
    Points(12) = parBasePlateLength / 2
    Points(13) = -parBasePlateWidth / 2
    Points(14) = 0
    
    Points(15) = parBasePlateLength / 2
    Points(16) = parBasePlateWidth / 2
    Points(17) = 0
    
    Points(18) = parBasePlateLength / 2
    Points(19) = parBasePlateWidth / 2
    Points(20) = parBasePlateHeight
    
    Points(21) = parBasePlateLength / 2
    Points(22) = -parBasePlateWidth / 2
    Points(23) = parBasePlateHeight
    
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(0), Points(1), Points(2), _
                        Points(3), Points(4), Points(5))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(3), Points(4), Points(5), _
                        Points(6), Points(7), Points(8))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(6), Points(7), Points(8), _
                        Points(9), Points(10), Points(11))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(9), Points(10), Points(11), _
                        Points(0), Points(1), Points(2))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(0), Points(1), Points(2), _
                        Points(12), Points(13), Points(14))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(12), Points(13), Points(14), _
                        Points(21), Points(22), Points(23))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(21), Points(22), Points(23), _
                        Points(9), Points(10), Points(11))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(18), Points(19), Points(20), _
                        Points(6), Points(7), Points(8))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(18), Points(19), Points(20), _
                        Points(15), Points(16), Points(17))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(3), Points(4), Points(5), _
                        Points(15), Points(16), Points(17))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(15), Points(16), Points(17), _
                        Points(12), Points(13), Points(14))
    ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputcoll.ResourceManager, _
                        Points(21), Points(22), Points(23), _
                        Points(18), Points(19), Points(20))
    'Set the Output
    For iCount = 1 To ObjEdgeColl.Count
        m_outputcoll.AddOutput "Edges_", ObjEdgeColl(iCount)
    Next iCount
    Set ObjEdgeColl = Nothing
        
    'Create salient points on each face of the body
    Dim ObjPointColl As New Collection
    'Points on Right and Left surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                             0, -parBasePlateWidth / 2, parBasePlateHeight / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                            0, parBasePlateWidth / 2, parBasePlateHeight / 2)
    'Points on Front and Back surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                           -parBasePlateLength / 2, 0, parBasePlateHeight / 2)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                            parBasePlateLength / 2, 0, parBasePlateHeight / 2)
    'Points on Top and Bottom surfaces
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                            0, 0, parBasePlateHeight)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputcoll.ResourceManager, _
                                            0, 0, 0)
    'Set the output
    For iCount = 1 To ObjPointColl.Count
        m_outputcoll.AddOutput "Points_", ObjPointColl(iCount)
    Next iCount
    Set ObjPointColl = Nothing
    Set oGeomFactory = Nothing

Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub
