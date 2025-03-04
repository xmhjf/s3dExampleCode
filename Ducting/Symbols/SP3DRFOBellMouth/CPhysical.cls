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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         NN
'   Creation Date:  Friday 1, Aug 2003
'   Description:
'        This class module is the place for user to implement graphical part of VBSymbol for this aspect
'        The symbol is constructed by two ruled surfaces and two planes.
'        Two arcs (each is half a circle) are constructed. Each ruled surface is between a half circle and the corresponding
'        arc in the flat oval contour. The plane is by three points, and imposed the triangle as boundary.
'        It is a triangle: one point where two half circles meet, and the two other points are the ends of the straight section.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim ObjBellMouthLeft As Object
    Dim ObjBellMouthRight As Object
    Dim ObjBellMouthFrontPlane As Object
    Dim ObjBellMouthBackPlane As Object

    Dim Width As Double
    Dim BellMouthOutletWidth As Double
    Dim BellMouthOutletDepth As Double
    Dim BellMouthHeight As Double

 ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    BellMouthOutletWidth = arrayOfInputs(3)
    BellMouthOutletDepth = arrayOfInputs(4)
    BellMouthHeight = arrayOfInputs(5)

    iOutput = 0

 ' Insert your code for output 1(ObjBellMouthLeft)
    Dim obaseCircle As IngrGeom3D.Arc3d
    Dim otopCircle As IngrGeom3D.Arc3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set otopCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, 0, 0, Width / 2, _
                                                                                                                0, -Width / 2, 0, _
                                                                                                                0, 0, -Width / 2)
    Set obaseCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, BellMouthHeight, -(BellMouthOutletWidth - BellMouthOutletDepth) / 2, BellMouthOutletDepth / 2, _
                                                                                                                 BellMouthHeight, -BellMouthOutletWidth / 2, 0, _
                                                                                                                  BellMouthHeight, -(BellMouthOutletWidth - BellMouthOutletDepth) / 2, -BellMouthOutletDepth / 2)

    Set ObjBellMouthLeft = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            obaseCircle, otopCircle, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouthLeft
    Set ObjBellMouthLeft = Nothing

' Insert your code for output 2(ObjBellMouthRight)
    Set otopCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, 0, 0, Width / 2, _
                                                                                                                0, Width / 2, 0, _
                                                                                                                0, 0, -Width / 2)
    Set obaseCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, BellMouthHeight, (BellMouthOutletWidth - BellMouthOutletDepth) / 2, BellMouthOutletDepth / 2, _
                                                                                                                 BellMouthHeight, BellMouthOutletWidth / 2, 0, _
                                                                                                                  BellMouthHeight, (BellMouthOutletWidth - BellMouthOutletDepth) / 2, -BellMouthOutletDepth / 2)

    Set ObjBellMouthRight = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            obaseCircle, otopCircle, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouthRight
    Set ObjBellMouthRight = Nothing

    Dim sp3dElem1 As IJDObject
    Dim sp3dElem2 As IJDObject

    Set sp3dElem1 = obaseCircle
    Set sp3dElem2 = otopCircle
    Set obaseCircle = Nothing
    Set otopCircle = Nothing
    sp3dElem1.Remove
    sp3dElem2.Remove

  ' Insert your code for output 3(ObjBellMouthFrontPlane)
    Dim lineStrPoints(0 To 8)  As Double
    lineStrPoints(0) = 0
    lineStrPoints(1) = 0
    lineStrPoints(2) = -Width / 2

    lineStrPoints(3) = BellMouthHeight
    lineStrPoints(4) = -(BellMouthOutletWidth - BellMouthOutletDepth) / 2
    lineStrPoints(5) = -BellMouthOutletDepth / 2

    lineStrPoints(6) = BellMouthHeight
    lineStrPoints(7) = (BellMouthOutletWidth - BellMouthOutletDepth) / 2
    lineStrPoints(8) = -BellMouthOutletDepth / 2

    Set ObjBellMouthFrontPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 3, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouthFrontPlane
    Set ObjBellMouthFrontPlane = Nothing

' Insert your code for output 4(ObjBellMouthBackPlane)
    lineStrPoints(0) = 0
    lineStrPoints(1) = 0
    lineStrPoints(2) = Width / 2

    lineStrPoints(3) = BellMouthHeight
    lineStrPoints(4) = -(BellMouthOutletWidth - BellMouthOutletDepth) / 2
    lineStrPoints(5) = BellMouthOutletDepth / 2

    lineStrPoints(6) = BellMouthHeight
    lineStrPoints(7) = (BellMouthOutletWidth - BellMouthOutletDepth) / 2
    lineStrPoints(8) = BellMouthOutletDepth / 2

    Set ObjBellMouthBackPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 3, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBellMouthBackPlane
    Set ObjBellMouthBackPlane = Nothing
    Set geomFactory = Nothing

' Place Bell Mouth Port

'=====================================
'BUILD HVAC NOZZLE: Bell Mouth Port
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim Depth As Double
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim cptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double

'    Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
    PortDepth = 0#
    cptOffset = 0#
    Depth = Width
'     To construct nozzle as crosssection only, use FlangeWidth of 0, and
'     some non-zero value for flange thickness
'     Flange thickness assumed to be negigible thickness

    FlowDir = DistribFlow_IN

'     Depth and Width of crosssection will be the same as diameter of the nozzle which
'     is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "BellPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)

    Set NozzleFactory = Nothing

'    Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos
    Set pos = Nothing


'    Direction specified here of the nozzle should be the direction in which pipe will be routed.
'    Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    Set dir = Nothing

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set sp3dElem1 = Nothing
    Set sp3dElem2 = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

