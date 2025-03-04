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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         kkk & svsmylav
'   Creation Date:  Friday 3, June 2005
'   Description:
'        This class module is the place for user to implement graphical part of VBSymbol for this aspect
'        The symbol is constructed by two ruled surfaces and two planes.
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
    Dim ObjExhGridLeft As Object
    Dim ObjExhGridRight As Object
    Dim ObjExhGridFrontPlane As Object
    Dim ObjExhGridBackPlane As Object

    Dim Width As Double                 ' A= width
    Dim Width1 As Double  ' B=Width1
    Dim Depth1 As Double  ' C=Depth1
    Dim Height1 As Double       ' D=ExhGridHeight

 ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Width1 = arrayOfInputs(3)
    Depth1 = arrayOfInputs(4)
    Height1 = arrayOfInputs(5)

    iOutput = 0

''' ' Insert your code for output 1(ObjExhGridLeft)
    Dim obaseCircle As IngrGeom3D.Arc3d
    Dim otopCircle As IngrGeom3D.Arc3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set otopCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, 0, 0, Width / 2, _
                                                                   0, -Width / 2, 0, _
                                                                    0, 0, -Width / 2)

    Dim oEqpLine As IngrGeom3D.Line3d
    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
    Height1, -(Width1 - Depth1) / 2, -Depth1 / 2, _
    Height1, -(Width1 - Depth1) / 2, Depth1 / 2)

    Set ObjExhGridLeft = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oEqpLine, otopCircle, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExhGridLeft
    Set ObjExhGridLeft = Nothing

' Insert your code for output 2(ObjExhGridRight)
    Set otopCircle = geomFactory.Arcs3d.CreateBy3Points(Nothing, 0, 0, Width / 2, _
                                                            0, Width / 2, 0, _
                                                            0, 0, -Width / 2)

     Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Height1, (Width1 - Depth1) / 2, -Depth1 / 2, _
     Height1, (Width1 - Depth1) / 2, Depth1 / 2)

    Set ObjExhGridRight = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            oEqpLine, otopCircle, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExhGridRight
    Set ObjExhGridRight = Nothing

    Dim sp3dElem1 As IJDObject

    Set sp3dElem1 = otopCircle
    Set obaseCircle = Nothing
    Set otopCircle = Nothing
    sp3dElem1.Remove

  ' Insert your code for output 3(ObjExhGridFrontPlane)
    ReDim lineStrPoints(0 To 8) As Double
    lineStrPoints(0) = 0
    lineStrPoints(1) = 0
    lineStrPoints(2) = -Width / 2

    lineStrPoints(3) = Height1
    lineStrPoints(4) = -(Width1 - Depth1) / 2
    lineStrPoints(5) = -Depth1 / 2

    lineStrPoints(6) = Height1
    lineStrPoints(7) = (Width1 - Depth1) / 2
    lineStrPoints(8) = -Depth1 / 2
    

    Set ObjExhGridFrontPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 3, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExhGridFrontPlane
    Set ObjExhGridFrontPlane = Nothing

' Insert your code for output 4(ObjExhGridBackPlane)
    lineStrPoints(0) = 0
    lineStrPoints(1) = 0
    lineStrPoints(2) = Width / 2

    lineStrPoints(3) = Height1
    lineStrPoints(4) = -(Width1 - Depth1) / 2
    lineStrPoints(5) = Depth1 / 2

    lineStrPoints(6) = Height1
    lineStrPoints(7) = (Width1 - Depth1) / 2
    lineStrPoints(8) = Depth1 / 2

    Set ObjExhGridBackPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 3, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExhGridBackPlane
    Set ObjExhGridBackPlane = Nothing
    
' Insert your code for output 4(ObjExhGridBackPlane)
    ReDim lineStrPoints(0 To 12) As Double
    lineStrPoints(0) = Height1
    lineStrPoints(1) = -(Width1 - Depth1) / 2 - 0.03
    lineStrPoints(2) = -Depth1 / 2 - 0.03
    
    lineStrPoints(3) = Height1
    lineStrPoints(4) = -(Width1 - Depth1) / 2 - 0.03
    lineStrPoints(5) = Depth1 / 2 + 0.03
    
    lineStrPoints(6) = Height1
    lineStrPoints(7) = (Width1 - Depth1) / 2 + 0.03
    lineStrPoints(8) = Depth1 / 2 + 0.03
    
    lineStrPoints(9) = Height1
    lineStrPoints(10) = (Width1 - Depth1) / 2 + 0.03
    lineStrPoints(11) = -Depth1 / 2 - 0.03
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    stPoint.Set Height1, -(Width1 - Depth1) / 2 - 0.03, -Depth1 / 2 - 0.03
    enPoint.Set Height1 + 0.001, (Width1 - Depth1) / 2 + 0.03, Depth1 / 2 + 0.03

    Dim ObjPlateAndSteelNet As Object
    Set ObjPlateAndSteelNet = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlateAndSteelNet
    Set ObjPlateAndSteelNet = Nothing

    Set geomFactory = Nothing


'=====================================
'BUILD HVAC NOZZLE
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
                                            "ExhGridPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)

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

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub


