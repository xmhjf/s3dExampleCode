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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Feb 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' SP3DRightReducerCableTray is placed with larger port (port-1) at y=0 and the Narrow Port(port-2)
' on the positive side of Y-axiz parFacetoFace. The Geometry is formed using three outputs:
' 1. Bottom of Reducer
' 2. Right side plane (Straight portion)
' 3. Projection of Left side curve (Reducing potion)
' Outputs 4 & 5 are Cable Tray Ports created by using 'CreateCableTrayPort' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim parActualWidth    As Double
    Dim parActualDepth    As Double
    Dim parActualWidth2   As Double
    Dim ObjLeftProjection As Object
    Dim ObjReducerBottom  As Object
    Dim ObjRightPlane     As Object
    Dim parFacetoFace As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim HD                 As Double
    Dim HW                 As Double
    Dim HW2                As Double
    Dim Port1S(0 To 11)    As Double
    Dim Port2S(0 To 11)    As Double
    Dim RPlane(0 To 11)    As Double

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parFacetoFace = oTrayPart.Length
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    iOutput = 0

    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Call RetrieveCableTrayPortProperties(2, oPartFclt, parActualWidth2, parActualDepth)
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    HW2 = parActualWidth2 / 2

' Insert your code for output 1(ReducerBottom)
'   Prepare port1 points
    oPort1.Set 0, 0, 0

'   Tray top edge on Negative X-Axis Side
    Port1S(0) = oPort1.x - HW
    Port1S(1) = oPort1.y
    Port1S(2) = oPort1.z + HD

'   Tray bottom on Negative X-Axis Side
    Port1S(3) = oPort1.x - HW
    Port1S(4) = oPort1.y
    Port1S(5) = oPort1.z - HD

'   Tray bottom on Positive X-Axis Side
    Port1S(6) = oPort1.x + HW
    Port1S(7) = oPort1.y
    Port1S(8) = oPort1.z - HD
    
'   Tray top on Positive X-Axis Side
    Port1S(9) = oPort1.x + HW
    Port1S(10) = oPort1.y
    Port1S(11) = oPort1.z + HD

'   Prepare port2 points
    oPort2.Set (oPort1.x + HW - HW2), parFacetoFace, 0

'   Tray top edge at Left side of Reduced end
    Port2S(0) = oPort2.x - HW2
    Port2S(1) = oPort2.y
    Port2S(2) = oPort2.z + HD
    
'   Tray bottom edge at Left side of Reduced end
    Port2S(3) = oPort2.x - HW2
    Port2S(4) = oPort2.y
    Port2S(5) = oPort2.z - HD
    
'   Tray bottom edge at Right side of Reduced end
    Port2S(6) = oPort2.x + HW2
    Port2S(7) = oPort2.y
    Port2S(8) = oPort2.z - HD
    
'   Tray top edge at Right side of Reduced end
    Port2S(9) = oPort2.x + HW2
    Port2S(10) = oPort2.y
    Port2S(11) = oPort2.z + HD

'   Co-ordinate for the Point where the Tray start converging
'   Co-ordinate for the Point where the Tray Stop converging
    Dim ConvergeStartPt(0 To 2) As Double
    Dim ConvergeEndPt(0 To 2) As Double
    ConvergeStartPt(0) = Port1S(3)
    ConvergeStartPt(1) = Port1S(4) + parFacetoFace / 3
    ConvergeStartPt(2) = Port1S(5)
    
    ConvergeEndPt(0) = Port2S(3)
    ConvergeEndPt(1) = Port2S(4) - parFacetoFace / 3
    ConvergeEndPt(2) = Port2S(5)

'   Co-ordinates for right side plane
    RPlane(0) = Port1S(6)
    RPlane(1) = Port1S(7)
    RPlane(2) = Port1S(8)
    
    RPlane(3) = Port2S(6)
    RPlane(4) = Port2S(7)
    RPlane(5) = Port2S(8)
    
    RPlane(6) = Port2S(9)
    RPlane(7) = Port2S(10)
    RPlane(8) = Port2S(11)
    
    RPlane(9) = Port1S(9)
    RPlane(10) = Port1S(10)
    RPlane(11) = Port1S(11)
    
'   Construct Port1 to Port2 Left hand side bottom curve
'   Construct Port1 Left hand side line
    Dim oP1LHSbottomLine As IngrGeom3D.Line3d
    Set oP1LHSbottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), _
                                    ConvergeStartPt(0), ConvergeStartPt(1), ConvergeStartPt(2))
'   Construct Port2 Left hand side line
    Dim oP2LHSbottomLine As IngrGeom3D.Line3d
    Set oP2LHSbottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ConvergeEndPt(0), ConvergeEndPt(1), ConvergeEndPt(2), _
                                    Port2S(3), Port2S(4), Port2S(5))
'   Construct intermediate line joining the above two lines
    Dim oIntermLine As IngrGeom3D.Line3d
    Set oIntermLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ConvergeStartPt(0), ConvergeStartPt(1), ConvergeStartPt(2), _
                                    ConvergeEndPt(0), ConvergeEndPt(1), ConvergeEndPt(2))
'   Construct Port2 Bottom line
    Dim oP2BottomLine As IngrGeom3D.Line3d
    Set oP2BottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), _
                            Port2S(6), Port2S(7), Port2S(8))
'   Construct Port1 to Port2 Right hand side bottom line
    Dim oP1P2RHSbottomLine As IngrGeom3D.Line3d
    Set oP1P2RHSbottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), _
                                    Port1S(6), Port1S(7), Port1S(8))
'   Construct Port1 Bottom line
    Dim oP1BottomLine As IngrGeom3D.Line3d
    Set oP1BottomLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), _
                                        Port1S(3), Port1S(4), Port1S(5))
'   Construct complete bottom boundary curves
    Dim obottomcurves       As Collection
    Dim ObjCBottomCurves    As IngrGeom3D.ComplexString3d
    Set obottomcurves = New Collection
    obottomcurves.Add oP1LHSbottomLine
    obottomcurves.Add oIntermLine
    obottomcurves.Add oP2LHSbottomLine
    obottomcurves.Add oP2BottomLine
    obottomcurves.Add oP1P2RHSbottomLine
    obottomcurves.Add oP1BottomLine
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set ObjCBottomCurves = PlaceTrCString(StartBC, obottomcurves)
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
'   Create the Reducer bottom
    Set ObjReducerBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjReducerBottom.AddBoundary(ObjCBottomCurves)
'   Remove cable tray bottom Header and Branch lines
    Dim ObjBcurves As IJDObject
    Set ObjBcurves = ObjCBottomCurves
    ObjBcurves.Remove
    Set ObjBcurves = Nothing
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducerBottom
    Set ObjReducerBottom = Nothing
    
' Insert your code for output 2(RightPlane)
    Set ObjRightPlane = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, RPlane)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightPlane
    Set ObjRightPlane = Nothing

' Insert your code for output 3(LeftProjection)
    Dim ObjP1P2LHSbottomcurve As IngrGeom3D.ComplexString3d
    Dim oP1P2LHSbottomcurve   As Collection
    Set oP1P2LHSbottomcurve = New Collection
    oP1P2LHSbottomcurve.Add oP1LHSbottomLine
    oP1P2LHSbottomcurve.Add oIntermLine
    oP1P2LHSbottomcurve.Add oP2LHSbottomLine
    Set ObjP1P2LHSbottomcurve = PlaceTrCString(StartBC, oP1P2LHSbottomcurve)
    Set ObjLeftProjection = PlaceProjection(m_OutputColl, ObjP1P2LHSbottomcurve, oDirProj, parActualDepth, True)
    Set ObjP1P2LHSbottomcurve = Nothing
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftProjection
    Set ObjLeftProjection = Nothing
    Set StartBC = Nothing

' Place Nozzle 1
'   Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim ObjCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

'   oPlacePoint.Set -parFacetoFace / 2, 0, 0
    oDir.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set ObjCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 2
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set parFacetoFace / 2, 0, 0
    oDir.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set ObjCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing

'   Remove curves
    Set oP1LHSbottomLine = Nothing
    Set oP2LHSbottomLine = Nothing
    Set oIntermLine = Nothing
    Set oP2BottomLine = Nothing
    Set oP1BottomLine = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To oP1P2LHSbottomcurve.Count
        oP1P2LHSbottomcurve.Remove 1
    Next iCount
    Set oP1P2LHSbottomcurve = Nothing
    
    For iCount = 1 To obottomcurves.Count
        obottomcurves.Remove 1
    Next iCount
    Set obottomcurves = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
''''   Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
