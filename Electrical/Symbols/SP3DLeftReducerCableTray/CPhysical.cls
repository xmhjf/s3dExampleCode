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
'   Author:          BG
'   Creation Date:  Wednesday, Mar 6 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'SP3DLeftReducerCableTray is placed with larger port (port-1) at y=0 and the Narrow Port(port-2)
'on the Negative side of Y-axiz (-parFacetoFace).The Geometry is formed using three outputs.
'1.Bottom of Tray
'2.Straight Vertical portion
'3.Vertical Portion with Reducer
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
''''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ActualWidth As Double
    Dim ActualDepth As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim HD              As Double
    Dim HW              As Double
    Dim HW2              As Double
    Dim parFacetoFace As Double
    Dim ActualWidth2 As Double
    Dim Port1S(0 To 11) As Double
    Dim Port2S(0 To 11) As Double
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    iOutput = 0
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
    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)
    Call RetrieveCableTrayPortProperties(2, oPartFclt, ActualWidth2, ActualDepth)
    HD = ActualDepth / 2
    HW = ActualWidth / 2
    HW2 = ActualWidth2 / 2
    
' Insert your code for output 1(Bottom Portion of Reducer)
    'Port 1 position
    oPort1.Set 0, 0, 0
    
    'Tray top edge on Negative X-Axis Side
    Port1S(0) = oPort1.x - HW
    Port1S(1) = oPort1.y
    Port1S(2) = oPort1.z + HD
    ' Tray bottom on Negative X-Axis Side
    Port1S(3) = oPort1.x - HW
    Port1S(4) = oPort1.y
    Port1S(5) = oPort1.z - HD
    ' Tray bottom on Positive X-Axis Side
    Port1S(6) = oPort1.x + HW
    Port1S(7) = oPort1.y
    Port1S(8) = oPort1.z - HD
    ' Tray top on Positive X-Axis Side
    Port1S(9) = oPort1.x + HW
    Port1S(10) = oPort1.y
    Port1S(11) = oPort1.z + HD
    
    'Port 2 position
    oPort2.Set oPort1.x + HW - HW2, -parFacetoFace, 0
    'Tray Port 2 U-shape points positions
    'Tray top edge at Left side of Reduced end
    Port2S(0) = oPort2.x - HW2
    Port2S(1) = oPort2.y
    Port2S(2) = oPort2.z + HD
    'Tray bottom edge at Left side of Reduced end
    Port2S(3) = oPort2.x - HW2
    Port2S(4) = oPort2.y
    Port2S(5) = oPort2.z - HD
    'Tray bottom edge at Right side of Reduced end
    Port2S(6) = oPort2.x + HW2
    Port2S(7) = oPort2.y
    Port2S(8) = oPort2.z - HD
    'Tray top edge at Right side of Reduced end
    Port2S(9) = oPort2.x + HW2
    Port2S(10) = oPort2.y
    Port2S(11) = oPort2.z + HD
    
    'Co-ordinate for the Point where the Tray start converging
    'Co-ordinate for the Point where the Tray Stop converging
    Dim ConvergeStartPt(0 To 2) As Double
    Dim ConvergeEndPt(0 To 2) As Double
    ConvergeStartPt(0) = Port1S(3)
    ConvergeStartPt(1) = Port1S(4) - parFacetoFace / 3
    ConvergeStartPt(2) = Port1S(5)
    
    ConvergeEndPt(0) = Port2S(3)
    ConvergeEndPt(1) = Port2S(4) + parFacetoFace / 3
    ConvergeEndPt(2) = Port2S(5)
    'Construct Tray bottom:
    'Construct Line Along Port1 (LAP1)
    Dim oLAP1 As IngrGeom3D.Line3d
    Set oLAP1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), _
                                                                                                    Port1S(3), Port1S(4), Port1S(5))
    'Construct Line on the Left Side Below Port1 (LLSBP1)
    Dim oLLSBP1 As IngrGeom3D.Line3d
    Set oLLSBP1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), _
                                                                ConvergeStartPt(0), ConvergeStartPt(1), ConvergeStartPt(2))
    'Construct Taper Line
    Dim oTaperLine           As IngrGeom3D.Line3d
    Set oTaperLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ConvergeStartPt(0), ConvergeStartPt(1), _
                                        ConvergeStartPt(2), ConvergeEndPt(0), ConvergeEndPt(1), ConvergeEndPt(2))
    'Construct Line on Left Side Above Port2 (LLSAP2)
    Dim oLLSAP2 As IngrGeom3D.Line3d
    Set oLLSAP2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, ConvergeEndPt(0), _
                                        ConvergeEndPt(1), ConvergeEndPt(2), Port2S(3), Port2S(4), Port2S(5))
    'Construct Line along Port2 (LAP2)
    Dim oLAP2 As IngrGeom3D.Line3d
    Set oLAP2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), _
                                                            Port2S(6), Port2S(7), Port2S(8))
    'Construct straight Line on Right Side Between Port1 and Port2 (LRSBP)
    Dim oLRSBP As IngrGeom3D.Line3d
    Set oLRSBP = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), _
                                                            Port1S(6), Port1S(7), Port1S(8))
    
    Dim oReducercurves           As Collection
    Dim objHRcurves         As IngrGeom3D.ComplexString3d
    Set oReducercurves = New Collection
    
    oReducercurves.Add oLAP1
    oReducercurves.Add oLLSBP1
    oReducercurves.Add oTaperLine
    oReducercurves.Add oLLSAP2
    oReducercurves.Add oLAP2
    oReducercurves.Add oLRSBP
                                                            
    Dim StartRC   As New AutoMath.DPosition
    StartRC.Set Port1S(6), Port1S(7), Port1S(8)
    Set objHRcurves = PlaceTrCString(StartRC, oReducercurves)
    Dim ObjTrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set ObjTrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                    Port1S(6), Port1S(7), Port1S(8), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjTrayBottom.AddBoundary(objHRcurves)
    'Remove cable tray bottom Header and Branch lines
    Dim ObjtopHRcurves As IJDObject
    Set ObjtopHRcurves = objHRcurves
    ObjtopHRcurves.Remove
    Set ObjtopHRcurves = Nothing
    Set objHRcurves = Nothing
    Set StartRC = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrayBottom
    Set ObjTrayBottom = Nothing
    
 ' Insert your code for output 2(Straight Vertical Portion)
    Dim VerticalStraight As IngrGeom3D.Plane3d
    Dim VSPoints(0 To 11) As Double
    
    VSPoints(0) = Port1S(9)
    VSPoints(1) = Port1S(10)
    VSPoints(2) = Port1S(11)
    
    VSPoints(3) = Port1S(6)
    VSPoints(4) = Port1S(7)
    VSPoints(5) = Port1S(8)
    
    VSPoints(6) = Port2S(6)
    VSPoints(7) = Port2S(7)
    VSPoints(8) = Port2S(8)
    
    VSPoints(9) = Port2S(9)
    VSPoints(10) = Port2S(10)
    VSPoints(11) = Port2S(11)
    
    Set VerticalStraight = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, VSPoints)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), VerticalStraight
    Set VerticalStraight = Nothing
    
 ' Insert your code for output 3(Taper Vertical Portion)
    Dim objTaperVcurves         As IngrGeom3D.ComplexString3d
    Dim oTaperVcurves           As Collection
    Set oTaperVcurves = New Collection
    oTaperVcurves.Add oLLSBP1
    oTaperVcurves.Add oTaperLine
    oTaperVcurves.Add oLLSAP2
    Dim StTaperCur   As New AutoMath.DPosition
    StTaperCur.Set Port1S(3), Port1S(4), Port1S(5)
    Set objTaperVcurves = PlaceTrCString(StTaperCur, oTaperVcurves)
    Dim ObjVerticalTaper As IJDObject
    Set ObjVerticalTaper = PlaceProjection(m_OutputColl, objTaperVcurves, oDirProj, ActualDepth, True)
    Set objTaperVcurves = Nothing
    Set oDirProj = Nothing
    Set StTaperCur = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerticalTaper
    Set ObjVerticalTaper = Nothing
    
'Remove Branch Curves
    Set oLAP1 = Nothing
    Set oLLSBP1 = Nothing
    Set oTaperLine = Nothing
    Set oLLSAP2 = Nothing
    Set oLAP2 = Nothing
    Set oLRSBP = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To oReducercurves.Count
        oReducercurves.Remove 1
    Next iCount
    Set oReducercurves = Nothing
    For iCount = 1 To oTaperVcurves.Count
        oTaperVcurves.Remove 1
    Next iCount
    Set oTaperVcurves = Nothing
    
' Place Port 1
    'Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    'Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    'oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oPortLocation = Nothing
    
' Place Port 2
    'Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    'oPlacePoint.Set 0, (parBendRadius + parActualWidth / 2), 0
    oDir.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort2 = Nothing
    Set oGeomFactory = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
