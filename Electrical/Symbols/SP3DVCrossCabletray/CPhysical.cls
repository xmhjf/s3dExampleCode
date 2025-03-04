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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         BBA/MS
'   Creation Date:  Wednesday, Nov 3 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    
    Dim ActualWidth As Double
    Dim ActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    iOutput = 0
'   Retrieve Part properties: BendRadius, TangentLength
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    parBendRadius = oTrayPart.BendRadius
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
'   parTangentLength = 0.0762
    parTangentLength = oTrayPart.TangentLength
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If parTangentLength = 0 Then parTangentLength = 0.000001
'   Retrieve Ports 1: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)

' Insert your code for output 1: CrossFront
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim oPort3 As New AutoMath.DPosition 'Port 3 center point
    Dim oPort4 As New AutoMath.DPosition 'Port 3 center point
    Dim HD                      As Double
    Dim HW                      As Double
    Dim Port1S(0 To 11)         As Double
    Dim Port2S(0 To 11)         As Double
    Dim Port3S(0 To 11)         As Double
    Dim Port4S(0 To 11)         As Double
    Dim BFLA(0 To 8)            As Double

    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

'   Port 1 position
    oPort1.Set -(parTangentLength + parBendRadius + ActualDepth / 2), 0, 0
'   Tray Port 1 U-shape points positions
    HD = ActualDepth / 2
    HW = ActualWidth / 2
'   Front-top
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y - HW
    Port1S(2) = oPort1.z + HD
'   Front-bottom
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y - HW
    Port1S(5) = oPort1.z - HD
'   Back-bottom
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y + HW
    Port1S(8) = oPort1.z - HD
'   Back-top
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y + HW
    Port1S(11) = oPort1.z + HD
    
'   Port 2 position
    oPort2.Set parTangentLength + parBendRadius + ActualDepth / 2, 0, 0
'   Tray Port 2 points positions
'   Front-top
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y - HW
    Port2S(2) = oPort2.z + HD
'   Front-bottom
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y - HW
    Port2S(5) = oPort2.z - HD
'   Back-bottom
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y + HW
    Port2S(8) = oPort2.z - HD
'   Back-top
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y + HW
    Port2S(11) = oPort2.z + HD
    
'   Port 3 position
    oPort3.Set 0, 0, parTangentLength + parBendRadius + ActualDepth / 2
'   Tray Port 3 U-shape points positions
'   Front-Left
    Port3S(0) = oPort3.x - HD
    Port3S(1) = oPort3.y - HW
    Port3S(2) = oPort3.z
'   Front-Right
    Port3S(3) = oPort3.x + HD
    Port3S(4) = oPort3.y - HW
    Port3S(5) = oPort3.z
'   Back-right
    Port3S(6) = oPort3.x + HD
    Port3S(7) = oPort3.y + HW
    Port3S(8) = oPort3.z
'   Back-left
    Port3S(9) = oPort3.x - HD
    Port3S(10) = oPort3.y + HW
    Port3S(11) = oPort3.z
    
'   Port 4 position
    oPort4.Set 0, 0, -(parTangentLength + parBendRadius + ActualDepth / 2)
'   Tray Port 4 U-shape points positions
'   Front-Left
    Port4S(0) = oPort4.x - HD
    Port4S(1) = oPort4.y - HW
    Port4S(2) = oPort4.z
'   Front-Right
    Port4S(3) = oPort4.x + HD
    Port4S(4) = oPort4.y - HW
    Port4S(5) = oPort4.z
'   Back-right
    Port4S(6) = oPort4.x + HD
    Port4S(7) = oPort4.y + HW
    Port4S(8) = oPort4.z
'   Back-left
    Port4S(9) = oPort4.x - HD
    Port4S(10) = oPort4.y + HW
    Port4S(11) = oPort4.z

'   Branch Front Left Arc BFLA
    BFLA(0) = Port1S(0) + parTangentLength
    BFLA(1) = Port1S(1)
    BFLA(2) = Port1S(2)

    BFLA(3) = Port3S(0)
    BFLA(4) = Port3S(1)
    BFLA(5) = Port3S(2) - parTangentLength
    
    BFLA(6) = Port1S(0) + parTangentLength
    BFLA(7) = Port3S(1)
    BFLA(8) = Port3S(2) - parTangentLength

'   Construct Front Cross: PortEdge and Branch curves
'   Construct Port1 Front Edge
    Dim oP1FEdge           As IngrGeom3D.Line3d
    Set oP1FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(0), Port1S(1), Port1S(2), Port1S(3), Port1S(4), Port1S(5))
    
'   Construct front left tangent line from the branch arc start point (bottom of port1)
    Dim oBottomFLTLine           As IngrGeom3D.Line3d
    Set oBottomFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), BFLA(0), BFLA(1), -BFLA(2))

'   Construct branch front left-bottom curve
    Dim objBFLBA  As IngrGeom3D.Arc3d
    Dim SBFLBA   As New AutoMath.DPosition
    Dim EBFLBA   As New AutoMath.DPosition
    Dim CBFLBA   As New AutoMath.DPosition
    SBFLBA.Set BFLA(0), BFLA(1), -BFLA(2)
    EBFLBA.Set BFLA(3), BFLA(4), -BFLA(5)
    CBFLBA.Set BFLA(6), BFLA(7), -BFLA(8)
    Set objBFLBA = PlaceTrArcByCenter(SBFLBA, EBFLBA, CBFLBA)
    Set SBFLBA = Nothing
    Set EBFLBA = Nothing
    Set CBFLBA = Nothing

'   Construct branch front left bottom tangent line from curve end point to port4
    Dim oBFLBLine           As IngrGeom3D.Line3d
    Set oBFLBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), BFLA(4), -BFLA(5), Port4S(0), Port4S(1), Port4S(2))

'   Construct Port4 Front Edge
    Dim oP4FEdge           As IngrGeom3D.Line3d
    Set oP4FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(0), Port4S(1), Port4S(2), Port4S(3), Port4S(4), Port4S(5))
    
'   Construct branch front right bottom tangent line  from port4 to curve end point
    Dim oBFRBLine           As IngrGeom3D.Line3d
    Set oBFRBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(3), Port4S(4), Port4S(5), -BFLA(3), BFLA(4), -BFLA(5))

'   Construct branch front right-bottom curve
    Dim objBFRBA  As IngrGeom3D.Arc3d
    Dim SBFRBA   As New AutoMath.DPosition
    Dim EBFRBA   As New AutoMath.DPosition
    Dim CBFRBA   As New AutoMath.DPosition
    SBFRBA.Set -BFLA(3), BFLA(4), -BFLA(5)
    EBFRBA.Set -BFLA(0), BFLA(1), -BFLA(2)
    CBFRBA.Set -BFLA(6), BFLA(7), -BFLA(8)
    Set objBFRBA = PlaceTrArcByCenter(SBFRBA, EBFRBA, CBFRBA)
    Set SBFRBA = Nothing
    Set EBFRBA = Nothing
    Set CBFRBA = Nothing
    
'   Construct front right tangent line from the branch arc start point (Bottom of port2)
    Dim oBottomFRTLine           As IngrGeom3D.Line3d
    Set oBottomFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), BFLA(1), -BFLA(2), Port2S(3), Port2S(4), Port2S(5))

'   Construct Port2 Front Edge
    Dim oP2FEdge           As IngrGeom3D.Line3d
    Set oP2FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(0), Port2S(1), Port2S(2))

'   Construct front right tangent line till the branch arc start point (Top of port2)
    Dim oTopFRTLine           As IngrGeom3D.Line3d
    Set oTopFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(0), Port2S(1), Port2S(2), -BFLA(0), BFLA(1), BFLA(2))

'   Construct branch front right-top curve
    Dim objBFRTA  As IngrGeom3D.Arc3d
    Dim SBFRTA   As New AutoMath.DPosition
    Dim EBFRTA   As New AutoMath.DPosition
    Dim CBFRTA   As New AutoMath.DPosition
    SBFRTA.Set -BFLA(0), BFLA(1), BFLA(2)
    EBFRTA.Set -BFLA(3), BFLA(4), BFLA(5)
    CBFRTA.Set -BFLA(6), BFLA(7), BFLA(8)
    Set objBFRTA = PlaceTrArcByCenter(SBFRTA, EBFRTA, CBFRTA)
    Set SBFRTA = Nothing
    Set EBFRTA = Nothing
    Set CBFRTA = Nothing

'   Construct branch front right tangent line starting from branch arc end point
    Dim oBFRTLine           As IngrGeom3D.Line3d
    Set oBFRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), BFLA(4), BFLA(5), Port3S(3), Port3S(4), Port3S(5))
    
'   Construct Port3 Front Edge
    Dim oP3FEdge           As IngrGeom3D.Line3d
    Set oP3FEdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(0), Port3S(1), Port3S(2))

'   Construct branch front left top tangent line
    Dim oBFLTLine           As IngrGeom3D.Line3d
    Set oBFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(0), Port3S(1), Port3S(2), BFLA(3), BFLA(4), BFLA(5))

'   Construct branch front left-top curve
    Dim objBFLTA  As IngrGeom3D.Arc3d
    Dim SBFLTA   As New AutoMath.DPosition
    Dim EBFLTA   As New AutoMath.DPosition
    Dim CBFLTA   As New AutoMath.DPosition
    SBFLTA.Set BFLA(3), BFLA(4), BFLA(5)
    EBFLTA.Set BFLA(0), BFLA(1), BFLA(2)
    CBFLTA.Set BFLA(6), BFLA(7), BFLA(8)
    Set objBFLTA = PlaceTrArcByCenter(SBFLTA, EBFLTA, CBFLTA)
    Set SBFLTA = Nothing
    Set EBFLTA = Nothing
    Set CBFLTA = Nothing

'   Construct front left tangent line from the branch arc start point (top of port1)
    Dim oTopFLTLine           As IngrGeom3D.Line3d
    Set oTopFLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), BFLA(1), BFLA(2), Port1S(0), Port1S(1), Port1S(2))

'   Construct entire PortEdge and branch curves
    Dim oFrontBranchcurves           As Collection
    Dim objEBcurves         As IngrGeom3D.ComplexString3d
    Set oFrontBranchcurves = New Collection
    oFrontBranchcurves.Add oP1FEdge
    oFrontBranchcurves.Add oBottomFLTLine
    oFrontBranchcurves.Add objBFLBA
    oFrontBranchcurves.Add oBFLBLine
    oFrontBranchcurves.Add oP4FEdge
    oFrontBranchcurves.Add oBFRBLine
    oFrontBranchcurves.Add objBFRBA
    oFrontBranchcurves.Add oBottomFRTLine
    oFrontBranchcurves.Add oP2FEdge
    oFrontBranchcurves.Add oTopFRTLine
    oFrontBranchcurves.Add objBFRTA
    oFrontBranchcurves.Add oBFRTLine
    oFrontBranchcurves.Add oP3FEdge
    oFrontBranchcurves.Add oBFLTLine
    oFrontBranchcurves.Add objBFLTA
    oFrontBranchcurves.Add oTopFLTLine

'   Construct front-plane
    Dim StartBC As New AutoMath.DPosition
    StartBC.Set Port1S(0), Port1S(1), Port1S(2)
    Set objEBcurves = PlaceTrCString(StartBC, oFrontBranchcurves)
    Dim ObjCrossFront As IngrGeom3D.Plane3d
    Dim oDirProj As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 1, 0
    Set ObjCrossFront = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                    Port1S(0), Port1S(1), Port1S(2), _
                                                                    oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjCrossFront.AddBoundary(objEBcurves)
'   Remove cable tray front-tee Header and Branch lines
    Dim ObjtopEBcurves As IJDObject
    Set ObjtopEBcurves = objEBcurves
    ObjtopEBcurves.Remove
    Set ObjtopEBcurves = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCrossFront
    Set ObjCrossFront = Nothing

' Insert your code for output 2: CrossRear
'   Construct Rear Cross: PortEdge and Branch curves
'   Construct Port1 Rear Edge
    Dim oP1REdge           As IngrGeom3D.Line3d
    Set oP1REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(9), Port1S(10), Port1S(11), Port1S(6), Port1S(7), Port1S(8))
    
'   Construct Rear left tangent line from the branch arc start point (bottom of port1)
    Dim oBottomRLTLine           As IngrGeom3D.Line3d
    Set oBottomRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), BFLA(0), -BFLA(1), -BFLA(2))

'   Construct branch Rear left-bottom curve
    Dim objBRLBA  As IngrGeom3D.Arc3d
    Dim SBRLBA   As New AutoMath.DPosition
    Dim EBRLBA   As New AutoMath.DPosition
    Dim CBRLBA   As New AutoMath.DPosition
    SBRLBA.Set BFLA(0), -BFLA(1), -BFLA(2)
    EBRLBA.Set BFLA(3), -BFLA(4), -BFLA(5)
    CBRLBA.Set BFLA(6), -BFLA(7), -BFLA(8)
    Set objBRLBA = PlaceTrArcByCenter(SBRLBA, EBRLBA, CBRLBA)
    Set SBRLBA = Nothing
    Set EBRLBA = Nothing
    Set CBRLBA = Nothing

'   Construct branch Rear left bottom tangent line
    Dim oBRLBLine           As IngrGeom3D.Line3d
    Set oBRLBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(3), -BFLA(4), -BFLA(5), Port4S(9), Port4S(10), Port4S(11))

'   Construct Port4 Rear Edge
    Dim oP4REdge           As IngrGeom3D.Line3d
    Set oP4REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(9), Port4S(10), Port4S(11), Port4S(6), Port4S(7), Port4S(8))
    
'   Construct branch Rear right bottom tangent line
    Dim oBRRBLine           As IngrGeom3D.Line3d
    Set oBRRBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(6), Port4S(7), Port4S(8), -BFLA(3), -BFLA(4), -BFLA(5))

'   Construct branch Rear right-bottom curve
    Dim objBRRBA  As IngrGeom3D.Arc3d
    Dim SBRRBA   As New AutoMath.DPosition
    Dim EBRRBA   As New AutoMath.DPosition
    Dim CBRRBA   As New AutoMath.DPosition
    SBRRBA.Set -BFLA(3), -BFLA(4), -BFLA(5)
    EBRRBA.Set -BFLA(0), -BFLA(1), -BFLA(2)
    CBRRBA.Set -BFLA(6), -BFLA(7), -BFLA(8)
    Set objBRRBA = PlaceTrArcByCenter(SBRRBA, EBRRBA, CBRRBA)
    Set SBRRBA = Nothing
    Set EBRRBA = Nothing
    Set CBRRBA = Nothing

'   Construct Rear right tangent line from the branch arc start point (Bottom of port2)
    Dim oBottomRRTLine           As IngrGeom3D.Line3d
    Set oBottomRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(0), -BFLA(1), -BFLA(2), Port2S(6), Port2S(7), Port2S(8))

'   Construct Port2 Rear Edge
    Dim oP2REdge           As IngrGeom3D.Line3d
    Set oP2REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), Port2S(9), Port2S(10), Port2S(11))

'   Construct Rear right tangent line till the branch arc start point (Top of port2)
    Dim oTopRRTLine           As IngrGeom3D.Line3d
    Set oTopRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(9), Port2S(10), Port2S(11), -BFLA(0), -BFLA(1), BFLA(2))

'   Construct branch Rear right-top curve
    Dim objBRRTA  As IngrGeom3D.Arc3d
    Dim SBRRTA   As New AutoMath.DPosition
    Dim EBRRTA   As New AutoMath.DPosition
    Dim CBRRTA   As New AutoMath.DPosition
    SBRRTA.Set -BFLA(0), -BFLA(1), BFLA(2)
    EBRRTA.Set -BFLA(3), -BFLA(4), BFLA(5)
    CBRRTA.Set -BFLA(6), -BFLA(7), BFLA(8)
    Set objBRRTA = PlaceTrArcByCenter(SBRRTA, EBRRTA, CBRRTA)
    Set SBRRTA = Nothing
    Set EBRRTA = Nothing
    Set CBRRTA = Nothing

'   Construct branch Rear right tangent line starting from branch arc end point
    Dim oBRRTLine           As IngrGeom3D.Line3d
    Set oBRRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BFLA(3), -BFLA(4), BFLA(5), Port3S(6), Port3S(7), Port3S(8))

'   Construct Port3 Rear Edge
    Dim oP3REdge           As IngrGeom3D.Line3d
    Set oP3REdge = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), Port3S(9), Port3S(10), Port3S(11))

'   Construct branch Rear left top tangent line
    Dim oBRLTLine           As IngrGeom3D.Line3d
    Set oBRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(9), Port3S(10), Port3S(11), BFLA(3), -BFLA(4), BFLA(5))

'   Construct branch Rear left-top curve
    Dim objBRLTA  As IngrGeom3D.Arc3d
    Dim SBRLTA   As New AutoMath.DPosition
    Dim EBRLTA   As New AutoMath.DPosition
    Dim CBRLTA   As New AutoMath.DPosition
    SBRLTA.Set BFLA(3), -BFLA(4), BFLA(5)
    EBRLTA.Set BFLA(0), -BFLA(1), BFLA(2)
    CBRLTA.Set BFLA(6), -BFLA(7), BFLA(8)
    Set objBRLTA = PlaceTrArcByCenter(SBRLTA, EBRLTA, CBRLTA)
    Set SBRLTA = Nothing
    Set EBRLTA = Nothing
    Set CBRLTA = Nothing

'   Construct Rear left tangent line from the branch arc start point (top of port1)
    Dim oTopRLTLine           As IngrGeom3D.Line3d
    Set oTopRLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BFLA(0), -BFLA(1), BFLA(2), Port1S(9), Port1S(10), Port1S(11))

'   Construct entire PortEdge and branch curves
    Dim oRearBranchcurves           As Collection
    Set oRearBranchcurves = New Collection
    oRearBranchcurves.Add oP1REdge
    oRearBranchcurves.Add oBottomRLTLine
    oRearBranchcurves.Add objBRLBA
    oRearBranchcurves.Add oBRLBLine
    oRearBranchcurves.Add oP4REdge
    oRearBranchcurves.Add oBRRBLine
    oRearBranchcurves.Add objBRRBA
    oRearBranchcurves.Add oBottomRRTLine
    oRearBranchcurves.Add oP2REdge
    oRearBranchcurves.Add oTopRRTLine
    oRearBranchcurves.Add objBRRTA
    oRearBranchcurves.Add oBRRTLine
    oRearBranchcurves.Add oP3REdge
    oRearBranchcurves.Add oBRLTLine
    oRearBranchcurves.Add objBRLTA
    oRearBranchcurves.Add oTopRLTLine

'   Construct Rear-plane
    StartBC.Set Port1S(9), Port1S(10), Port1S(11)
    Set objEBcurves = PlaceTrCString(StartBC, oRearBranchcurves)
    Dim ObjCrossRear As IngrGeom3D.Plane3d
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 1, 0
    Set ObjCrossRear = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                    Port1S(9), Port1S(10), Port1S(11), _
                                                                    oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjCrossRear.AddBoundary(objEBcurves)
'   Remove cable tray Rear-tee Header and Branch lines
    Set ObjtopEBcurves = objEBcurves
    ObjtopEBcurves.Remove
    Set ObjtopEBcurves = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCrossRear
    Set ObjCrossRear = Nothing
    
' Insert your code for output 3: TopBranchPort1Side
    Dim iCount As Integer
    StartBC.Set Port3S(0), Port3S(1), Port3S(2)
    Dim objBpscurve         As IngrGeom3D.ComplexString3d
    Dim oBpscurve           As Collection
    Set oBpscurve = New Collection
    oBpscurve.Add oBFLTLine
    oBpscurve.Add objBFLTA
    oBpscurve.Add oTopFLTLine
    Set objBpscurve = PlaceTrCString(StartBC, oBpscurve)
    Dim objTopBranchPort1Side As IJDObject
    Set objTopBranchPort1Side = PlaceProjection(m_OutputColl, objBpscurve, oDirProj, ActualWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopBranchPort1Side
    Set objTopBranchPort1Side = Nothing
    Set objBpscurve = Nothing
    For iCount = 1 To oBpscurve.Count
        oBpscurve.Remove 1
    Next iCount
    Set oBpscurve = Nothing

' Insert your code for output 4: BottomBranchPort1Side
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set oBpscurve = New Collection
    oBpscurve.Add oBottomFLTLine
    oBpscurve.Add objBFLBA
    oBpscurve.Add oBFLBLine
    Set objBpscurve = PlaceTrCString(StartBC, oBpscurve)
    Dim objBottomBranchPort1Side As IJDObject
    Set objBottomBranchPort1Side = PlaceProjection(m_OutputColl, objBpscurve, oDirProj, ActualWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomBranchPort1Side
    Set objBottomBranchPort1Side = Nothing
    Set objBpscurve = Nothing
    For iCount = 1 To oBpscurve.Count
        oBpscurve.Remove 1
    Next iCount
    Set oBpscurve = Nothing

' Insert your code for output 5: TopBranchPort2Side
    StartBC.Set Port2S(0), Port2S(1), Port2S(2)
    Set oBpscurve = New Collection
    oBpscurve.Add oTopFRTLine
    oBpscurve.Add objBFRTA
    oBpscurve.Add oBFRTLine
    Set objBpscurve = PlaceTrCString(StartBC, oBpscurve)
    Dim objTopBranchPort2Side As IJDObject
    Set objTopBranchPort2Side = PlaceProjection(m_OutputColl, objBpscurve, oDirProj, ActualWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopBranchPort2Side
    Set objTopBranchPort2Side = Nothing
    Set objBpscurve = Nothing
    For iCount = 1 To oBpscurve.Count
        oBpscurve.Remove 1
    Next iCount
    Set oBpscurve = Nothing

' Insert your code for output 6: BottomBranchPort2Side
    StartBC.Set Port4S(3), Port4S(4), Port4S(5)
    Set oBpscurve = New Collection
    oBpscurve.Add oBFRBLine
    oBpscurve.Add objBFRBA
    oBpscurve.Add oBottomFRTLine
    Set objBpscurve = PlaceTrCString(StartBC, oBpscurve)
    Dim objBottomBranchPort2Side As IJDObject
    Set objBottomBranchPort2Side = PlaceProjection(m_OutputColl, objBpscurve, oDirProj, ActualWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomBranchPort2Side
    Set objBottomBranchPort2Side = Nothing
    Set oDirProj = Nothing
    Set objBpscurve = Nothing
    For iCount = 1 To oBpscurve.Count
        oBpscurve.Remove 1
    Next iCount
    Set oBpscurve = Nothing
    Set StartBC = Nothing

'   Remove branch curves
    Set oP1FEdge = Nothing
    Set oBottomFLTLine = Nothing
    Set objBFLBA = Nothing
    Set oBFLBLine = Nothing
    Set oP4FEdge = Nothing
    Set oBFRBLine = Nothing
    Set objBFRBA = Nothing
    Set oBottomFRTLine = Nothing
    Set oP2FEdge = Nothing
    Set oTopFRTLine = Nothing
    Set objBFRTA = Nothing
    Set oBFRTLine = Nothing
    Set oP3FEdge = Nothing
    Set oBFLTLine = Nothing
    Set objBFLTA = Nothing
    Set oTopFLTLine = Nothing
    Set oP1REdge = Nothing
    Set oBottomRLTLine = Nothing
    Set objBRLBA = Nothing
    Set oBRLBLine = Nothing
    Set oP4REdge = Nothing
    Set oBRRBLine = Nothing
    Set objBRRBA = Nothing
    Set oBottomRRTLine = Nothing
    Set oP2REdge = Nothing
    Set oTopRRTLine = Nothing
    Set objBRRTA = Nothing
    Set oBRRTLine = Nothing
    Set oP3REdge = Nothing
    Set oBRLTLine = Nothing
    Set objBRLTA = Nothing
    Set oTopRLTLine = Nothing

'    Dim iCount As Integer
    For iCount = 1 To oFrontBranchcurves.Count
        oFrontBranchcurves.Remove 1
    Next iCount
    Set oFrontBranchcurves = Nothing
    
    For iCount = 1 To oRearBranchcurves.Count
        oRearBranchcurves.Remove 1
    Next iCount
    Set oRearBranchcurves = Nothing
    
' Place Port 1
'   Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
'   oPlacePoint.Set -(parBendRadius + ActualWidth / 2), 0, 0
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 2
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set 0, (parBendRadius + ActualWidth / 2), 0
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
' Place Port 3
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set 0, (parBendRadius + ActualWidth / 2), 0
    oDir.Set 0, 0, 1
    oRadialOrient.Set 1, 0, 0
    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oGeomFactory = Nothing
    Set oPortLocation = Nothing
    
' Place Port 4
'   Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
'   oPlacePoint.Set 0, (parBendRadius + ActualWidth / 2), 0
    oDir.Set 0, 0, -1
    oRadialOrient.Set 1, 0, 0
    oPortLocation.Set oPort4.x - dInsertionDepth * oDir.x, oPort4.y - dInsertionDepth * oDir.y, oPort4.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 4, oPortLocation, oDir, oRadialOrient, m_OutputColl)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort4 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


