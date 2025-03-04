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
'   Copyright (c) 2005-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Dec 9, 2005
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Horizontal Expanding/Reducing Cross (HX) symbol. Details of the symbol
'    are taken from Cooper B-Line products at www.b-line.com.
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
    
    Dim parActualWidth As Double
    Dim parActualWidth3 As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
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
    parTangentLength = oTrayPart.TangentLength

    iOutput = 0
    
'   Retrieve Port 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Call RetrieveCableTrayPortProperties(3, oPartFclt, parActualWidth3, parActualDepth)

'   Check to validate that if the tangentLength is zero, set it to a very small value
    If parTangentLength = 0 Then parTangentLength = 0.0001

' Insert your code for output 1: Tray Bottom
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim oPort3 As New AutoMath.DPosition 'Port 3 center point
    Dim oPort4 As New AutoMath.DPosition 'Port 4 center point
    Dim HD              As Double
    Dim HW              As Double
    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim Port3S(0 To 11)  As Double
    Dim Port4S(0 To 11)  As Double
    Dim BLTA(0 To 8)  As Double

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

'   Port 1 position
    oPort1.Set -(parTangentLength + parBendRadius + parActualWidth3 / 2), 0, 0
'   Tray Port 1 U-shape points positions
    HD = parActualDepth / 2
    HW = parActualWidth / 2
'   Tray top edge
    Port1S(0) = oPort1.x
    Port1S(1) = oPort1.y + HW
    Port1S(2) = oPort1.z + HD
'   Tray bottom
    Port1S(3) = oPort1.x
    Port1S(4) = oPort1.y + HW
    Port1S(5) = oPort1.z - HD
    
    Port1S(6) = oPort1.x
    Port1S(7) = oPort1.y - HW
    Port1S(8) = oPort1.z - HD
'   Tray top edge
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    
'   Port 2 position
    oPort2.Set (parTangentLength + parBendRadius + parActualWidth3 / 2), 0, 0
'   Tray Port 2 U-shape points positions
'   Tray top edge
    Port2S(0) = oPort2.x
    Port2S(1) = oPort2.y + HW
    Port2S(2) = oPort2.z + HD

'   Tray bottom
    Port2S(3) = oPort2.x
    Port2S(4) = oPort2.y + HW
    Port2S(5) = oPort2.z - HD
    
    Port2S(6) = oPort2.x
    Port2S(7) = oPort2.y - HW
    Port2S(8) = oPort2.z - HD

'   Tray top edge
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD
    
'   Port 3 position
    oPort3.Set 0, (parTangentLength + parBendRadius + parActualWidth / 2), 0
'   Tray Port 3 U-shape points positions
    Dim HW3 As Double
    
    HW3 = parActualWidth3 / 2

'   Tray top edge
    Port3S(0) = oPort3.x - HW3
    Port3S(1) = oPort3.y
    Port3S(2) = oPort3.z + HD

'   Tray bottom
    Port3S(3) = oPort3.x - HW3
    Port3S(4) = oPort3.y
    Port3S(5) = oPort3.z - HD
    
    Port3S(6) = oPort3.x + HW3
    Port3S(7) = oPort3.y
    Port3S(8) = oPort3.z - HD

'   Tray top edge
    Port3S(9) = oPort3.x + HW3
    Port3S(10) = oPort3.y
    Port3S(11) = oPort3.z + HD
    
'   Port 4 position
    oPort4.Set 0, -(parTangentLength + parBendRadius + parActualWidth / 2), 0
'   Tray Port 4 U-shape points positions
'   Tray top edge
    Port4S(0) = oPort4.x - HW3
    Port4S(1) = oPort4.y
    Port4S(2) = oPort4.z + HD

'   Tray bottom
    Port4S(3) = oPort4.x - HW3
    Port4S(4) = oPort4.y
    Port4S(5) = oPort4.z - HD
    
    Port4S(6) = oPort4.x + HW3
    Port4S(7) = oPort4.y
    Port4S(8) = oPort4.z - HD

'   Tray top edge
    Port4S(9) = oPort4.x + HW3
    Port4S(10) = oPort4.y
    Port4S(11) = oPort4.z + HD

'   Branch  Left Top Arc BLTA
    BLTA(0) = Port1S(3) + parTangentLength
    BLTA(1) = Port1S(4)
    BLTA(2) = Port1S(5)
'   Tray bottom
    BLTA(3) = Port3S(3)
    BLTA(4) = Port3S(4) - parTangentLength
    BLTA(5) = Port3S(5)
    
    BLTA(6) = Port1S(3) + parTangentLength
    BLTA(7) = Port3S(4) - parTangentLength
    BLTA(8) = Port1S(5)

'   Construct Tray bottom: Branch curves
'   Construct Branch curve
    Dim oBLine           As IngrGeom3D.Line3d
    Set oBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Port1 Left Top Tangent line
    Dim oLTopTLine           As IngrGeom3D.Line3d
    Set oLTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), BLTA(0), BLTA(1), BLTA(2))
'   Construct branch left top curve
    Dim objBLTA  As IngrGeom3D.Arc3d
    Dim SBLTA   As New AutoMath.DPosition
    Dim EBLTA   As New AutoMath.DPosition
    Dim CBLTA   As New AutoMath.DPosition
    SBLTA.Set BLTA(0), BLTA(1), BLTA(2)
    EBLTA.Set BLTA(3), BLTA(4), BLTA(5)
    CBLTA.Set BLTA(6), BLTA(7), BLTA(8)
    Set objBLTA = PlaceTrArcByCenter(SBLTA, EBLTA, CBLTA)
'   Construct branch left top tangent line
    Dim oBLTopTLine           As IngrGeom3D.Line3d
    Set oBLTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLTA(3), BLTA(4), BLTA(5), Port3S(3), Port3S(4), Port3S(5))
'   Construct branch right top tangent line
    Dim oBRTopTLine           As IngrGeom3D.Line3d
    Set oBRTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), -BLTA(3), BLTA(4), BLTA(5))
'   Construct branch right top curve
    Dim objBRTA  As IngrGeom3D.Arc3d
    Dim SBRTA   As New AutoMath.DPosition
    Dim EBRTA   As New AutoMath.DPosition
    Dim CBRTA   As New AutoMath.DPosition
    SBRTA.Set -BLTA(3), BLTA(4), BLTA(5)
    EBRTA.Set -BLTA(0), BLTA(1), BLTA(2)
    CBRTA.Set -BLTA(6), BLTA(7), BLTA(8)
    Set objBRTA = PlaceTrArcByCenter(SBRTA, EBRTA, CBRTA)
'   Construct  Right Top Tangent line at Port2
    Dim oRTopTLine           As IngrGeom3D.Line3d
    Set oRTopTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLTA(0), BLTA(1), BLTA(2), Port2S(3), Port2S(4), Port2S(5))
'   Construct Port1 bottom curve
    Dim oPort1botLine           As IngrGeom3D.Line3d
    Set oPort1botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port1S(3), Port1S(4), Port1S(5))
'   Construct Port2 bottom curve
    Dim oPort2botLine           As IngrGeom3D.Line3d
    Set oPort2botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(6), Port2S(7), Port2S(8))
'   Construct Port3 bottom curve
    Dim oPort3botLine           As IngrGeom3D.Line3d
    Set oPort3botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Port4 bottom curve
    Dim oPort4botLine           As IngrGeom3D.Line3d
    Set oPort4botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(6), Port4S(7), Port4S(8), Port4S(3), Port4S(4), Port4S(5))
    
'   Bottom side Curves and Tangents
'   Construct  Right Bottom Tangent line at Port2
    Dim oRBottomTLine           As IngrGeom3D.Line3d
    Set oRBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), -BLTA(0), -BLTA(1), BLTA(2))
'   Construct branch right bottom curve
    Dim objBRightBottomA  As IngrGeom3D.Arc3d
    Dim SBBottomTA   As New AutoMath.DPosition
    Dim EBBottomTA   As New AutoMath.DPosition
    Dim CBBottomTA   As New AutoMath.DPosition
    SBBottomTA.Set -BLTA(0), -BLTA(1), BLTA(2)
    EBBottomTA.Set -BLTA(3), -BLTA(4), BLTA(5)
    CBBottomTA.Set -BLTA(6), -BLTA(7), BLTA(8)
    Set objBRightBottomA = PlaceTrArcByCenter(SBBottomTA, EBBottomTA, CBBottomTA)
'   Construct  Right Bottom Tangent line at Port4
    Dim oBRBottomTLine           As IngrGeom3D.Line3d
    Set oBRBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -BLTA(3), -BLTA(4), BLTA(5), Port4S(6), Port4S(7), Port4S(8))
'   Construct branch left Bottom tangent line
    Dim oBLBottomTLine           As IngrGeom3D.Line3d
    Set oBLBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port4S(3), Port4S(4), Port4S(5), BLTA(3), -BLTA(4), BLTA(5))
'   Construct branch left Bottom curve
    Dim objBLeftBottomA  As IngrGeom3D.Arc3d
    Dim SBLBottomTA   As New AutoMath.DPosition
    Dim EBLBottomTA   As New AutoMath.DPosition
    Dim CBLBottomTA   As New AutoMath.DPosition
    SBLBottomTA.Set BLTA(3), -BLTA(4), BLTA(5)
    EBLBottomTA.Set BLTA(0), -BLTA(1), BLTA(2)
    CBLBottomTA.Set BLTA(6), -BLTA(7), BLTA(8)
    Set objBLeftBottomA = PlaceTrArcByCenter(SBLBottomTA, EBLBottomTA, CBLBottomTA)
'   Construct Port1 Left Bottom Tangent line
    Dim oLBottomTLine           As IngrGeom3D.Line3d
    Set oLBottomTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLTA(0), -BLTA(1), BLTA(2), Port1S(6), Port1S(7), Port1S(8))

'   Construct entire Bottom header and branch curves
    Dim obranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Set obranchcurves = New Collection
    obranchcurves.Add oLTopTLine
    obranchcurves.Add objBLTA
    obranchcurves.Add oBLTopTLine
    obranchcurves.Add oPort3botLine
    obranchcurves.Add oBRTopTLine
    obranchcurves.Add objBRTA
    obranchcurves.Add oRTopTLine
    obranchcurves.Add oPort2botLine
    obranchcurves.Add oRBottomTLine
    obranchcurves.Add objBRightBottomA
    obranchcurves.Add oBRBottomTLine
    obranchcurves.Add oPort4botLine
    obranchcurves.Add oBLBottomTLine
    obranchcurves.Add objBLeftBottomA
    obranchcurves.Add oLBottomTLine
    obranchcurves.Add oPort1botLine
    
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    Dim ObjTrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set ObjTrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call ObjTrayBottom.AddBoundary(objHBcurves)
    'Remove cable tray bottom Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrayBottom
    Set ObjTrayBottom = Nothing
    
' Insert your code for output 2: TopBranchPort1Side
    Dim objBp1scurve         As IngrGeom3D.ComplexString3d
    Dim oBp1scurve           As Collection
    Set oBp1scurve = New Collection
    oBp1scurve.Add oLTopTLine
    oBp1scurve.Add objBLTA
    oBp1scurve.Add oBLTopTLine
    Set objBp1scurve = PlaceTrCString(StartBC, oBp1scurve)
    Dim ObjTopBranchPort1Side As IJDObject
    Set ObjTopBranchPort1Side = PlaceProjection(m_OutputColl, objBp1scurve, oDirProj, parActualDepth, True)
    Set objBp1scurve = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBranchPort1Side
    Set ObjTopBranchPort1Side = Nothing
    
' Insert your code for output 3: TopBranchPort2Side
    Dim objBp2scurve         As IngrGeom3D.ComplexString3d
    Dim oBp2scurve           As Collection
    Set oBp2scurve = New Collection
    oBp2scurve.Add oBRTopTLine
    oBp2scurve.Add objBRTA
    oBp2scurve.Add oRTopTLine
    Dim StartBC2   As New AutoMath.DPosition
    StartBC2.Set Port3S(6), Port3S(7), Port3S(8)
    Set objBp2scurve = PlaceTrCString(StartBC2, oBp2scurve)
    Dim ObjTopBranchPort2Side As IJDObject
    Set ObjTopBranchPort2Side = PlaceProjection(m_OutputColl, objBp2scurve, oDirProj, parActualDepth, True)
    Set objBp2scurve = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBranchPort2Side
    Set ObjTopBranchPort2Side = Nothing
    
' Insert your code for output 4: BottomBranchPort2Side
    Dim objBp3scurve         As IngrGeom3D.ComplexString3d
    Dim oBp3scurve           As Collection
    Set oBp3scurve = New Collection
    oBp3scurve.Add oRBottomTLine
    oBp3scurve.Add objBRightBottomA
    oBp3scurve.Add oBRBottomTLine
    Dim StartBC3   As New AutoMath.DPosition
    StartBC3.Set Port2S(6), Port2S(7), Port2S(8)
    Set objBp3scurve = PlaceTrCString(StartBC3, oBp3scurve)
    Dim ObjBottomBranchPort2Side As IJDObject
    Set ObjBottomBranchPort2Side = PlaceProjection(m_OutputColl, objBp3scurve, oDirProj, parActualDepth, True)
    Set objBp3scurve = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomBranchPort2Side
    Set ObjBottomBranchPort2Side = Nothing

' Insert your code for output 5: BottomBranchPort1Side
    Dim objBp4scurve         As IngrGeom3D.ComplexString3d
    Dim oBp4scurve           As Collection
    Set oBp4scurve = New Collection
    oBp4scurve.Add oBLBottomTLine
    oBp4scurve.Add objBLeftBottomA
    oBp4scurve.Add oLBottomTLine
    
    Dim StartBC4   As New AutoMath.DPosition
    StartBC4.Set Port4S(3), Port4S(4), Port4S(5)
    Set objBp4scurve = PlaceTrCString(StartBC4, oBp4scurve)
    Dim ObjBottomBranchPort1Side As IJDObject
    Set ObjBottomBranchPort1Side = PlaceProjection(m_OutputColl, objBp4scurve, oDirProj, parActualDepth, True)
    Set objBp4scurve = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomBranchPort1Side
    Set ObjBottomBranchPort1Side = Nothing
    
'   Remove branch curves
    Set oLTopTLine = Nothing
    Set objBLTA = Nothing
    Set oBLTopTLine = Nothing
    Set oBLine = Nothing
    Set oBRTopTLine = Nothing
    Set objBRTA = Nothing
    Set oRTopTLine = Nothing
    Set oPort2botLine = Nothing
'   Set oHLine = Nothing
    Set oPort1botLine = Nothing
    Set oRBottomTLine = Nothing
    Set objBRightBottomA = Nothing
    Set oBRBottomTLine = Nothing
    Set oBLBottomTLine = Nothing
    Set objBLeftBottomA = Nothing
    Set oLBottomTLine = Nothing

    Dim iCount As Integer
    For iCount = 1 To obranchcurves.Count
        obranchcurves.Remove 1
    Next iCount
    Set obranchcurves = Nothing
    
    For iCount = 1 To oBp1scurve.Count
        oBp1scurve.Remove 1
    Next iCount
    Set oBp1scurve = Nothing
    
    For iCount = 1 To oBp2scurve.Count
        oBp2scurve.Remove 1
    Next iCount
    Set oBp2scurve = Nothing
    
    For iCount = 1 To oBp3scurve.Count
        oBp3scurve.Remove 1
    Next iCount
    Set oBp3scurve = Nothing
    
    For iCount = 1 To oBp4scurve.Count
        oBp4scurve.Remove 1
    Next iCount
    Set oBp4scurve = Nothing
    
' Place Port 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

    oDir.Set -1, 0, 0
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
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'    Set oPlacePoint = Nothing
    Set oPort2 = Nothing
    Set oPortLocation = Nothing
    
' Place Port 3
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1

    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort3 = Nothing
    Set oPortLocation = Nothing

' Place Port 4
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort4.x - dInsertionDepth * oDir.x, oPort4.y - dInsertionDepth * oDir.y, oPort4.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 4, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort4 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set SBLTA = Nothing
    Set EBLTA = Nothing
    Set CBLTA = Nothing
    Set SBRTA = Nothing
    Set EBRTA = Nothing
    Set CBRTA = Nothing
    Set SBBottomTA = Nothing
    Set EBBottomTA = Nothing
    Set CBBottomTA = Nothing
    Set SBLBottomTA = Nothing
    Set EBLBottomTA = Nothing
    Set CBLBottomTA = Nothing
    Set StartBC = Nothing
    Set oDirProj = Nothing
    Set StartBC2 = Nothing
    Set StartBC3 = Nothing
    Set StartBC4 = Nothing
    Set oTrayPart = Nothing
    Set oPort3botLine = Nothing
    Set oPort4botLine = Nothing
    Set objHBcurves = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
