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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Monday, Aug 28 2006
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mm.yyyy          who                       change description
'   -----------         -----                        ------------------
'    05.Oct.2006         KKC          TR-106918  New cabletray reducing tee places lopsided.Corrected the ports location.
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
    Dim ObjTrayBottom As Object
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim parActualWidth1 As Double
    Dim parActualWidth2 As Double
    Dim parActualWidth3 As Double
    Dim doffsetvalue As Double
    Dim parTeeBranchLength As Double

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
    parTangentLength = oTrayPart.TangentLength
    parTeeBranchLength = arrayOfInputs(2)
    
'Check to validate that if the tangentLength is zero, set it to a very small value
    If parTangentLength = 0 Then parTangentLength = 0.0001
'   Retrieve Ports 1,2 & 3 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth1, parActualDepth)
    Call RetrieveCableTrayPortProperties(2, oPartFclt, parActualWidth2, parActualDepth)
    Call RetrieveCableTrayPortProperties(3, oPartFclt, parActualWidth3, parActualDepth)
     
    doffsetvalue = (parActualWidth1 / 2 - parActualWidth2 / 2)
 ' Insert your code for output 1: Tray Bottom
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim oPort3 As New AutoMath.DPosition 'Port 2 center point
    Dim HD              As Double
    Dim HW              As Double
    Dim HW3              As Double
    Dim Port1S(0 To 11)  As Double
    Dim Port2S(0 To 11)  As Double
    Dim Port3S(0 To 11)  As Double
    Dim BLA(0 To 17)  As Double
    Dim BRA(0 To 17)  As Double
    Dim BHC(0 To 17)  As Double
    Dim HSide(0 To 11)  As Double

    Dim Angle As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Angle = 2 * Atn(1)
'   Port 1 position
    oPort1.Set -(parTangentLength + parBendRadius + parActualWidth3 / 2), 0, 0
'   Tray Port 1 U-shape points positions
    HD = parActualDepth / 2
    HW = parActualWidth1 / 2
    HW3 = parActualWidth3 / 2
'   Tray top edge close to curved branch
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
'   Tray top edge far from curved branch
    Port1S(9) = oPort1.x
    Port1S(10) = oPort1.y - HW
    Port1S(11) = oPort1.z + HD
    
'   Port 2 position
    oPort2.Set (parTangentLength + parBendRadius + parActualWidth3 / 2), -doffsetvalue, 0
'   Tray Port 2 U-shape points positions
'   Tray top edge close to curved branch
  
    HW = parActualWidth2 / 2
    
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
'   Tray top edge far from curved branch
    Port2S(9) = oPort2.x
    Port2S(10) = oPort2.y - HW
    Port2S(11) = oPort2.z + HD
    
'   Port 3 position
    oPort3.Set 0, (parTeeBranchLength - parActualWidth1 / 2), 0
'   Tray Port 3 U-shape points positions
'   Tray top edge close to curved branch
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
'   Tray top edge far from curved branch
    Port3S(9) = oPort3.x + HW3
    Port3S(10) = oPort3.y
    Port3S(11) = oPort3.z + HD
    
'   Branch Left Arc BLA
    BLA(0) = Port1S(3) + parTangentLength
    BLA(1) = Port1S(4)
    BLA(2) = Port1S(5)
'   Tray bottom
    BLA(3) = Port3S(3)
    BLA(4) = Port1S(4) + parBendRadius
    BLA(5) = Port3S(5)
    
    BLA(6) = Port1S(3) + parTangentLength
    BLA(7) = Port1S(4) + parBendRadius
    BLA(8) = Port1S(5)

'   Branch Right Arc BLA
    BLA(9) = Port2S(3) - parTangentLength
    BLA(10) = Port2S(4)
    BLA(11) = Port2S(5)
'   Tray bottom
    BLA(12) = Port3S(6)
    BLA(13) = Port2S(4) + parBendRadius
    BLA(14) = Port3S(8)
    
    BLA(15) = Port2S(3) - parTangentLength
    BLA(16) = Port2S(4) + parBendRadius
    BLA(17) = Port2S(5)

'   Construct Tray bottom: Header and Branch curves
'   Construct bottom header curve
    Dim oHLine           As IngrGeom3D.Line3d
    Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(6), Port2S(7), Port2S(8), Port1S(6), Port1S(7), Port1S(8))
'   Construct Branch curve
    Dim oBLine           As IngrGeom3D.Line3d
    Set oBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(3), Port3S(4), Port3S(5), Port3S(6), Port3S(7), Port3S(8))
'   Construct Header Left Tangent line
    Dim oLTLine           As IngrGeom3D.Line3d
    Set oLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(3), Port1S(4), Port1S(5), BLA(0), BLA(1), BLA(2))
'   Construct branch left curve
    Dim objBLA  As IngrGeom3D.Arc3d
    Dim SBLA   As New AutoMath.DPosition
    Dim EBLA   As New AutoMath.DPosition
    Dim CBLA   As New AutoMath.DPosition
    SBLA.Set BLA(0), BLA(1), BLA(2)
    EBLA.Set BLA(3), BLA(4), BLA(5)
    CBLA.Set BLA(6), BLA(7), BLA(8)
    Set objBLA = PlaceTrArcByCenter(SBLA, EBLA, CBLA)
'   Construct branch left tangent line
    Dim oBLTLine           As IngrGeom3D.Line3d
    Set oBLTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLA(3), BLA(4), BLA(5), Port3S(3), Port3S(4), Port3S(5))
'   Construct branch right tangent line
    Dim oBRTLine           As IngrGeom3D.Line3d
    Set oBRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port3S(6), Port3S(7), Port3S(8), BLA(12), BLA(13), BLA(14))
'   Construct branch right curve
    Dim objBRA  As IngrGeom3D.Arc3d
    Dim SBRA   As New AutoMath.DPosition
    Dim EBRA   As New AutoMath.DPosition
    Dim CBRA   As New AutoMath.DPosition
    SBRA.Set BLA(12), BLA(13), BLA(14)
    EBRA.Set BLA(9), BLA(10), BLA(11)
    CBRA.Set BLA(15), BLA(16), BLA(17)
    Set objBRA = PlaceTrArcByCenter(SBRA, EBRA, CBRA)
'   Construct Header Right Tangent line
    Dim oRTLine           As IngrGeom3D.Line3d
    Set oRTLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLA(9), BLA(10), BLA(11), Port2S(3), Port2S(4), Port2S(5))
'   Construct Port1 bottom curve
    Dim oPort1botLine           As IngrGeom3D.Line3d
    Set oPort1botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port1S(6), Port1S(7), Port1S(8), Port1S(3), Port1S(4), Port1S(5))
'   Construct Port2 bottom curve
    Dim oPort2botLine           As IngrGeom3D.Line3d
    Set oPort2botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Port2S(3), Port2S(4), Port2S(5), Port2S(6), Port2S(7), Port2S(8))
'   Construct entire Bottom header and branch curves
    Dim obranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Dim objHcurves As IngrGeom3D.ComplexString3d
    Set obranchcurves = New Collection
    obranchcurves.Add oLTLine
    obranchcurves.Add objBLA
    obranchcurves.Add oBLTLine
    obranchcurves.Add oBLine
    obranchcurves.Add oBRTLine
    obranchcurves.Add objBRA
    obranchcurves.Add oRTLine
    obranchcurves.Add oPort2botLine
    obranchcurves.Add oHLine
    obranchcurves.Add oPort1botLine
    
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Port1S(3), Port1S(4), Port1S(5)
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    Dim TrayBottom As IngrGeom3D.Plane3d
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, 1
    Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, Port1S(3), Port1S(4), Port1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call TrayBottom.AddBoundary(objHBcurves)

 'Remove cable tray bottom Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), TrayBottom
    Set TrayBottom = Nothing
    
' Insert your code for output 2: Header side
    Dim HeaderSide As IngrGeom3D.Plane3d
    HSide(0) = Port1S(6)
    HSide(1) = Port1S(7)
    HSide(2) = Port1S(8)
    
    HSide(3) = Port2S(6)
    HSide(4) = Port2S(7)
    HSide(5) = Port2S(8)
    
    HSide(6) = Port2S(9)
    HSide(7) = Port2S(10)
    HSide(8) = Port2S(11)
    
    HSide(9) = Port1S(9)
    HSide(10) = Port1S(10)
    HSide(11) = Port1S(11)

    Set HeaderSide = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, HSide)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), HeaderSide
    Set HeaderSide = Nothing
    
' Insert your code for output 3: BranchPort1Side
    Dim objBp1scurve         As IngrGeom3D.ComplexString3d
    Dim oBp1scurve           As Collection
    Set oBp1scurve = New Collection
    oBp1scurve.Add oLTLine
    oBp1scurve.Add objBLA
    oBp1scurve.Add oBLTLine
    Set objBp1scurve = PlaceTrCString(StartBC, oBp1scurve)
    Dim ObjBranchPort1Side As IJDObject
    Set ObjBranchPort1Side = PlaceProjection(m_OutputColl, objBp1scurve, oDirProj, parActualDepth, True)
    Set objBp1scurve = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchPort1Side
    Set ObjBranchPort1Side = Nothing
    
' Insert your code for output 4: BranchPort2Side
    Dim objBp2scurve         As IngrGeom3D.ComplexString3d
    Dim oBp2scurve           As Collection
    Set oBp2scurve = New Collection
    oBp2scurve.Add oBRTLine
    oBp2scurve.Add objBRA
    oBp2scurve.Add oRTLine
    Dim StartBC2   As New AutoMath.DPosition
    StartBC2.Set Port3S(6), Port3S(7), Port3S(8)
    Set objBp2scurve = PlaceTrCString(StartBC2, oBp2scurve)
    Dim ObjBranchPort2Side As IJDObject
    Set ObjBranchPort2Side = PlaceProjection(m_OutputColl, objBp2scurve, oDirProj, parActualDepth, True)
    Set objBp2scurve = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchPort2Side
    Set ObjBranchPort2Side = Nothing

'   Remove branch curves
    Set oLTLine = Nothing
    Set objBLA = Nothing
    Set oBLTLine = Nothing
    Set oBLine = Nothing
    Set oBRTLine = Nothing
    Set objBRA = Nothing
    Set oRTLine = Nothing
    Set oPort2botLine = Nothing
    Set oHLine = Nothing
    Set oPort1botLine = Nothing

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
    Set oDir = Nothing
    Set oRadialOrient = Nothing
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

    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
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
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
