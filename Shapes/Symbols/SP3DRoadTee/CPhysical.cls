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
'   Author:          NN
'   Creation Date:  Sunday, Aug 03 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim objRoadBottom As IngrGeom3D.Plane3d
    Dim objRoadTop As IngrGeom3D.Plane3d
    Dim objBranchSide1 As IJDObject
    Dim objBranchSide2 As IJDObject
    Dim objTeeCenter As IngrGeom3D.Point3d
    Dim objHeaderPoint1 As IngrGeom3D.Point3d
    Dim objHeaderPoint2 As IngrGeom3D.Point3d
    Dim objBranchPoint3 As IngrGeom3D.Point3d
    Dim ObjPlane1 As Object
    Dim ObjPlane2 As Object
    Dim ObjPlane3 As Object
    
    Dim parRoadWidth As Double
    Dim parRoadDepth As Double
    Dim parRoadRadius As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parRoadWidth = arrayOfInputs(2)
    parRoadDepth = arrayOfInputs(3)
    parRoadRadius = arrayOfInputs(4)
    
    iOutput = 0

    Dim oPoint1 As New AutoMath.DPosition 'Point 1
    Dim oPoint2 As New AutoMath.DPosition 'Point 2
    Dim oPoint3 As New AutoMath.DPosition 'Point 3
    Dim HD              As Double
    Dim HW              As Double
    Dim Point1S(0 To 11)  As Double
    Dim Point2S(0 To 11)  As Double
    Dim Point3S(0 To 11)  As Double
    Dim BLA(0 To 8)  As Double
    Dim BRA(0 To 8)  As Double
'    Dim BHC(0 To 17)  As Double
    Dim HSide(0 To 11)  As Double

    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'   Point 1 position
    oPoint1.Set -(parRoadRadius + parRoadWidth / 2), 0, 0
'   Road Point 1 U-shape points positions
    HD = parRoadDepth / 2
    HW = parRoadWidth / 2
    
'   Road top edge close to curved branch
    Point1S(0) = oPoint1.x
    Point1S(1) = oPoint1.y + HW
    Point1S(2) = oPoint1.z + HD
'   Road bottom
    Point1S(3) = oPoint1.x
    Point1S(4) = oPoint1.y + HW
    Point1S(5) = oPoint1.z - HD
    
    Point1S(6) = oPoint1.x
    Point1S(7) = oPoint1.y - HW
    Point1S(8) = oPoint1.z - HD
'   Road top edge far from curved branch
    Point1S(9) = oPoint1.x
    Point1S(10) = oPoint1.y - HW
    Point1S(11) = oPoint1.z + HD
    
'   Point 2 position
    oPoint2.Set (parRoadRadius + parRoadWidth / 2), 0, 0
'   Road Point 2 U-shape points positions
'   Road top edge close to curved branch
    Point2S(0) = oPoint2.x
    Point2S(1) = oPoint2.y + HW
    Point2S(2) = oPoint2.z + HD
'   Road bottom
    Point2S(3) = oPoint2.x
    Point2S(4) = oPoint2.y + HW
    Point2S(5) = oPoint2.z - HD
    
    Point2S(6) = oPoint2.x
    Point2S(7) = oPoint2.y - HW
    Point2S(8) = oPoint2.z - HD
'   Road top edge far from curved branch
    Point2S(9) = oPoint2.x
    Point2S(10) = oPoint2.y - HW
    Point2S(11) = oPoint2.z + HD
    
'   Point 3 position
    oPoint3.Set 0, (parRoadRadius + parRoadWidth / 2), 0
'   Road Point 3 U-shape points positions
'   Road top edge close to curved branch
    Point3S(0) = oPoint3.x - HW
    Point3S(1) = oPoint3.y
    Point3S(2) = oPoint3.z + HD
'   Road bottom
    Point3S(3) = oPoint3.x - HW
    Point3S(4) = oPoint3.y
    Point3S(5) = oPoint3.z - HD
    
    Point3S(6) = oPoint3.x + HW
    Point3S(7) = oPoint3.y
    Point3S(8) = oPoint3.z - HD
'   Road top edge far from curved branch
    Point3S(9) = oPoint3.x + HW
    Point3S(10) = oPoint3.y
    Point3S(11) = oPoint3.z + HD

' Insert your code for output 1: Plane 1
    Set ObjPlane1 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, Point1S)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane1
    Set ObjPlane1 = Nothing

' Insert your code for output 2: Plane 2
'' dPlanePoints is used to get the Normal of plane outward.
    Dim dPlanePoints(0 To 11) As Double
    
    dPlanePoints(0) = Point2S(0)
    dPlanePoints(1) = Point2S(1)
    dPlanePoints(2) = Point2S(2)
    
    dPlanePoints(3) = Point2S(9)
    dPlanePoints(4) = Point2S(10)
    dPlanePoints(5) = Point2S(11)
    
    dPlanePoints(6) = Point2S(6)
    dPlanePoints(7) = Point2S(7)
    dPlanePoints(8) = Point2S(8)
    
    dPlanePoints(9) = Point2S(3)
    dPlanePoints(10) = Point2S(4)
    dPlanePoints(11) = Point2S(5)
    Set ObjPlane2 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, dPlanePoints)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane2
    Set ObjPlane2 = Nothing
    
' Insert your code for output 3: Plane 3
'' dPlanePoints is used to get the Normal of plane outward.
    dPlanePoints(0) = Point3S(9)
    dPlanePoints(1) = Point3S(10)
    dPlanePoints(2) = Point3S(11)
    
    dPlanePoints(3) = Point3S(6)
    dPlanePoints(4) = Point3S(7)
    dPlanePoints(5) = Point3S(8)
    
    dPlanePoints(6) = Point3S(3)
    dPlanePoints(7) = Point3S(4)
    dPlanePoints(8) = Point3S(5)
    
    dPlanePoints(9) = Point3S(0)
    dPlanePoints(10) = Point3S(1)
    dPlanePoints(11) = Point3S(2)
    Set ObjPlane3 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, dPlanePoints)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane3
    Set ObjPlane3 = Nothing
    
' Insert your code for output 4: Road Bottom (default Surface)
'   Branch Left Arc BLA
    BLA(0) = Point1S(3)
    BLA(1) = Point1S(4)
    BLA(2) = Point1S(5)
'   Road bottom
    BLA(3) = Point3S(3)
    BLA(4) = Point3S(4)
    BLA(5) = Point3S(5)
    
    BLA(6) = Point1S(3)
    BLA(7) = Point3S(4)
    BLA(8) = Point1S(5)

'   Construct Road bottom: Header and Branch curves
'   Construct bottom header curve
    Dim oHLine           As IngrGeom3D.Line3d
    Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point2S(6), Point2S(7), Point2S(8), Point1S(6), Point1S(7), Point1S(8))
'   Construct Branch curve
    Dim oBLine           As IngrGeom3D.Line3d
    Set oBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point3S(3), Point3S(4), Point3S(5), Point3S(6), Point3S(7), Point3S(8))
'   Construct branch left curve
    Dim objBLA  As IngrGeom3D.Arc3d
    Dim SBLA   As New AutoMath.DPosition
    Dim EBLA   As New AutoMath.DPosition
    Dim CBLA   As New AutoMath.DPosition
    SBLA.Set BLA(0), BLA(1), BLA(2)
    EBLA.Set BLA(3), BLA(4), BLA(5)
    CBLA.Set BLA(6), BLA(7), BLA(8)
    Set objBLA = PlaceTrArcByCenter(SBLA, EBLA, CBLA)
'   Construct branch right curve
    Dim objBRA  As IngrGeom3D.Arc3d
    Dim SBRA   As New AutoMath.DPosition
    Dim EBRA   As New AutoMath.DPosition
    Dim CBRA   As New AutoMath.DPosition
    SBRA.Set -BLA(3), BLA(4), BLA(5)
    EBRA.Set -BLA(0), BLA(1), BLA(2)
    CBRA.Set -BLA(6), BLA(7), BLA(8)
    Set objBRA = PlaceTrArcByCenter(SBRA, EBRA, CBRA)
'   Construct Branch inner line
    Dim oBILine           As IngrGeom3D.Line3d
    Set oBILine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, BLA(3), BLA(4), BLA(5), -BLA(3), BLA(4), BLA(5))
'   Construct Point1 bottom curve
    Dim oPoint1botLine           As IngrGeom3D.Line3d
    Set oPoint1botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point1S(6), Point1S(7), Point1S(8), Point1S(3), Point1S(4), Point1S(5))
'   Construct Point2 bottom curve
    Dim oPoint2botLine           As IngrGeom3D.Line3d
    Set oPoint2botLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point2S(3), Point2S(4), Point2S(5), Point2S(6), Point2S(7), Point2S(8))
'   Construct entire Bottom header and branch curves
    Dim obranchcurves           As Collection
    Dim objHBcurves         As IngrGeom3D.ComplexString3d
    Dim objHcurves As IngrGeom3D.ComplexString3d
    Set obranchcurves = New Collection
    obranchcurves.Add objBLA
    obranchcurves.Add oBLine
    obranchcurves.Add objBRA
    obranchcurves.Add oPoint2botLine
    obranchcurves.Add oHLine
    obranchcurves.Add oPoint1botLine
    
    Dim StartBC   As New AutoMath.DPosition
    StartBC.Set Point1S(3), Point1S(4), Point1S(5)
    Set objHBcurves = PlaceTrCString(StartBC, obranchcurves)
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    oDirProj.Set 0, 0, -1
    Set objRoadBottom = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, Point1S(3), Point1S(4), Point1S(5), oDirProj.x, oDirProj.y, oDirProj.z)
    Call objRoadBottom.AddBoundary(objHBcurves)
    'Remove road bottom Header and Branch lines
    Dim ObjbotHBcurves As IJDObject
    Set ObjbotHBcurves = objHBcurves
    ObjbotHBcurves.Remove
    Set ObjbotHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRoadBottom
    Set objRoadBottom = Nothing
 
' Insert your code for output 5: Road Top Plane
' Road Top Plane
    Dim TBLA(0 To 8)  As Double
    Dim TBRA(0 To 8)  As Double

'   Branch Left Arc TBLA
    TBLA(0) = Point1S(3)
    TBLA(1) = Point1S(4)
    TBLA(2) = Point1S(2)
'   Road bottom
    TBLA(3) = Point3S(3)
    TBLA(4) = Point3S(4)
    TBLA(5) = Point3S(2)
    
    TBLA(6) = Point1S(3)
    TBLA(7) = Point3S(4)
    TBLA(8) = Point1S(2)

'   Construct Road bottom: Header and Branch curves
'   Construct bottom header curve
    Dim oTHLine           As IngrGeom3D.Line3d
    Set oTHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point1S(6), Point1S(7), Point1S(11), Point2S(6), Point2S(7), Point2S(11))
'   Construct Branch curve
    Dim oTBLine           As IngrGeom3D.Line3d
    Set oTBLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point3S(6), Point3S(7), Point3S(11), Point3S(3), Point3S(4), Point3S(2))
'   Construct branch left curve
    Dim objTBLA  As IngrGeom3D.Arc3d
    Dim STBLA   As New AutoMath.DPosition
    Dim ETBLA   As New AutoMath.DPosition
    Dim CTBLA   As New AutoMath.DPosition
    ETBLA.Set TBLA(0), TBLA(1), TBLA(2)
    STBLA.Set TBLA(3), TBLA(4), TBLA(5)
    CTBLA.Set TBLA(6), TBLA(7), TBLA(8)
    Set objTBLA = PlaceTrArcByCenter(STBLA, ETBLA, CTBLA)
'   Construct branch right curve
    Dim objTBRA  As IngrGeom3D.Arc3d
    Dim STBRA   As New AutoMath.DPosition
    Dim ETBRA   As New AutoMath.DPosition
    Dim CTBRA   As New AutoMath.DPosition
    ETBRA.Set -TBLA(3), TBLA(4), TBLA(5)
    STBRA.Set -TBLA(0), TBLA(1), TBLA(2)
    CTBRA.Set -TBLA(6), TBLA(7), TBLA(8)
    Set objTBRA = PlaceTrArcByCenter(STBRA, ETBRA, CTBRA)
'   Construct Branch inner line
    Dim oTBILine           As IngrGeom3D.Line3d
    Set oTBILine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -TBLA(3), TBLA(4), TBLA(5), TBLA(3), TBLA(4), TBLA(5))
'   Construct Point1 bottom curve
    Dim oPoint1TopLine           As IngrGeom3D.Line3d
    Set oPoint1TopLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point1S(3), Point1S(4), Point1S(2), Point1S(6), Point1S(7), Point1S(11))
'   Construct Point2 bottom curve
    Dim oPoint2TopLine           As IngrGeom3D.Line3d
    Set oPoint2TopLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point2S(6), Point2S(7), Point2S(11), Point2S(3), Point2S(4), Point2S(2))
'   Construct entire Bottom header and branch curves
    Dim oTbranchcurves           As Collection
    Dim objTHBcurves         As IngrGeom3D.ComplexString3d
    Dim objTHcurves As IngrGeom3D.ComplexString3d
    Set oTbranchcurves = New Collection
    
    oTbranchcurves.Add oPoint1TopLine
    oTbranchcurves.Add oTHLine
    oTbranchcurves.Add oPoint2TopLine
    oTbranchcurves.Add objTBRA
    oTbranchcurves.Add oTBLine
    oTbranchcurves.Add objTBLA
    
    Dim TStartBC   As New AutoMath.DPosition
    TStartBC.Set Point1S(3), Point1S(4), Point1S(2)
    Set objTHBcurves = PlaceTrCString(TStartBC, oTbranchcurves)
    Dim oTDirProj        As AutoMath.DVector
    Set oTDirProj = New AutoMath.DVector
    oTDirProj.Set 0, 0, 1
    Set objRoadTop = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, Point1S(3), Point1S(4), Point1S(2), oTDirProj.x, oTDirProj.y, oTDirProj.z)
    Call objRoadTop.AddBoundary(objTHBcurves)
    'Remove road top Header and Branch lines
    Dim ObjtopHBcurves As IJDObject
    Set ObjtopHBcurves = objTHBcurves
    ObjtopHBcurves.Remove
    Set ObjtopHBcurves = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRoadTop
    Set objRoadTop = Nothing

' Insert your code for output 6: Header side
    Dim HeaderSide As IngrGeom3D.Plane3d
    HSide(0) = Point1S(6)
    HSide(1) = Point1S(7)
    HSide(2) = Point1S(8)
    
    HSide(3) = Point2S(6)
    HSide(4) = Point2S(7)
    HSide(5) = Point2S(8)
    
    HSide(6) = Point2S(9)
    HSide(7) = Point2S(10)
    HSide(8) = Point2S(11)
    
    HSide(9) = Point1S(9)
    HSide(10) = Point1S(10)
    HSide(11) = Point1S(11)

    
    Set HeaderSide = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, HSide)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), HeaderSide
    Set HeaderSide = Nothing
    
' Insert your code for output 7: BranchPoint1Side
    Set objBranchSide1 = PlaceProjection(m_OutputColl, objBLA, oTDirProj, parRoadDepth, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchSide1
    Set objBranchSide1 = Nothing
    
' Insert your code for output 8: BranchPoint2Side
    Set objBranchSide2 = PlaceProjection(m_OutputColl, objBRA, oTDirProj, parRoadDepth, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchSide2
    Set objBranchSide2 = Nothing
    
    
 ' Insert your code for output 9: Road Tee center point
    Dim CenterPos  As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0
 
' Place Tee center
    Set objTeeCenter = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                                                CenterPos.x, CenterPos.y, CenterPos.z)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTeeCenter
    Set objTeeCenter = Nothing
 
 ' Insert your code for output 10
    Set objHeaderPoint1 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                                CenterPos.x - (parRoadWidth / 2 + parRoadRadius), CenterPos.y, CenterPos.z)
     

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderPoint1
    Set objHeaderPoint1 = Nothing
    
 ' Insert your code for output 11
    Set objHeaderPoint2 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                                CenterPos.x + (parRoadWidth / 2 + parRoadRadius), CenterPos.y, CenterPos.z)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHeaderPoint2
    Set objHeaderPoint2 = Nothing
    
 ' Insert your code for output 12
    Set objBranchPoint3 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                                CenterPos.x, CenterPos.y + (parRoadWidth / 2 + parRoadRadius), CenterPos.z)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchPoint3
    Set objBranchPoint3 = Nothing

'   Remove branch curves
    Set objBLA = Nothing
    Set oBLine = Nothing
    Set objBRA = Nothing
    Set oPoint2botLine = Nothing
    Set oHLine = Nothing
    Set oPoint1botLine = Nothing

    Dim iCount As Integer
    For iCount = 1 To obranchcurves.Count
        obranchcurves.Remove 1
    Next iCount
    Set obranchcurves = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


