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
'   Author:         ACM
'   Creation Date:  Monday, Aug 04 2003
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Ingr SmartPlant 3D Symbol Road Cross.
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
 
    Dim objRoadCrossBottom As IngrGeom3D.Plane3d
    Dim objRoadCrossTop As IngrGeom3D.Plane3d
    Dim ObjPlane1 As Object
    Dim ObjPlane2 As Object
    Dim ObjPlane3 As Object
    Dim ObjPlane4 As Object
    Dim ObjTurnSurface1 As Object
    Dim ObjTurnSurface2 As Object
    Dim ObjTurnSurface3 As Object
    Dim ObjTurnSurface4 As Object
    Dim objCrossCenter  As Object
    Dim objPoint1 As IngrGeom3D.Point3d
    Dim objPoint2 As IngrGeom3D.Point3d
    Dim objPoint3 As IngrGeom3D.Point3d
    Dim objPoint4 As IngrGeom3D.Point3d
 
    Dim parRoadWidth As Double
    Dim parRoadDepth As Double
    Dim parRoadRadius As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parRoadWidth = arrayOfInputs(2)
    parRoadDepth = arrayOfInputs(3)
    parRoadRadius = arrayOfInputs(4)
    
    iOutput = 0
 
 ''Construct Side planes
 
    Dim oPoint As New AutoMath.DPosition 'Point 1
    Dim HD              As Double
    Dim HW              As Double
    Dim Point1S(0 To 11)  As Double
    Dim Point2S(0 To 11)  As Double
    Dim Point3S(0 To 11)  As Double
    Dim Point4S(0 To 11)  As Double
    Dim dArcPoints(0 To 8)  As Double

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'   Point 1 position
    oPoint.Set -(parRoadRadius + parRoadWidth / 2), 0, 0
'   Road Point 1 U-shape points positions
    HD = parRoadDepth / 2
    HW = parRoadWidth / 2
'   Road top edge close to curved branch
    Point1S(0) = oPoint.x
    Point1S(1) = oPoint.y + HW
    Point1S(2) = oPoint.z + HD
'   Road bottom
    Point1S(3) = oPoint.x
    Point1S(4) = oPoint.y + HW
    Point1S(5) = oPoint.z - HD
    
    Point1S(6) = oPoint.x
    Point1S(7) = oPoint.y - HW
    Point1S(8) = oPoint.z - HD
'   Road top edge far from curved branch
    Point1S(9) = oPoint.x
    Point1S(10) = oPoint.y - HW
    Point1S(11) = oPoint.z + HD
    
'   Point 2 position
    oPoint.Set 0, -(parRoadRadius + parRoadWidth / 2), 0
'   Road Point 2 U-shape points positions
'   Road top edge close to curved branch
    Point2S(0) = oPoint.x + HW
    Point2S(1) = oPoint.y
    Point2S(2) = oPoint.z + HD
'   Road bottom
    Point2S(3) = oPoint.x - HW
    Point2S(4) = oPoint.y
    Point2S(5) = oPoint.z + HD
    
    Point2S(6) = oPoint.x - HW
    Point2S(7) = oPoint.y
    Point2S(8) = oPoint.z - HD
'   Road top edge far from curved branch
    Point2S(9) = oPoint.x + HW
    Point2S(10) = oPoint.y
    Point2S(11) = oPoint.z - HD
    
'   Point 3 position
    oPoint.Set (parRoadRadius + parRoadWidth / 2), 0, 0
'   Road Point 3 U-shape points positions
'   Road top edge close to curved branch
    Point3S(0) = oPoint.x
    Point3S(1) = oPoint.y + HW
    Point3S(2) = oPoint.z + HD
'   Road bottom
    Point3S(3) = oPoint.x
    Point3S(4) = oPoint.y - HW
    Point3S(5) = oPoint.z + HD
    
    Point3S(6) = oPoint.x
    Point3S(7) = oPoint.y - HW
    Point3S(8) = oPoint.z - HD
'   Road top edge far from curved branch
    Point3S(9) = oPoint.x
    Point3S(10) = oPoint.y + HW
    Point3S(11) = oPoint.z - HD
    
'   Point 4 position
    oPoint.Set 0, (parRoadRadius + parRoadWidth / 2), 0
'   Road Point 2 U-shape points positions
'   Road top edge close to curved branch
    Point4S(0) = oPoint.x + HW
    Point4S(1) = oPoint.y
    Point4S(2) = oPoint.z + HD
'   Road bottom
    Point4S(3) = oPoint.x + HW
    Point4S(4) = oPoint.y
    Point4S(5) = oPoint.z - HD
    
    Point4S(6) = oPoint.x - HW
    Point4S(7) = oPoint.y
    Point4S(8) = oPoint.z - HD
'   Road top edge far from curved branch
    Point4S(9) = oPoint.x - HW
    Point4S(10) = oPoint.y
    Point4S(11) = oPoint.z + HD
    
    
 ' Insert your code for output 1: Plane 1
    Set ObjPlane1 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, Point1S)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane1
    Set ObjPlane1 = Nothing
    
' Insert your code for output 2: Plane 2
    Set ObjPlane2 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, Point2S)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane2
    Set ObjPlane2 = Nothing
    
' Insert your code for output 3: Plane 3
    Set ObjPlane3 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, Point3S)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane3
    Set ObjPlane3 = Nothing
    
' Insert your code for output 4: Plane 4
    Set ObjPlane4 = oGeomFactory.Planes3d.CreateByPoints(Nothing, 4, Point4S)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlane4
    Set ObjPlane4 = Nothing
    
''Construct Cross Corner Curved Portions
    Dim oArc  As IngrGeom3D.Arc3d
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEndPoint   As New AutoMath.DPosition
    Dim oCenterPos   As New AutoMath.DPosition
    Dim oLine As IngrGeom3D.Line3d
    Dim oCollection As Collection
    Set oCollection = New Collection
    Dim oDirProj        As AutoMath.DVector
    Set oDirProj = New AutoMath.DVector
    
'   Top Left Corner Arc TLCA
    dArcPoints(0) = Point1S(3)
    dArcPoints(1) = Point1S(4)
    dArcPoints(2) = Point1S(5)
'   Road bottom
    dArcPoints(3) = Point4S(6)
    dArcPoints(4) = Point4S(7)
    dArcPoints(5) = Point4S(8)
    'Arc Center point
    dArcPoints(6) = Point1S(3)
    dArcPoints(7) = Point4S(7)
    dArcPoints(8) = Point1S(5)
    
'   Construct Top left Corner curve
    oStPoint.Set dArcPoints(0), dArcPoints(1), dArcPoints(2)
    oEndPoint.Set dArcPoints(3), dArcPoints(4), dArcPoints(5)
    oCenterPos.Set dArcPoints(6), dArcPoints(7), dArcPoints(8)
    
    Set oArc = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenterPos)
 
    ''Adding to Collection
    oCollection.Add oArc

' Insert your code for output 5: Corner Curve 1
    oDirProj.Set 0, 0, 1
    Set ObjTurnSurface1 = PlaceProjection(m_OutputColl, oArc, oDirProj, parRoadDepth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurnSurface1
    Set ObjTurnSurface1 = Nothing
    Set oArc = Nothing
    
'   Construct Horizontal line at Point 4
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point4S(6), Point4S(7), Point4S(8), _
                                                            Point4S(3), Point4S(4), Point4S(5))
    ''Adding to Collection
    oCollection.Add oLine
    Set oLine = Nothing

''   Top right Corner Arc TRCA
    oStPoint.Set -dArcPoints(3), dArcPoints(4), dArcPoints(5)
    oEndPoint.Set -dArcPoints(0), dArcPoints(1), dArcPoints(2)
    oCenterPos.Set -dArcPoints(6), dArcPoints(7), dArcPoints(8)
    
    Set oArc = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenterPos)

   ''Adding to Collection
    oCollection.Add oArc

' Insert your code for output 6: Corner Curve 2
    oDirProj.Set 0, 0, 1
    Set ObjTurnSurface2 = PlaceProjection(m_OutputColl, oArc, oDirProj, parRoadDepth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurnSurface2
    Set ObjTurnSurface2 = Nothing
    Set oArc = Nothing

'  Construct Horizontal line at Point 3
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point3S(9), Point3S(10), Point3S(11), _
                                                              Point3S(6), Point3S(7), Point3S(8))

   ''Adding to Collection
    oCollection.Add oLine
    Set oLine = Nothing

''   Bottom Right Corner Arc TLCA
'   Construct Bottom Right Corner curve
    oStPoint.Set -dArcPoints(0), -dArcPoints(1), dArcPoints(2)
    oEndPoint.Set -dArcPoints(3), -dArcPoints(4), dArcPoints(5)
    oCenterPos.Set -dArcPoints(6), -dArcPoints(7), dArcPoints(8)
    
    Set oArc = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenterPos)

   ''Adding to Collection
    oCollection.Add oArc
    
' Insert your code for output 7: Corner Curve 3
    oDirProj.Set 0, 0, 1
    Set ObjTurnSurface3 = PlaceProjection(m_OutputColl, oArc, oDirProj, parRoadDepth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurnSurface3
    Set ObjTurnSurface3 = Nothing
    Set oArc = Nothing

'   Construct Horizontal line at Point 2
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point2S(9), Point2S(10), Point2S(11), _
                                                             Point2S(6), Point2S(7), Point2S(8))

   ''Adding to Collection
    oCollection.Add oLine
    Set oLine = Nothing
    
''   Bottom Left Corner Arc BLCA
    
    oStPoint.Set dArcPoints(3), -dArcPoints(4), dArcPoints(5)
    oEndPoint.Set dArcPoints(0), -dArcPoints(1), dArcPoints(2)
    oCenterPos.Set dArcPoints(6), -dArcPoints(7), dArcPoints(8)

    Set oArc = PlaceTrArcByCenter(oStPoint, oEndPoint, oCenterPos)
    
    ''Adding to Collection
    oCollection.Add oArc

' Insert your code for output 8: Corner Curve 4
    oDirProj.Set 0, 0, 1
    Set ObjTurnSurface4 = PlaceProjection(m_OutputColl, oArc, oDirProj, parRoadDepth, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTurnSurface4
    Set ObjTurnSurface4 = Nothing

    Set oArc = Nothing

'  Construct Horizontal line at Point 1
   Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, Point1S(6), Point1S(7), Point1S(8), _
                                                             Point1S(3), Point1S(4), Point1S(5))
 
  ''Adding to Collection
    oCollection.Add oLine
    Set oLine = Nothing

' Insert your code for output 9: Cross Bottom Plane (Default Surface)
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    
    oStPoint.Set Point1S(3), Point1S(4), Point1S(5)
    Set oComplexStr = PlaceTrCString(oStPoint, oCollection)

    oDirProj.Set 0, 0, -1
    Set objRoadCrossBottom = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                    Point1S(3), Point1S(4), Point1S(5), _
                                                    oDirProj.x, oDirProj.y, oDirProj.z)
    
    Call objRoadCrossBottom.AddBoundary(oComplexStr)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRoadCrossBottom
    Set objRoadCrossBottom = Nothing
    
' Insert your code for output 10: Cross Top Plane
    Dim oTransVec As New AutoMath.DVector
    Dim oTransMatrix As New AutoMath.DT4x4
    
''Appy Translation to the Bottom Corss Complex string to get Top Surface Complex string by parRoadDepth.
    oTransMatrix.LoadIdentity
    oTransVec.Set 0, 0, parRoadDepth
    oTransMatrix.Translate oTransVec

    oComplexStr.Transform oTransMatrix
    
    oDirProj.Set 0, 0, 1
    Set objRoadCrossTop = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                    Point1S(3), Point1S(4), Point1S(5) + parRoadDepth, _
                                                    oDirProj.x, oDirProj.y, oDirProj.z)

    Call objRoadCrossTop.AddBoundary(oComplexStr)

 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRoadCrossTop
    Set objRoadCrossTop = Nothing
    
    oComplexStr.RemoveCurve True
    
    Dim iCount As Integer
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    
    Set oPoint = Nothing
    Set oComplexStr = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oDirProj = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    Set oTransMatrix = Nothing
    Set oTransVec = Nothing
    
    
' Insert your code for output 11(Point at center)
' Place Cross center(Origin)

    oCenterPos.Set 0, 0, 0
    
    Set objCrossCenter = oGeomFactory.Points3d.CreateByPoint(Nothing, oCenterPos.x, _
                                                                            oCenterPos.y, oCenterPos.z)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCrossCenter
    Set objCrossCenter = Nothing
    
' Insert your code for output 12
    Set objPoint1 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                            oCenterPos.x - (parRoadWidth / 2 + parRoadRadius), oCenterPos.y, oCenterPos.z)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint1
    Set objPoint1 = Nothing
    
' Insert your code for output 13
    Set objPoint2 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                            oCenterPos.x, oCenterPos.y - (parRoadWidth / 2 + parRoadRadius), oCenterPos.z)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint2
    Set objPoint2 = Nothing
    
' Insert your code for output 14
    Set objPoint3 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                            oCenterPos.x + (parRoadWidth / 2 + parRoadRadius), oCenterPos.y, oCenterPos.z)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint3
    Set objPoint3 = Nothing
    
' Insert your code for output 15
    Set objPoint4 = oGeomFactory.Points3d.CreateByPoint(Nothing, _
                            oCenterPos.x, oCenterPos.y + (parRoadWidth / 2 + parRoadRadius), oCenterPos.z)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint4
    Set objPoint4 = Nothing
    
    Set oCenterPos = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


