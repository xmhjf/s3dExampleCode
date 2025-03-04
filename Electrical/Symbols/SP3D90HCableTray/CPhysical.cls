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
'   Copyright (c) 2002-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Friday, Feb 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------    -----                   ------------------
'  09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'  08.SEP.2006     KKC                DI-95670  Replace names with initials in all revision history sheets and symbols
'  25.Sep.2006     KKC                TR-100618 Modified symbol code to honour Tangent Length
'  11.June.2007    KKC                DM-121367  Implemented part data basis for the Square Outside Corners
'  21.JUN.2007     dkl                CR-114880 Implemented part data basis to enable placement of segmented bend.
'  3.JUL.2007      dkl                CR-122171 Implemented part data bases 119 and 121 for Schneider Electric Catalog.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjHoriTangent  As Object
    Dim ObjInclTangent  As Object
    
    Dim parActualWidth As Double
    Dim parActualWidth1 As Double
    Dim parActualDepth As Double
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    Dim ProjVector As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
    Dim LineStrCP As AutoMath.DPosition
    Set LineStrCP = New AutoMath.DPosition
       
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parCTDimension = arrayOfInputs(2)
'    parTangentLength1 = arrayOfInputs(3)
'    parTangentLength2 = arrayOfInputs(4)
'    parEnd1toEnd2OuterEdge = arrayOfInputs(5)
'    parEnd2toEnd1OuterEdge = arrayOfInputs(6)

    iOutput = 0
    
    Dim parAngle As Double
    parAngle = PI / 2
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Call RetrieveCableTrayPortProperties(2, oPartFclt, parActualWidth1, parActualDepth)
    Dim HD As Double
    Dim HW As Double
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    
    Dim dFacetoCenter1 As Double   'Face to Center along Port 1.
    Dim dFacetoCenter2 As Double   'Face to Center along Port 2.
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    
'   The following Part data Bases are implemented,
'   1) 90 degree Horizontal Curved Bend, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) Bend Radius Dimension Exclusion basis.
'   3) 90 degree Horizontal Curved Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
'   4) 90 degree Horizontal Segmented Bend, Symmetrical, specified by Bend Radius and Tangent Length.
'   5) 90 degree Horizontal Segmented Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
'   6) 90 degree Horizontal Curved Bend, Symmetrical, specified by End-1 to End-2 outer edge and Bend Radius.
'   7) 90 degree Horizontal Curved Bend, Asymmetrical, specified by End-1 to End-2 outer edge, End-2 to End-1 outer edge and Bend Radius.
    
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 103    ' 103 - 90 degree Horizontal Curved Bend, Symmetrical, specified by
                            ' Bend Radius and Tangent Length (Default implementation).
            ' Retrieve Part properties: TangentLength
            Dim parTangentLength As Double
            parTangentLength = oTrayPart.TangentLength
            Dim parBendRadius As Double
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFacetoCenter1 = (parBendRadius + parActualWidth1 / 2) * Tan(parAngle / 2) + parTangentLength
            dFacetoCenter2 = dFacetoCenter1
            
        Case 5      'Bend Radius Dimension Exclusion basis.
            Dim parCTDimension As Double
            parCTDimension = arrayOfInputs(2)
            parTangentLength = oTrayPart.TangentLength
            dFacetoCenter1 = parCTDimension + parActualWidth1 / 2
            dFacetoCenter2 = dFacetoCenter1
            
        Case 105  ' 105 - 90 degree Horizontal Curved Bend, Asymmetrical, specified by
                 'Bend Radius, Tangent Length 1 and Tangent Length 2.
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = arrayOfInputs(3)
            dTangentLength2 = arrayOfInputs(4)
            dFacetoCenter1 = (parBendRadius + parActualWidth1 / 2) * Tan(parAngle / 2) + dTangentLength1
            dFacetoCenter2 = (parBendRadius + HW) * Tan(parAngle / 2) + dTangentLength2
            
        Case 107     ' 107 - 90 degree Horizontal Segmented Bend, Symmetrical, specified by
                    ' Bend Radius and Tangent Length.
            ' Retrieve Part properties: TangentLength
            parTangentLength = oTrayPart.TangentLength
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFacetoCenter1 = (parBendRadius + parActualWidth1 / 2) * Tan(parAngle / 2) + parTangentLength
            dFacetoCenter2 = dFacetoCenter1
            
        Case 109    ' 109 - 90 degree Horizontal Segmented Bend, Asymmetrical, specified by
                   'Bend Radius, Tangent Length 1 and Tangent Length 2.
            dTangentLength1 = arrayOfInputs(3)
            dTangentLength2 = arrayOfInputs(4)
            parBendRadius = oTrayPart.BendRadius
            dFacetoCenter1 = (parBendRadius + parActualWidth1 / 2) * Tan(parAngle / 2) + dTangentLength1
            dFacetoCenter2 = (parBendRadius + HW) * Tan(parAngle / 2) + dTangentLength2
        
        Case 119    '90 degree Horizontal Curved Bend, Symmetrical, specified by
                    'End-1 to End-2 outer edge and Bend Radius.
            Dim parEnd1toEnd2OuterEdge As Double
            parEnd1toEnd2OuterEdge = arrayOfInputs(5)
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parEnd1toEnd2OuterEdge - parBendRadius - parActualWidth1
            dTangentLength2 = dTangentLength1
            dFacetoCenter1 = parEnd1toEnd2OuterEdge - parActualWidth1 / 2
            dFacetoCenter2 = dFacetoCenter1
            
        Case 121    '90 degree Horizontal Curved Bend, Asymmetrical, specified by
                    'End-1 to End-2 outer edge, End-2 to End-1 outer edge and Bend Radius.
            Dim parEnd2toEnd1OuterEdge As Double
            parEnd1toEnd2OuterEdge = arrayOfInputs(5)
            parEnd2toEnd1OuterEdge = arrayOfInputs(6)
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = parEnd1toEnd2OuterEdge - parBendRadius - parActualWidth1
            dTangentLength2 = parEnd2toEnd1OuterEdge - parBendRadius - parActualWidth
            dFacetoCenter1 = parEnd1toEnd2OuterEdge - parActualWidth1 / 2
            dFacetoCenter2 = parEnd2toEnd1OuterEdge - parActualWidth / 2
        
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
      
    oPort1.Set -dFacetoCenter1, 0, 0
    oPort2.Set dFacetoCenter2 * Cos(parAngle), dFacetoCenter2 * Sin(parAngle), 0
      
' Insert your code for output 1(U Shape Horizontal Tangent)
    Dim CP As New AutoMath.DPosition 'arc center point
    
    Dim LineStrPoints(0 To 11)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
    LineStrPoints(0) = oPort1.x
    LineStrPoints(1) = oPort1.y + HW
    LineStrPoints(2) = oPort1.z + HD
    
    LineStrPoints(3) = oPort1.x
    LineStrPoints(4) = oPort1.y + HW
    LineStrPoints(5) = oPort1.z - HD
    
    LineStrPoints(6) = oPort1.x
    LineStrPoints(7) = oPort1.y - HW
    LineStrPoints(8) = oPort1.z - HD
    
    LineStrPoints(9) = oPort1.x
    LineStrPoints(10) = oPort1.y - HW
    LineStrPoints(11) = oPort1.z + HD
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set 1, 0, 0
    
    Set ObjHoriTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength1, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTangent
    Set ObjHoriTangent = Nothing
    Set oLineString = Nothing
    
    LineStrCP.Set -dFacetoCenter1 + dTangentLength1, 0, 0
    
' Insert your code for output 2(Elbow)
    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 103, 105, 119, 121
        
            Dim oElbowLineStr As IngrGeom3D.LineString3d
            LineStrPoints(0) = LineStrCP.x
            LineStrPoints(1) = LineStrCP.y + HW
            LineStrPoints(2) = LineStrCP.z + HD
            
            LineStrPoints(3) = LineStrCP.x
            LineStrPoints(4) = LineStrCP.y + HW
            LineStrPoints(5) = LineStrCP.z - HD
            
            LineStrPoints(6) = LineStrCP.x
            LineStrPoints(7) = LineStrCP.y - HW
            LineStrPoints(8) = LineStrCP.z - HD
            
            LineStrPoints(9) = LineStrCP.x
            LineStrPoints(10) = LineStrCP.y - HW
            LineStrPoints(11) = LineStrCP.z + HD
            Set oElbowLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
            ProjVector.Set 0, 0, 1
            CP.Set -dFacetoCenter1 + dTangentLength1, parBendRadius + HW, 0
            
            Dim ObjElbow As Object
            Set ObjElbow = PlaceRevolution(m_OutputColl, oElbowLineStr, ProjVector, CP, parAngle, False)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
            Set ObjElbow = Nothing
            Set oElbowLineStr = Nothing
    
        Case 5
        
            Dim oProjLineStr As IngrGeom3D.LineString3d
            Dim LineStrPoints1(0 To 14)  As Double
            LineStrPoints1(0) = LineStrCP.x
            LineStrPoints1(1) = LineStrCP.y + HW
            LineStrPoints1(2) = LineStrCP.z - HD
            
            LineStrPoints1(3) = LineStrCP.x
            LineStrPoints1(4) = LineStrCP.y - HW
            LineStrPoints1(5) = LineStrPoints1(2)
            
            LineStrPoints1(6) = parActualWidth1 / 2
            LineStrPoints1(7) = LineStrCP.y - HW
            LineStrPoints1(8) = LineStrPoints1(2)
            
            LineStrPoints1(9) = parActualWidth1 / 2
            LineStrPoints1(10) = parCTDimension + HW - dTangentLength2
            LineStrPoints1(11) = LineStrPoints1(2)
            
            LineStrPoints1(12) = -parActualWidth1 / 2
            LineStrPoints1(13) = parCTDimension + HW - dTangentLength2
            LineStrPoints1(14) = LineStrPoints1(2)
            
            
            Dim oHLine  As IngrGeom3D.Line3d
            Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5), LineStrPoints1(6), LineStrPoints1(7), LineStrPoints1(8))
            
            Dim oILine1 As IngrGeom3D.Line3d
            Set oILine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(6), LineStrPoints1(7), LineStrPoints1(8), LineStrPoints1(9), LineStrPoints1(10), LineStrPoints1(11))
            
            Dim oILine2  As IngrGeom3D.Line3d
            Set oILine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(9), LineStrPoints1(10), LineStrPoints1(11), LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14))
            
            Dim oILine3  As IngrGeom3D.Line3d
            Set oILine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14), LineStrPoints1(0), LineStrPoints1(1), LineStrPoints1(2))
                        
            Dim oVLine  As IngrGeom3D.Line3d
            Set oVLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(0), LineStrPoints1(1), LineStrPoints1(2), LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5))
            
            Dim oLiCollection  As Collection
            Set oLiCollection = New Collection
            oLiCollection.Add oHLine
            oLiCollection.Add oILine1
            Set ProjVector = New AutoMath.DVector
            Dim objBp1scurve  As IngrGeom3D.ComplexString3d
            Dim oStPoint   As AutoMath.DPosition
            Set oStPoint = New AutoMath.DPosition
            
            oStPoint.Set LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5)
            Set objBp1scurve = PlaceTrCString(oStPoint, oLiCollection)
            
            ProjVector.Set 0, 0, 1
            Dim ObjProj As IJDObject
            Set ObjProj = PlaceProjection(m_OutputColl, objBp1scurve, ProjVector, parActualDepth, True)
            
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProj
            Set ObjProj = Nothing
            Set oLiCollection = Nothing
            
            Dim oLiCollection1 As Collection
            Set oLiCollection1 = New Collection
            oLiCollection1.Add oILine3
            Set oStPoint = New AutoMath.DPosition
            
            oStPoint.Set LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14)
            Set objBp1scurve = New IngrGeom3D.ComplexString3d
            Set objBp1scurve = PlaceTrCString(oStPoint, oLiCollection1)
            Dim ObjProj1 As IJDObject
            Set ObjProj1 = PlaceProjection(m_OutputColl, objBp1scurve, ProjVector, parActualDepth, False)
            
        ' Set the output
            m_OutputColl.AddOutput "Proj_", ObjProj1
            Set ObjProj1 = Nothing
            Set objBp1scurve = Nothing
            Set oLiCollection1 = Nothing
            
            Dim obranchcurves As Collection
            Set obranchcurves = New Collection
            obranchcurves.Add oHLine
            obranchcurves.Add oILine1
            obranchcurves.Add oILine2
            obranchcurves.Add oILine3
            obranchcurves.Add oVLine
            
            Dim objHBcurves         As IngrGeom3D.ComplexString3d
            oStPoint.Set LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5)
            Set objHBcurves = PlaceTrCString(oStPoint, obranchcurves)
            
            Dim TrayBottom As IngrGeom3D.Plane3d
            Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                   -0, 0, -HD, ProjVector.x, ProjVector.y, ProjVector.z)
            
            Call TrayBottom.AddBoundary(objHBcurves)
                        
        ' Set the output
            m_OutputColl.AddOutput "Proj_", TrayBottom
            Set TrayBottom = Nothing
            Set objHBcurves = Nothing
            Set obranchcurves = Nothing
            Set oStPoint = Nothing
        
        Case 107, 109
            
            'Create the bottom surface adjoining Tangent 1 of the segment.
            Dim dPointOuterBendTan1 As AutoMath.DPosition 'Point at outer bend along Tangent 1.
            Set dPointOuterBendTan1 = New AutoMath.DPosition
            dPointOuterBendTan1.Set LineStrCP.x + 2 * HW / Tan(3 * PI / 8), _
                        LineStrCP.y - HW, LineStrCP.z - HD
                        
            Dim dPoints() As Double
            ReDim dPoints(0 To 8)
            
            dPoints(0) = LineStrCP.x
            dPoints(1) = LineStrCP.y + HW
            dPoints(2) = LineStrCP.z - HD
            
            dPoints(3) = dPointOuterBendTan1.x
            dPoints(4) = dPointOuterBendTan1.y
            dPoints(5) = dPoints(2)

            dPoints(6) = dPoints(0)
            dPoints(7) = LineStrCP.y - HW
            dPoints(8) = dPoints(2)
            
            Dim objBottomSegmentSurface1 As IngrGeom3D.Plane3d
            Set objBottomSegmentSurface1 = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            3, dPoints)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomSegmentSurface1
            Set objBottomSegmentSurface1 = Nothing
            
            'Create the bottom surface in-between the surfaces adjoining Tangent 1 and tangent 2.
            Dim oLineStrCP2 As AutoMath.DPosition   'Point where Tangent 2 begins.
            Set oLineStrCP2 = New AutoMath.DPosition
            oLineStrCP2.Set oPort2.x, oPort2.y - dTangentLength2, oPort2.z
            
            Dim dPointOuterBendTan2 As AutoMath.DPosition 'Point at outer bend along Tangent 2.
            Set dPointOuterBendTan2 = New AutoMath.DPosition
            dPointOuterBendTan2.Set oLineStrCP2.x + parActualWidth1 / 2, _
                oLineStrCP2.y - parActualWidth1 / Tan(3 * PI / 8), oLineStrCP2.z - HD
            
            ReDim dPoints(0 To 11)
            dPoints(0) = LineStrCP.x
            dPoints(1) = LineStrCP.y + HW
            dPoints(2) = LineStrCP.z - HD

            dPoints(3) = dPointOuterBendTan1.x
            dPoints(4) = dPointOuterBendTan1.y
            dPoints(5) = dPoints(2)

            dPoints(6) = dPointOuterBendTan2.x
            dPoints(7) = dPointOuterBendTan2.y
            dPoints(8) = dPoints(2)

            dPoints(9) = oLineStrCP2.x - parActualWidth1 / 2
            dPoints(10) = oLineStrCP2.y
            dPoints(11) = dPoints(2)
            
            Dim objBottomSegmentSurfaceMiddle As IngrGeom3D.Plane3d
            Set objBottomSegmentSurfaceMiddle = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            4, dPoints)
                      
            ' Set the output
            m_OutputColl.AddOutput "BottomSegmentSurfaceMiddle", objBottomSegmentSurfaceMiddle
            Set objBottomSegmentSurfaceMiddle = Nothing

            'Create the bottom surface adjoining Tangent 2 of the segment.
            
            ReDim dPoints(0 To 8)
            dPoints(0) = oLineStrCP2.x - parActualWidth1 / 2
            dPoints(1) = oLineStrCP2.y
            dPoints(2) = oLineStrCP2.z - HD
            
            dPoints(3) = dPointOuterBendTan2.x
            dPoints(4) = dPointOuterBendTan2.y
            dPoints(5) = dPoints(2)

            dPoints(6) = oLineStrCP2.x + parActualWidth1 / 2
            dPoints(7) = dPoints(1)
            dPoints(8) = dPoints(2)
            
            Dim objBottomSegmentSurface2 As IngrGeom3D.Plane3d
            Set objBottomSegmentSurface2 = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            3, dPoints)
                      
            ' Set the output
            m_OutputColl.AddOutput "BottomSegmentSurface2", objBottomSegmentSurface2
            Set objBottomSegmentSurface2 = Nothing
            
            'Create the inner inclined surface of the segment.
            ReDim dPoints(0 To 11)
            dPoints(0) = LineStrCP.x
            dPoints(1) = LineStrCP.y + HW
            dPoints(2) = LineStrCP.z - HD
            
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = LineStrCP.z + HD
            
            dPoints(6) = oLineStrCP2.x - parActualWidth1 / 2
            dPoints(7) = oLineStrCP2.y
            dPoints(8) = oLineStrCP2.z + HD
            
            dPoints(9) = dPoints(6)
            dPoints(10) = dPoints(7)
            dPoints(11) = oLineStrCP2.z - HD
            
            Dim objInnerSegmentSurfaceInclined As IngrGeom3D.Plane3d
            Set objInnerSegmentSurfaceInclined = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            4, dPoints)
            
            ' Set the output
            m_OutputColl.AddOutput "InnerSegmentSurfaceInclined", objInnerSegmentSurfaceInclined
            Set objInnerSegmentSurfaceInclined = Nothing
            
            'Create the Outer inclined surface of the segment.
            
            Dim oTraceLineSegment As IngrGeom3D.Line3d
            Set oTraceLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                LineStrCP.x, LineStrCP.y - HW, LineStrCP.z - HD, _
                dPointOuterBendTan1.x, dPointOuterBendTan1.y, dPointOuterBendTan1.z)
            
            Dim oTraceCurveColl  As Collection
            Set oTraceCurveColl = New Collection
            oTraceCurveColl.Add oTraceLineSegment
            
            Set oTraceLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPointOuterBendTan1.x, dPointOuterBendTan1.y, dPointOuterBendTan1.z, _
                dPointOuterBendTan2.x, dPointOuterBendTan2.y, dPointOuterBendTan2.z)
            oTraceCurveColl.Add oTraceLineSegment
            
            Set oTraceLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPointOuterBendTan2.x, dPointOuterBendTan2.y, dPointOuterBendTan2.z, _
                oLineStrCP2.x + parActualWidth1 / 2, oLineStrCP2.y, oLineStrCP2.z - HD)
            oTraceCurveColl.Add oTraceLineSegment
            
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set LineStrCP.x, LineStrCP.y - HW, LineStrCP.z - HD
            
            Dim objTraceCurve  As IngrGeom3D.ComplexString3d
            Set objTraceCurve = PlaceTrCString(oStPoint, oTraceCurveColl)
                
            Dim objOuterSegmentSurfaceInclined As IJDObject
            ProjVector.Set 0, 0, 1
            Set objOuterSegmentSurfaceInclined = PlaceProjection(m_OutputColl, objTraceCurve, ProjVector, parActualDepth, True)
                      
            ' Set the output
            Dim ObjSurface As Object
            m_OutputColl.AddOutput "OuterSegmentSurfaceInclined", objOuterSegmentSurfaceInclined
            Set objOuterSegmentSurfaceInclined = Nothing
            Set objTraceCurve = Nothing
            Set oTraceCurveColl = Nothing
            Set oTraceLineSegment = Nothing
            Set oStPoint = Nothing
            Set oLineStrCP2 = Nothing
            Set dPointOuterBendTan1 = Nothing
            Set dPointOuterBendTan2 = Nothing
                       
        Case Else
            GoTo ErrorLabel:
            
    End Select
     
     ' Insert your code for output 3(U Shape Inclined Tangent)
    Dim oInclinedLineStr As IngrGeom3D.LineString3d
        
    LineStrPoints(0) = oPort2.x - HW * Sin(parAngle)
    LineStrPoints(1) = oPort2.y + HW * Cos(parAngle)
    LineStrPoints(2) = oPort2.z + HD
    
    LineStrPoints(3) = oPort2.x - HW * Sin(parAngle)
    LineStrPoints(4) = oPort2.y + HW * Cos(parAngle)
    LineStrPoints(5) = oPort2.z - HD
    
    LineStrPoints(6) = oPort2.x + HW * Sin(parAngle)
    LineStrPoints(7) = oPort2.y - HW * Cos(parAngle)
    LineStrPoints(8) = oPort2.z - HD
    
    LineStrPoints(9) = oPort2.x + HW * Sin(parAngle)
    LineStrPoints(10) = oPort2.y - HW * Cos(parAngle)
    LineStrPoints(11) = oPort2.z + HD
    Set oInclinedLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set -Cos(parAngle), -Sin(parAngle), 0
    
    Set ObjInclTangent = PlaceProjection(m_OutputColl, oInclinedLineStr, ProjVector, dTangentLength2, False)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInclTangent
    Set ObjInclTangent = Nothing
    Set ProjVector = Nothing
    Set oInclinedLineStr = Nothing
    Set oGeomFactory = Nothing
        
' Place Nozzle 1
     
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

' Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition

    oDir.Set Cos(parAngle), Sin(parAngle), 0
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

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub

