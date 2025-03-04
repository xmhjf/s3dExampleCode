VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday, Aug 25 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Style HP-70 Victaulic Coupling. Symbol details are provided in 06.12.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI       As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
     Set m_oSymGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim ObjClampBodyFront As Object
    Dim ObjClampBodyRear As Object
    Dim ObjSupFrontTop As Object
    Dim ObjSupFrontBottom As Object
    Dim ObjSupRearTop As Object
    Dim ObjSupRearBottom As Object
    Dim ObjLeftFrontBottomNut  As Object
    Dim ObjLeftFrontTopNut  As Object
    Dim ObjLeftRearBottomNut  As Object
    Dim ObjLeftRearTopNut  As Object
    Dim ObjRightFrontBottomNut  As Object
    Dim ObjRightFrontTopNut  As Object
    Dim ObjRightRearBottomNut  As Object
    Dim ObjRightRearTopNut  As Object
    
    Dim ClampTk As Double
    Dim ClampSideWallTk As Double
    Dim ClampTotalTk As Double
    Dim BoltSupLen As Double
    Dim NutHeight   As Double
    Dim supArcAngle As Double
    Dim GapBetweenUCandLC   As Double
    Dim iCount As Integer

    Dim parClampOuterDiameter As Double
    Dim parClampLength As Double
    Dim parClampWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampOuterDiameter = arrayOfInputs(2)
    parClampLength = arrayOfInputs(3)
    parClampWidth = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    Dim ClampInnerDiameter As Double
    Dim BoltCenterDistance As Double
    Dim LugSeparation As Double
    Dim BoltDiameter As Double
    Dim BoltLength As Double
    Dim BoltLugWidth As Double
    
'   Assumptions: Following dimensions are taken to be proportions of the input parameters
    ClampInnerDiameter = parClampOuterDiameter / 1.2
    LugSeparation = 0.6 * parClampOuterDiameter
    BoltDiameter = 0.1 * parClampOuterDiameter
    BoltLength = LugSeparation + 4 * BoltDiameter
    BoltLugWidth = parClampWidth
    BoltCenterDistance = parClampLength - 3 * BoltDiameter

    m_oSymGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

'   Assumptions: Clamp thickness 0.1 times the Clamp Outer Diameter, Clamp side wall thickness
'   is 0.1 times the Clamp Width
    ClampTk = 0.1 * parClampOuterDiameter
    ClampSideWallTk = 0.1 * parClampWidth

'   Assumption: Gap between the clamps is taken as 4 mm
    GapBetweenUCandLC = 0.004
    ClampTotalTk = LugSeparation / 2 - GapBetweenUCandLC / 2

'   Assumption: Nut height is taken 30% of available length for fastening
    NutHeight = (BoltLength - LugSeparation) * 0.3
 
'   Intersection of support with the clamp body in side view is an arc. Compute its angle.
    supArcAngle = Atn(ClampTotalTk / Sqr(Abs((parClampOuterDiameter / 2) ^ 2 - ClampTotalTk ^ 2)))
    
    Dim IntersecDistance As Double
    IntersecDistance = 0.5 * parClampOuterDiameter * (1 - Cos(supArcAngle))

    BoltSupLen = (parClampLength - parClampOuterDiameter) / 2 + IntersecDistance

' Insert your code for output 2(Clamp Body Front-Half)
    Dim LineStrPoints()  As Double
    ReDim LineStrPoints(0 To 26) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim RevVector As AutoMath.DVector
    Dim CenPtDPosition As AutoMath.DPosition
    
    Set RevVector = New AutoMath.DVector
    Set CenPtDPosition = New AutoMath.DPosition
    
'   Point at Y-axis and Negative X-axis near Center(1point)
    LineStrPoints(0) = -parClampWidth / 2
    LineStrPoints(1) = ClampInnerDiameter / 2
    LineStrPoints(2) = GapBetweenUCandLC / 2
    
'   Point at Y-axis and  Negative X-axis away from Center (2point)
    LineStrPoints(3) = -parClampWidth / 2
    LineStrPoints(4) = parClampOuterDiameter / 2
    LineStrPoints(5) = LineStrPoints(2)
    
'   Point at Y-axis and Positive X-axis away from Center(3point)
    LineStrPoints(6) = parClampWidth / 2
    LineStrPoints(7) = parClampOuterDiameter / 2
    LineStrPoints(8) = LineStrPoints(2)
    
'   Point at Y-axis and Positive X-axis near Center (4point)
    LineStrPoints(9) = parClampWidth / 2
    LineStrPoints(10) = ClampInnerDiameter / 2
    LineStrPoints(11) = LineStrPoints(2)
    
'   Point at Y-axis and Positive X-axis near Center (5point)
    LineStrPoints(12) = parClampWidth / 2 - ClampSideWallTk
    LineStrPoints(13) = ClampInnerDiameter / 2
    LineStrPoints(14) = LineStrPoints(2)
    
'   Point at Y-axis and Positive X-axis inner wall of clamp (6point)
    Dim ClampDiameter As Double
    ClampDiameter = parClampOuterDiameter - 2 * ClampTk

    LineStrPoints(15) = parClampWidth / 2 - ClampSideWallTk
    LineStrPoints(16) = ClampDiameter / 2
    LineStrPoints(17) = LineStrPoints(2)
    
'   Point at Y-axis and Negative X-axis inner wall of clamp (7point)
    LineStrPoints(18) = -parClampWidth / 2 + ClampSideWallTk
    LineStrPoints(19) = ClampDiameter / 2
    LineStrPoints(20) = LineStrPoints(2)
    
'   Point at Y-axis and Negative X-axis (8point)
    LineStrPoints(21) = -parClampWidth / 2 + ClampSideWallTk
    LineStrPoints(22) = ClampInnerDiameter / 2
    LineStrPoints(23) = LineStrPoints(2)
    
'   Starting Point
    LineStrPoints(24) = LineStrPoints(0)
    LineStrPoints(25) = LineStrPoints(1)
    LineStrPoints(26) = LineStrPoints(2)

    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
   
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStrPoints)

    RevVector.Set 1, 0, 0
    CenPtDPosition.Set 0, 0, GapBetweenUCandLC / 2
    Set ObjClampBodyFront = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampBodyFront
    Set ObjClampBodyFront = Nothing
    Set oLineString = Nothing
    
' Insert your code for output 3(Clamp Body Rear-Half)
    For iCount = 1 To 9
        LineStrPoints(3 * iCount - 1) = -GapBetweenUCandLC / 2
    Next iCount
    RevVector.Set -1, 0, 0
    CenPtDPosition.Set 0, 0, -GapBetweenUCandLC / 2
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStrPoints)
    Set ObjClampBodyRear = PlaceRevolution(m_OutputColl, oLineString, RevVector, CenPtDPosition, PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampBodyRear
    Set ObjClampBodyRear = Nothing
    Set RevVector = Nothing
    Set CenPtDPosition = Nothing
    Set oLineString = Nothing

' Insert your code for output 4(Front Top Support)
'   Prepare section-curve
    ReDim LineStrPoints(0 To 20) As Double
    LineStrPoints(0) = -BoltLugWidth / 2
    LineStrPoints(1) = (parClampOuterDiameter / 2) * Cos(supArcAngle)
    LineStrPoints(2) = GapBetweenUCandLC / 2 + 0.5 * parClampOuterDiameter * Sin(supArcAngle)
    
    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = (parClampOuterDiameter / 2) * Cos(supArcAngle) + BoltSupLen
    LineStrPoints(5) = LineStrPoints(2)
    
    LineStrPoints(6) = LineStrPoints(0)
    LineStrPoints(7) = LineStrPoints(4)
    LineStrPoints(8) = GapBetweenUCandLC / 2 + ClampTotalTk / 2      '20% of clamp total thickness
    
    LineStrPoints(9) = LineStrPoints(0)
    LineStrPoints(10) = (parClampOuterDiameter / 2) * Cos(supArcAngle) + BoltSupLen * 0.75 '75% of the bolt support length
    LineStrPoints(11) = LineStrPoints(8)
    
    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = (parClampOuterDiameter / 2) * Cos(supArcAngle) + BoltSupLen * 0.5  '50% of the bolt support length
    LineStrPoints(14) = GapBetweenUCandLC / 2
    
    LineStrPoints(15) = LineStrPoints(0)
    LineStrPoints(16) = LineStrPoints(1)
    LineStrPoints(17) = LineStrPoints(14)
    
    LineStrPoints(18) = LineStrPoints(0)
    LineStrPoints(19) = LineStrPoints(1)
    LineStrPoints(20) = LineStrPoints(2)
   
'   Store the coordinates of Upper Support section
    Dim SupFrontTopPts(0 To 20) As Double
    For iCount = 0 To 20
        SupFrontTopPts(iCount) = LineStrPoints(iCount)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Dim projVector As New AutoMath.DVector
    projVector.Set 1, 0, 0
    Set ObjSupFrontTop = PlaceProjection(m_OutputColl, oLineString, projVector, BoltLugWidth, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupFrontTop
    Set ObjSupFrontTop = Nothing
     
' Insert your code for output 5(Front Bottom Support)
'   Prepare section-curve points
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = SupFrontTopPts(3 * iCount - 3)
        LineStrPoints(3 * iCount - 2) = -SupFrontTopPts(3 * iCount - 2)
        LineStrPoints(3 * iCount - 1) = SupFrontTopPts(3 * iCount - 1)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjSupFrontBottom = PlaceProjection(m_OutputColl, oLineString, projVector, BoltLugWidth, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupFrontBottom
    Set ObjSupFrontBottom = Nothing
    
' Insert your code for output 6(Rear Top Support)
'   Prepare section-curve points
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = SupFrontTopPts(3 * iCount - 3)
        LineStrPoints(3 * iCount - 2) = SupFrontTopPts(3 * iCount - 2)
        LineStrPoints(3 * iCount - 1) = -SupFrontTopPts(3 * iCount - 1)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjSupRearTop = PlaceProjection(m_OutputColl, oLineString, projVector, BoltLugWidth, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupRearTop
    Set ObjSupRearTop = Nothing
   
' Insert your code for output 7(Rear Bottom Support)
'   Prepare section-curve points
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = SupFrontTopPts(3 * iCount - 3)
        LineStrPoints(3 * iCount - 2) = -SupFrontTopPts(3 * iCount - 2)
        LineStrPoints(3 * iCount - 1) = -SupFrontTopPts(3 * iCount - 1)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjSupRearBottom = PlaceProjection(m_OutputColl, oLineString, projVector, BoltLugWidth, True)
    Set oLineString = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupRearBottom
    Set ObjSupRearBottom = Nothing
    
' Insert your code for output 8(Left Top Bolt)
    Dim BoltStPt As AutoMath.DPosition
    Dim BoltEndPt As AutoMath.DPosition
    Set BoltStPt = New AutoMath.DPosition
    Set BoltEndPt = New AutoMath.DPosition

    BoltStPt.Set 0, BoltCenterDistance / 2, BoltLength / 2
    BoltEndPt.Set BoltStPt.x, BoltStPt.y, BoltStPt.z - BoltLength

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), BoltStPt, BoltEndPt, BoltDiameter
    
' Insert your code for output 9(Left Bottom Bolt)
    BoltStPt.Set 0, -BoltCenterDistance / 2, BoltLength / 2
    BoltEndPt.Set BoltStPt.x, BoltStPt.y, BoltStPt.z - BoltLength

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), BoltStPt, BoltEndPt, BoltDiameter

' Insert your code for output 12(Left Front Top Nut)
    Dim HalfNutWidthAcrossCorners As Double

    HalfNutWidthAcrossCorners = BoltDiameter * Sqr(3) / 2

    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = HalfNutWidthAcrossCorners * Cos(iCount * PI / 3)
        LineStrPoints(3 * iCount - 2) = BoltCenterDistance / 2 + HalfNutWidthAcrossCorners * Sin(iCount * PI / 3)
        LineStrPoints(3 * iCount - 1) = LugSeparation / 2
    Next iCount
    
    projVector.Set 0, 0, 1
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjLeftFrontTopNut = PlaceProjection(m_OutputColl, oLineString, projVector, NutHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftFrontTopNut
    Set ObjLeftFrontTopNut = Nothing

' Insert your code for output 13(Left Rear Top Nut)
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = HalfNutWidthAcrossCorners * Cos(iCount * PI / 3)
        LineStrPoints(3 * iCount - 2) = BoltCenterDistance / 2 + HalfNutWidthAcrossCorners * Sin(iCount * PI / 3)
        LineStrPoints(3 * iCount - 1) = -LugSeparation / 2
    Next iCount
    projVector.Set 0, 0, -1
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjLeftRearTopNut = PlaceProjection(m_OutputColl, oLineString, projVector, NutHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftRearTopNut
    Set ObjLeftRearTopNut = Nothing

' Insert your code for output 16(Left Front Bottom Nut)
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = HalfNutWidthAcrossCorners * Cos(iCount * PI / 3)
        LineStrPoints(3 * iCount - 2) = -BoltCenterDistance / 2 + HalfNutWidthAcrossCorners * Sin(iCount * PI / 3)
        LineStrPoints(3 * iCount - 1) = LugSeparation / 2
    Next iCount
    projVector.Set 0, 0, 1
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjLeftFrontBottomNut = PlaceProjection(m_OutputColl, oLineString, projVector, NutHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftFrontBottomNut
    Set ObjLeftFrontBottomNut = Nothing
    
' Insert your code for output 17(Left Rear Bottom Nut)
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = HalfNutWidthAcrossCorners * Cos(iCount * PI / 3)
        LineStrPoints(3 * iCount - 2) = -BoltCenterDistance / 2 + HalfNutWidthAcrossCorners * Sin(iCount * PI / 3)
        LineStrPoints(3 * iCount - 1) = -LugSeparation / 2
    Next iCount
    projVector.Set 0, 0, -1
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set ObjLeftRearBottomNut = PlaceProjection(m_OutputColl, oLineString, projVector, NutHeight, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftRearBottomNut
    Set ObjLeftRearBottomNut = Nothing
    Set projVector = Nothing
    Set BoltStPt = Nothing
    Set BoltEndPt = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
