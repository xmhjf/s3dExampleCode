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
'   Copyright (c) 2003-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Tuesday, Feb 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'Veritical CableTray Inside consist of a Horizontal Tangent and Veritical Curve and Vertical Tangent
'The shape is formed by the following Three steps:
'1.For Horizontal Tangent by Constructing a Line String and Projecting along X-Axis
'2.For Vertical Curve by Constructing a Line String and Revolving about Y-Axis
'3.For Vertical Tangent by Constructing a Line String and Projecting along Z-Axis
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.May.2004     MS     TR 58960 Removed hardcoded TangentLength value.
'                                           Now the TangentLength is retrieved from oTrayPart
'   13.May.2004     MS     CR 59321 Provided a check for TangentLength, if the
'                                           TangentLength value retrieved from the part is zero then put it to a very small value 0.000001
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  30.MAY.2007     dkl  CR-114880 used approprite CmpdblXXX functions for comparing double variables.
'                                 Implemented part data basis to enable placement of segmented bend.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim objVertiTangent As Object
    Dim objHoriTangent As Object
    Dim objVertiCurve As Object
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parTangentLength1 = arrayOfInputs(2)
'    parTangentLength2 = arrayOfInputs(3)
              
    iOutput = 0
    
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    
'   The following Part data Bases are implemented,
'   1) 90 degree Vertical Curved Upward Bend, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) 90 degree Vertical Curved Upward Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
'   3) 90 degree Vertical Segmented Upward Bend, Symmetrical, specified by Bend Radius and Tangent Length.
'   4) 90 degree Vertical Segmented Upward Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
    
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
    
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 31, 35 ' 31 - 90 degree Vertical Curved Upward Bend, Symmetrical, specified by
                             'Bend Radius and Tangent Length (Default implementation).
                             '35 - 90 degree Vertical Segmented Upward Bend, Symmetrical, specified by
                             'Bend Radius and Tangent Length.
                             
            ' Retrieve Part properties: TangentLength
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            
        Case 33, 37  ' 33 - 90 degree Vertical Curved Upward Bend, Asymmetrical, specified by
                     'Bend Radius, Tangent Length 1 and Tangent Length 2.
                     '37 - 90 degree Vertical Segmented Upward Bend, Asymmetrical, specified by
                     'Bend Radius, Tangent Length 1 and Tangent Length 2.
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
                      
        Case Else
            GoTo ErrorLabel:
    End Select
    
'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
    
' Insert your code for output 1(Horizontal Tangent)
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth              As Double
    Dim HalfWidth              As Double
    Dim ProjVector        As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
    oPort1.Set -(parBendRadius + HalfDepth + dTangentLength1), 0, 0

    Dim LineStrPoints()  As Double
    ReDim LineStrPoints(0 To 11)
    
    LineStrPoints(0) = oPort1.x
    LineStrPoints(1) = oPort1.y - HalfWidth
    LineStrPoints(2) = oPort1.z + HalfDepth
    
    LineStrPoints(3) = oPort1.x
    LineStrPoints(4) = oPort1.y - HalfWidth
    LineStrPoints(5) = oPort1.z - HalfDepth
    
    LineStrPoints(6) = oPort1.x
    LineStrPoints(7) = oPort1.y + HalfWidth
    LineStrPoints(8) = oPort1.z - HalfDepth
    
    LineStrPoints(9) = oPort1.x
    LineStrPoints(10) = oPort1.y + HalfWidth
    LineStrPoints(11) = oPort1.z + HalfDepth
    ProjVector.Set 1, 0, 0
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength1, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing

' Insert your code for output 2(Vertical Bend)
    Call RetrieveCableTrayPortProperties(2, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth2 As Double
    Dim HalfWidth2 As Double
    HalfDepth2 = parActualDepth / 2
    HalfWidth2 = parActualWidth / 2
    
    Dim LineStrCP As New AutoMath.DPosition
    Dim CP As New AutoMath.DPosition
    LineStrCP.Set -(parBendRadius + HalfDepth2), 0, 0
    
    Select Case lCableTrayPartDataBasis
        Case Is <= 1, 31, 33
        
            LineStrPoints(0) = LineStrCP.x
            LineStrPoints(1) = LineStrCP.y - HalfWidth
            LineStrPoints(2) = LineStrCP.z + HalfDepth
            
            LineStrPoints(3) = LineStrCP.x
            LineStrPoints(4) = LineStrCP.y - HalfWidth
            LineStrPoints(5) = LineStrCP.z - HalfDepth
            
            LineStrPoints(6) = LineStrCP.x
            LineStrPoints(7) = LineStrCP.y + HalfWidth
            LineStrPoints(8) = LineStrCP.z - HalfDepth
            
            LineStrPoints(9) = LineStrCP.x
            LineStrPoints(10) = LineStrCP.y + HalfWidth
            LineStrPoints(11) = LineStrCP.z + HalfDepth
            
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, _
                                                                            LineStrPoints)
            
            ProjVector.Set 0, -1, 0
            CP.Set -(parBendRadius + HalfDepth2), 0, parBendRadius + HalfDepth
            
            Set objVertiCurve = PlaceRevolution(m_OutputColl, oLineString, ProjVector, CP, _
                                                                            PI / 2, False)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiCurve
            Set objVertiCurve = Nothing
            
        Case 35, 37
            Dim oPointEndSegment2 As AutoMath.DPosition    'Point at a End of the segment (before
            Set oPointEndSegment2 = New AutoMath.DPosition 'the tangent at port 2).
            oPointEndSegment2.Set 0, 0, parBendRadius + HalfDepth
              
            'Create the bottom surface of the base of the segment.
            LineStrPoints(0) = LineStrCP.x
            LineStrPoints(1) = LineStrCP.y + HalfWidth
            LineStrPoints(2) = LineStrCP.z - HalfDepth
            
            LineStrPoints(3) = LineStrPoints(0) + HalfDepth / Tan(3 * PI / 8)
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)

            LineStrPoints(6) = LineStrCP.x + HalfDepth / Tan(3 * PI / 8)
            LineStrPoints(7) = LineStrCP.y - HalfWidth
            LineStrPoints(8) = LineStrCP.z - HalfDepth
            
            LineStrPoints(9) = LineStrCP.x
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)

            Dim objBottomSegmentSurface1 As IngrGeom3D.Plane3d
            Set objBottomSegmentSurface1 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            4, LineStrPoints)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomSegmentSurface1
            Set objBottomSegmentSurface1 = Nothing
            
            'Create the inclined surface of the base of the segment.
            LineStrPoints(0) = LineStrCP.x + HalfDepth / Tan(3 * PI / 8)
            LineStrPoints(1) = LineStrCP.y + HalfWidth
            LineStrPoints(2) = LineStrCP.z - HalfDepth
            
            LineStrPoints(3) = HalfDepth2
            LineStrPoints(4) = LineStrCP.y + HalfWidth2
            LineStrPoints(5) = oPointEndSegment2.z - HalfDepth2 / Tan(3 * PI / 8)
            
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = LineStrCP.y - HalfWidth2
            LineStrPoints(8) = LineStrPoints(5)
            
            LineStrPoints(9) = LineStrPoints(0)
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            Dim objBottomSegmentSurfaceInclined As IngrGeom3D.Plane3d
            Set objBottomSegmentSurfaceInclined = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            4, LineStrPoints)
            
            ' Set the output
            m_OutputColl.AddOutput "BottomSegmentSurfaceInclined", objBottomSegmentSurfaceInclined
            Set objBottomSegmentSurfaceInclined = Nothing
            
            'Create the top surface of the base of the segment.
            LineStrPoints(0) = HalfDepth2
            LineStrPoints(1) = LineStrCP.y + HalfWidth2
            LineStrPoints(2) = oPointEndSegment2.z - HalfDepth2 / Tan(3 * PI / 8)
         
            LineStrPoints(3) = LineStrPoints(0)
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = oPointEndSegment2.z
           
            LineStrPoints(6) = LineStrPoints(0)
            LineStrPoints(7) = LineStrCP.y - HalfWidth2
            LineStrPoints(8) = oPointEndSegment2.z
              
            LineStrPoints(9) = LineStrPoints(0)
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            Dim objBottomSegmentSurface2 As IngrGeom3D.Plane3d
            Set objBottomSegmentSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            4, LineStrPoints)
                      
            ' Set the output
            m_OutputColl.AddOutput "BottomSegmentSurface2", objBottomSegmentSurface2
            Set objBottomSegmentSurface2 = Nothing
            
            'Create the Front Segment Surface.
            ReDim LineStrPoints(0 To 17)
            
            LineStrPoints(0) = LineStrCP.x
            LineStrPoints(1) = LineStrCP.y - HalfWidth
            LineStrPoints(2) = LineStrCP.z - HalfDepth

            LineStrPoints(3) = LineStrPoints(0) + HalfDepth / Tan(3 * PI / 8)
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)

            LineStrPoints(6) = HalfDepth
            LineStrPoints(7) = LineStrPoints(1)
            LineStrPoints(8) = oPointEndSegment2.z - HalfDepth2 / Tan(3 * PI / 8)

            LineStrPoints(9) = LineStrPoints(6)
            LineStrPoints(10) = LineStrPoints(1)
            LineStrPoints(11) = oPointEndSegment2.z
            
            LineStrPoints(12) = -HalfDepth2
            LineStrPoints(13) = LineStrPoints(1)
            LineStrPoints(14) = LineStrPoints(11)
            
            LineStrPoints(15) = LineStrCP.x
            LineStrPoints(16) = LineStrPoints(1)
            LineStrPoints(17) = LineStrCP.z + HalfDepth
            
            Dim objFrontSegmentSurface As IngrGeom3D.Plane3d
            Set objFrontSegmentSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            6, LineStrPoints)
                      
            ' Set the output
            m_OutputColl.AddOutput "SegmentSurfaceOnTheFrontSide", objFrontSegmentSurface
            Set objFrontSegmentSurface = Nothing

            'Create the Rear Segment Surface.
            
            LineStrPoints(0) = LineStrCP.x
            LineStrPoints(1) = LineStrCP.y + HalfWidth
            LineStrPoints(2) = LineStrCP.z - HalfDepth
            
            LineStrPoints(3) = LineStrPoints(0) + HalfDepth / Tan(3 * PI / 8)
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)

            LineStrPoints(6) = HalfDepth
            LineStrPoints(7) = LineStrPoints(1)
            LineStrPoints(8) = oPointEndSegment2.z - HalfDepth2 / Tan(3 * PI / 8)
         
            LineStrPoints(9) = LineStrPoints(6)
            LineStrPoints(10) = LineStrPoints(1)
            LineStrPoints(11) = oPointEndSegment2.z
            
            LineStrPoints(12) = -HalfDepth2
            LineStrPoints(13) = LineStrPoints(1)
            LineStrPoints(14) = LineStrPoints(11)
            
            LineStrPoints(15) = LineStrCP.x
            LineStrPoints(16) = LineStrPoints(1)
            LineStrPoints(17) = LineStrCP.z + HalfDepth
            
            Dim objRearSegmentSurface As IngrGeom3D.Plane3d
            Set objRearSegmentSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                                            6, LineStrPoints)
                      
            ' Set the output
            m_OutputColl.AddOutput "SegmentSurfaceOnTheRearSide", objRearSegmentSurface
            Set objRearSegmentSurface = Nothing
            
        Case Else
            GoTo ErrorLabel:
            
    End Select

' Insert your code for output 3(Vertical Tangent)
    oPort2.Set 0, 0, (parBendRadius + HalfDepth + dTangentLength2)
    LineStrPoints(0) = oPort2.x - HalfDepth2
    LineStrPoints(1) = oPort2.y - HalfWidth2
    LineStrPoints(2) = oPort2.z
    
    LineStrPoints(3) = oPort2.x + HalfDepth2
    LineStrPoints(4) = oPort2.y - HalfWidth2
    LineStrPoints(5) = oPort2.z
    
    LineStrPoints(6) = oPort2.x + HalfDepth2
    LineStrPoints(7) = oPort2.y + HalfWidth2
    LineStrPoints(8) = oPort2.z
    
    LineStrPoints(9) = oPort2.x - HalfDepth2
    LineStrPoints(10) = oPort2.y + HalfWidth2
    LineStrPoints(11) = oPort2.z
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set 0, 0, -1
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing
    
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
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort1 = Nothing
    Set oPortLocation = Nothing
    
' Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set 0, 0, 1
    oRadialOrient.Set -1, 0, 0
    
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort2 = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
