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
'   Author:         MS
'   Creation Date:  Wednesday, Nov 03 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  17.JUN.2007     dkl    CR-114880 Added implementation for part data basis for segemnted bends.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages
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
'   1) 45 degree Vertical Curved Upward Bend, Symmetrical, specified by Bend Radius and Tangent Length (Default implementation).
'   2) 45 degree Vertical Curved Upward Bend, Asymmetrical, specified by Bend Radius, Tangent Length 1 and Tangent Length 2.
'   3) 45 degree Vertical Segmented Upward Bend, Symmetrical, specified by Tangent Length.
'   4) 45 degree Vertical Segmented Upward Bend, Asymmetrical, specified by Tangent Length 1 and Tangent Length 2.
    
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
    
    Dim dCentoEndofTangent1 As Double
    Dim Angle As Double
    Angle = PI / 4
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth              As Double
    Dim HalfWidth              As Double
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
    
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 47    ' 47 - 45 degree Vertical Curved Upward Bend, Symmetrical, specified by
                            ' Bend Radius and Tangent Length (Default implementation).
            ' Retrieve Part properties: TangentLength
            parBendRadius = oTrayPart.BendRadius
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dCentoEndofTangent1 = (parBendRadius + HalfDepth) * Tan(Angle / 2)
            
            oPort1.Set -(dCentoEndofTangent1 + dTangentLength1), 0, 0
            oPort2.Set (dCentoEndofTangent1 + dTangentLength2) * Cos(Angle), 0, (dCentoEndofTangent1 + dTangentLength2) * Sin(Angle)
            
        Case 49  ' 49 - 45 degree Vertical Curved Upward Bend, Asymmetrical, specified by
                 'Bend Radius, Tangent Length 1 and Tangent Length 2.
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
            dCentoEndofTangent1 = (parBendRadius + HalfDepth) * Tan(Angle / 2)
            
            oPort1.Set -(dCentoEndofTangent1 + dTangentLength1), 0, 0
            oPort2.Set (dCentoEndofTangent1 + dTangentLength2) * Cos(Angle), 0, (dCentoEndofTangent1 + dTangentLength2) * Sin(Angle)
            
        Case 51     ' 51 - 45 degree Vertical Segmented Upward Bend, Symmetrical, specified by
                    ' Tangent Length.
            ' Retrieve Part properties: TangentLength
            parTangentLength = oTrayPart.TangentLength
            dCentoEndofTangent1 = HalfDepth / Tan(3 * PI / 8)
                       
            'Retrieved Tangent Length is along the outer edge. dTangentLength is computed along the
            'cable tray axis.
            dTangentLength1 = parTangentLength - 2 * dCentoEndofTangent1
            dTangentLength2 = parTangentLength - 2 * dCentoEndofTangent1
            
            oPort1.Set -(dCentoEndofTangent1 + dTangentLength1), 0, 0
            oPort2.Set (dCentoEndofTangent1 + dTangentLength2) * Cos(Angle), 0, (dCentoEndofTangent1 + dTangentLength2) * Sin(Angle)
            
        Case 53    ' 53 - 45 degree Vertical Segmented Upward Bend, Asymmetrical, specified by
                   ' Tangent Length 1 and Tangent Length 2.
            dCentoEndofTangent1 = HalfDepth / Tan(3 * PI / 8)
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
                                   
            'Retrieved Tangent Length is along the outer edge. dTangentLength is computed along the
            'cable tray axis.
            dTangentLength1 = arrayOfInputs(2) - 2 * dCentoEndofTangent1
            dTangentLength2 = arrayOfInputs(3) - 2 * dCentoEndofTangent1
                        
            oPort1.Set -(dCentoEndofTangent1 + dTangentLength1), 0, 0
            oPort2.Set (dCentoEndofTangent1 + dTangentLength2) * Cos(Angle), 0, (dCentoEndofTangent1 + dTangentLength2) * Sin(Angle)
                        
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
    
' Insert your code for output 1(Horizontal Tangent)

    Dim LineStrPoints(0 To 11)  As Double
    Dim ProjVector        As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
        
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
    Set oLineString = Nothing

' Insert your code for output 2(Vertical Bend)
    Dim LineStrCP As AutoMath.DPosition
    Set LineStrCP = New AutoMath.DPosition
    Dim CP As New AutoMath.DPosition
    LineStrCP.Set -dCentoEndofTangent1, 0, 0

    Select Case lCableTrayPartDataBasis
        
        Case Is <= 1, 47, 49
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
            
            CP.Set -dCentoEndofTangent1, 0, (parBendRadius + HalfDepth)
            Set objVertiCurve = PlaceRevolution(m_OutputColl, oLineString, ProjVector, _
                                                                        CP, Angle, False)
                       
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiCurve
            Set objVertiCurve = Nothing
            Set oLineString = Nothing
            Set CP = Nothing
            
        Case 51, 53
        
        ' Create the bend bottom
            Dim oSweepLineSegment As IngrGeom3D.Line3d
            Set oSweepLineSegment = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                    -dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
                    -dCentoEndofTangent1, HalfWidth, -HalfDepth)
            
            Dim oTraceLineSegment As IngrGeom3D.Line3d
            Set oTraceLineSegment = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                   -dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
                   dCentoEndofTangent1, -HalfWidth, -HalfDepth)
            Dim oTraceCurveColl  As Collection
            Set oTraceCurveColl = New Collection
            oTraceCurveColl.Add oTraceLineSegment
            
            Set oTraceLineSegment = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
                dCentoEndofTangent1 * (1 + 2 * Cos(PI / 4)), -HalfWidth, _
                            -HalfDepth + 2 * dCentoEndofTangent1 * Cos(PI / 4))
            oTraceCurveColl.Add oTraceLineSegment
        
            Dim oStPoint As AutoMath.DPosition
            Set oStPoint = New AutoMath.DPosition
            oStPoint.Set -dCentoEndofTangent1, -HalfWidth, -HalfDepth
            
            Dim objTraceCurve  As IngrGeom3D.ComplexString3d
            Set objTraceCurve = PlaceTrCString(oStPoint, oTraceCurveColl)
                
            Dim objBendBottom As IngrGeom3D.IJElements
            Dim stnorm() As Double
            Dim ednorm() As Double
                              
            Set objBendBottom = geomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                      objTraceCurve, oSweepLineSegment, 0, 0, stnorm, ednorm, False)
                      
            ' Set the output
            Dim ObjSurface As Object
            For Each ObjSurface In objBendBottom
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
            Next ObjSurface
            Set objBendBottom = Nothing
            Set ObjSurface = Nothing
            Set objTraceCurve = Nothing
            Set oSweepLineSegment = Nothing
            Set oTraceCurveColl = Nothing
            Set oTraceLineSegment = Nothing
            Set oStPoint = Nothing
    
        ' Create Bend Front Surface.
            Dim dPoints(0 To 11) As Double
            
            dPoints(0) = -dCentoEndofTangent1
            dPoints(1) = -HalfWidth
            dPoints(2) = HalfDepth
            
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = -HalfDepth
            
            dPoints(6) = dCentoEndofTangent1
            dPoints(7) = dPoints(1)
            dPoints(8) = dPoints(5)
            
            dPoints(9) = dCentoEndofTangent1 * (1 + 2 * Cos(PI / 4))
            dPoints(10) = dPoints(1)
            dPoints(11) = -HalfDepth + 2 * dCentoEndofTangent1 * Cos(PI / 4)
            
            Dim objBendFrontSurface As IngrGeom3D.Plane3d
            Set objBendFrontSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                    4, dPoints)
        
        ' Set the output
            m_OutputColl.AddOutput "BendFrontSurface", objBendFrontSurface
            Set objBendFrontSurface = Nothing
        
        ' Create Bend Rear Surface.
            
            dPoints(0) = -dCentoEndofTangent1
            dPoints(1) = HalfWidth
            dPoints(2) = HalfDepth
            
            dPoints(3) = dPoints(0)
            dPoints(4) = dPoints(1)
            dPoints(5) = -HalfDepth
            
            dPoints(6) = dCentoEndofTangent1
            dPoints(7) = dPoints(1)
            dPoints(8) = dPoints(5)
            
            dPoints(9) = dCentoEndofTangent1 * (1 + 2 * Cos(PI / 4))
            dPoints(10) = dPoints(1)
            dPoints(11) = -HalfDepth + 2 * dCentoEndofTangent1 * Cos(PI / 4)
        
            Dim objBendRearSurface As IngrGeom3D.Plane3d
            Set objBendRearSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                    4, dPoints)
        
        ' Set the output
            m_OutputColl.AddOutput "BendRearSurface", objBendRearSurface
            Set objBendRearSurface = Nothing
         
        Case Else
            GoTo ErrorLabel:
    
    End Select

' Insert your code for output 3(Vertical Tangent)
    
    LineStrPoints(0) = (dCentoEndofTangent1) * Cos(Angle) - HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(1) = oPort2.y - HalfWidth
    LineStrPoints(2) = (dCentoEndofTangent1) * Sin(Angle) + HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(3) = (dCentoEndofTangent1) * Cos(Angle) + HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(4) = oPort2.y - HalfWidth
    LineStrPoints(5) = (dCentoEndofTangent1) * Sin(Angle) - HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(6) = (dCentoEndofTangent1) * Cos(Angle) + HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(7) = oPort2.y + HalfWidth
    LineStrPoints(8) = (dCentoEndofTangent1) * Sin(Angle) - HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(9) = (dCentoEndofTangent1) * Cos(Angle) - HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(10) = oPort2.y + HalfWidth
    LineStrPoints(11) = (dCentoEndofTangent1) * Sin(Angle) + HalfDepth * Sin(PI / 2 - Angle)

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set Cos(Angle), 0, Sin(Angle)
    
    
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength2, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing
    Set ProjVector = Nothing
    Set oLineString = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1
'    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    
    ' Set oPlacePoint = New AutoMath.DPosition
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
    oDir.Set Cos(Angle), 0, Sin(Angle)
    oRadialOrient.Set -Sin(Angle), 0, Cos(Angle)
    Set oPortLocation = New AutoMath.DPosition
    
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
