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
'   Veritical CableTray Inside consist of a Horizontal Tangent and Veritical Curve and Vertical Tangent
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    
    iOutput = 0
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
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
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If parTangentLength = 0 Then parTangentLength = 0.000001
    

' Insert your code for output 1(Horizontal Tangent)

    Dim LineStrPoints(0 To 11)  As Double
    Dim Angle As Double
    Dim ProjVector        As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Angle = PI / 3
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth              As Double
    Dim HalfWidth              As Double
    
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
    
    Dim dcentoPort1 As Double
    
    dcentoPort1 = (parBendRadius + HalfDepth) * Tan(Angle / 2)
        
    oPort1.Set -(dcentoPort1 + parTangentLength), 0, 0

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
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, parTangentLength, False)
    'Remove cable tray line string U-shape
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing

' Insert your code for output 2(Vertical Bend)
    Dim LineStrCP As New AutoMath.DPosition
    Dim CP As New AutoMath.DPosition
    LineStrCP.Set -dcentoPort1, 0, 0

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
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set 0, -1, 0
    
    CP.Set -dcentoPort1, 0, (parBendRadius + HalfDepth)
    Set objVertiCurve = PlaceRevolution(m_OutputColl, oLineString, ProjVector, CP, Angle, False)
    'Remove cable tray line string U-shape
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiCurve
    Set objVertiCurve = Nothing

' Insert your code for output 3(Vertical Tangent)
    oPort2.Set (dcentoPort1 + parTangentLength) * Cos(Angle), 0, (dcentoPort1 + parTangentLength) * Sin(Angle)
    
    LineStrPoints(0) = (dcentoPort1) * Cos(Angle) - HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(1) = oPort2.y - HalfWidth
    LineStrPoints(2) = (dcentoPort1) * Sin(Angle) + HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(3) = (dcentoPort1) * Cos(Angle) + HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(4) = oPort2.y - HalfWidth
    LineStrPoints(5) = (dcentoPort1) * Sin(Angle) - HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(6) = (dcentoPort1) * Cos(Angle) + HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(7) = oPort2.y + HalfWidth
    LineStrPoints(8) = (dcentoPort1) * Sin(Angle) - HalfDepth * Sin(PI / 2 - Angle)

    LineStrPoints(9) = (dcentoPort1) * Cos(Angle) - HalfDepth * Cos(PI / 2 - Angle)
    LineStrPoints(10) = oPort2.y + HalfWidth
    LineStrPoints(11) = (dcentoPort1) * Sin(Angle) + HalfDepth * Sin(PI / 2 - Angle)

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set Cos(Angle), 0, Sin(Angle)
    
    
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, parTangentLength, False)
    'Remove cable tray line string U-shape
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing
    
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
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set Cos(Angle), 0, Sin(Angle)

    oRadialOrient.Set -Sin(Angle), 0, Cos(Angle)

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
