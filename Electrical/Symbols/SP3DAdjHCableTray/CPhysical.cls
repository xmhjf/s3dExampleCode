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
'   Creation Date:  Friday, Mar 1 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjElbow As Object
    Dim ObjInclTangent  As Object

    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim ProjVector As New AutoMath.DVector
    Dim LineStrCP As New AutoMath.DPosition
    Dim parAdjBendAngle As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parAdjBendAngle = arrayOfInputs(2)
    
    iOutput = 0
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
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)

' Insert your code for output 1(U Shape Horizontal Tangent)
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim HD              As Double
    Dim HW              As Double
    Dim LineStrPoints(0 To 11)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    Dim parFacetoCenter As Double
    parFacetoCenter = (parBendRadius + HW) * Tan(parAdjBendAngle / 2) + parTangentLength
    
    oPort1.Set -parFacetoCenter, 0, 0
    
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
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set 1, 0, 0
    Set ObjHoriTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, parTangentLength, False)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTangent
    Set ObjHoriTangent = Nothing
'   Remove cable tray line string U-shape
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
 ' Insert your code for output 2(U Shape Elbow)
    Dim oElbowLineStr As IngrGeom3D.LineString3d
    LineStrCP.Set -parFacetoCenter + parTangentLength, 0, 0
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
    Set oElbowLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set 0, 0, 1
    CP.Set -parFacetoCenter + parTangentLength, parBendRadius + HW, 0
    Set ObjElbow = PlaceRevolution(m_OutputColl, oElbowLineStr, ProjVector, CP, parAdjBendAngle, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
'   Remove cable tray line string U-shape
    Set objLineString = oElbowLineStr
    Set oElbowLineStr = Nothing
    objLineString.Remove
    
' Insert your code for output 3(U Shape Inclined Tangent)
    Dim oInclinedLineStr As IngrGeom3D.LineString3d
    oPort2.Set parFacetoCenter * Cos(parAdjBendAngle), parFacetoCenter * Sin(parAdjBendAngle), 0
    
    LineStrPoints(0) = oPort2.x - HW * Sin(parAdjBendAngle)
    LineStrPoints(1) = oPort2.y + HW * Cos(parAdjBendAngle)
    LineStrPoints(2) = oPort2.z + HD
    
    LineStrPoints(3) = oPort2.x - HW * Sin(parAdjBendAngle)
    LineStrPoints(4) = oPort2.y + HW * Cos(parAdjBendAngle)
    LineStrPoints(5) = oPort2.z - HD
    
    LineStrPoints(6) = oPort2.x + HW * Sin(parAdjBendAngle)
    LineStrPoints(7) = oPort2.y - HW * Cos(parAdjBendAngle)
    LineStrPoints(8) = oPort2.z - HD
    
    LineStrPoints(9) = oPort2.x + HW * Sin(parAdjBendAngle)
    LineStrPoints(10) = oPort2.y - HW * Cos(parAdjBendAngle)
    LineStrPoints(11) = oPort2.z + HD
    Set oInclinedLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set -Cos(parAdjBendAngle), -Sin(parAdjBendAngle), 0
    Set ObjInclTangent = PlaceProjection(m_OutputColl, oInclinedLineStr, ProjVector, parTangentLength, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInclTangent
    Set ObjInclTangent = Nothing
    Set ProjVector = Nothing
'   Remove cable tray line string U-shape
    Set objLineString = oInclinedLineStr
    Set oInclinedLineStr = Nothing
    objLineString.Remove

    
' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

'   oPlacePoint.Set -(parBendRadius + parActualWidth / 2), 0, 0
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
'   Set oPlacePoint = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing

' Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    
    oDir.Set Cos(parAdjBendAngle), Sin(parAdjBendAngle), 0
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





