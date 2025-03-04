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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Tuesday, Jun 24 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   24.JUN.2008     MA     CR-144224 Created the symbol
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
    
    Dim iOutput     As Double
    Dim ObjHoriTangent  As Object
    Dim ObjProj As IJDObject
    Dim ObjProj1 As IJDObject
    Dim ObjInclTangent  As Object
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As AutoMath.DPosition  'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    Dim oLineStrCP As AutoMath.DPosition
    Set oLineStrCP = New AutoMath.DPosition
    Dim lCableTrayPartDataBasis As Long
    Dim dCTDimension As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    iOutput = 0
    parBendRadius = oTrayPart.BendRadius
    parTangentLength = oTrayPart.TangentLength
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(parTangentLength, 0) Then parTangentLength = 0.000001
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)

    'Insert your code for output 1(U Shape Horizontal Tangent)
    Dim HD As Double
    Dim HW As Double
    Dim dLineStrPoints(0 To 11)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim dAngle As Double
    dAngle = 75 * (PI / 180)
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    
    Dim dFacetoCenter As Double
    dCTDimension = (parBendRadius + parActualWidth) * Tan(dAngle / 2)
    dFacetoCenter = dCTDimension - parActualWidth / (2 * Tan(52.5 * PI / 180)) ' Angle 52.5 (105/2)
    oPort1.Set -dFacetoCenter, 0, 0
    
    dLineStrPoints(0) = oPort1.x
    dLineStrPoints(1) = oPort1.y + HW
    dLineStrPoints(2) = oPort1.z + HD
    
    dLineStrPoints(3) = oPort1.x
    dLineStrPoints(4) = oPort1.y + HW
    dLineStrPoints(5) = oPort1.z - HD
    
    dLineStrPoints(6) = oPort1.x
    dLineStrPoints(7) = oPort1.y - HW
    dLineStrPoints(8) = oPort1.z - HD
    
    dLineStrPoints(9) = oPort1.x
    dLineStrPoints(10) = oPort1.y - HW
    dLineStrPoints(11) = oPort1.z + HD
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    oProjVector.Set 1, 0, 0
    Set ObjHoriTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, parTangentLength, False)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriTangent
    Set ObjHoriTangent = Nothing
    'Remove cable tray line string U-shape
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
    'Insert your code for output 2(U Shape Plane apart form tangents)
    Dim oProjLineStr As IngrGeom3D.LineString3d
    oLineStrCP.Set -dFacetoCenter + parTangentLength, 0, -HD
    Dim dLineStrPoints1(0 To 17) As Double
    
    dLineStrPoints1(0) = oLineStrCP.x
    dLineStrPoints1(1) = oLineStrCP.y + HW
    dLineStrPoints1(2) = oLineStrCP.z
    
    dLineStrPoints1(3) = oLineStrCP.x
    dLineStrPoints1(4) = oLineStrCP.y - HW
    dLineStrPoints1(5) = oLineStrCP.z
    
    dLineStrPoints1(6) = HW * Tan(dAngle / 2)
    dLineStrPoints1(7) = -HW
    dLineStrPoints1(8) = oLineStrCP.z
    
    dLineStrPoints1(9) = dFacetoCenter * Cos(dAngle) + HW * Sin(dAngle) - parTangentLength * Cos(dAngle)
    dLineStrPoints1(10) = dFacetoCenter * Sin(dAngle) - HW * Cos(dAngle) - parTangentLength * Sin(dAngle)
    dLineStrPoints1(11) = oLineStrCP.z
    
    dLineStrPoints1(12) = dFacetoCenter * Cos(dAngle) - HW * Sin(dAngle) - parTangentLength * Cos(dAngle)
    dLineStrPoints1(13) = dFacetoCenter * Sin(dAngle) + HW * Cos(dAngle) - parTangentLength * Sin(dAngle)
    dLineStrPoints1(14) = oLineStrCP.z
    
    dLineStrPoints1(15) = -HW * Tan(dAngle / 2)
    dLineStrPoints1(16) = HW
    dLineStrPoints1(17) = oLineStrCP.z
    
    Dim oHLine As IngrGeom3D.Line3d
    Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(3), dLineStrPoints1(4), dLineStrPoints1(5), dLineStrPoints1(6), dLineStrPoints1(7), dLineStrPoints1(8))
    
    Dim oILine1 As IngrGeom3D.Line3d
    Set oILine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(6), dLineStrPoints1(7), dLineStrPoints1(8), dLineStrPoints1(9), dLineStrPoints1(10), dLineStrPoints1(11))
    
    Dim oILine2 As IngrGeom3D.Line3d
    Set oILine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(9), dLineStrPoints1(10), dLineStrPoints1(11), dLineStrPoints1(12), dLineStrPoints1(13), dLineStrPoints1(14))
    
    Dim oILine3  As IngrGeom3D.Line3d
    Set oILine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(12), dLineStrPoints1(13), dLineStrPoints1(14), dLineStrPoints1(15), dLineStrPoints1(16), dLineStrPoints1(17))
    
    Dim oHLine1  As IngrGeom3D.Line3d
    Set oHLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(15), dLineStrPoints1(16), dLineStrPoints1(17), dLineStrPoints1(0), dLineStrPoints1(1), dLineStrPoints1(2))
    
    Dim oVLine   As IngrGeom3D.Line3d
    Set oVLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLineStrPoints1(0), dLineStrPoints1(1), dLineStrPoints1(2), dLineStrPoints1(3), dLineStrPoints1(4), dLineStrPoints1(5))
    
    Dim oLiCollection As Collection
    Set oLiCollection = New Collection
    oLiCollection.Add oHLine
    oLiCollection.Add oILine1
    Set oProjVector = New AutoMath.DVector
    Dim objBp1scurve  As IngrGeom3D.ComplexString3d
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    
    oStPoint.Set dLineStrPoints1(3), dLineStrPoints1(4), dLineStrPoints1(5)
    Set objBp1scurve = PlaceTrCString(oStPoint, oLiCollection)
    
    oProjVector.Set 0, 0, 1
    
    Set ObjProj = PlaceProjection(m_OutputColl, objBp1scurve, oProjVector, parActualDepth, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProj
    Set ObjProj = Nothing
    Set oLiCollection = Nothing
    Set oLineStrCP = Nothing
    
    Dim oLiCollection1 As Collection
    Set oLiCollection1 = New Collection
    oLiCollection1.Add oILine3
    oLiCollection1.Add oHLine1
    Set oStPoint = New AutoMath.DPosition
    
    oStPoint.Set dLineStrPoints1(12), dLineStrPoints1(13), dLineStrPoints1(14)
    Set objBp1scurve = New IngrGeom3D.ComplexString3d
    Set objBp1scurve = PlaceTrCString(oStPoint, oLiCollection1)
    Set ObjProj1 = PlaceProjection(m_OutputColl, objBp1scurve, oProjVector, parActualDepth, False)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProj1
    Set ObjProj1 = Nothing
    Set objBp1scurve = Nothing
    Set oLiCollection1 = Nothing
    
    Dim oBranchcurves As Collection
    Set oBranchcurves = New Collection
    oBranchcurves.Add oHLine
    oBranchcurves.Add oILine1
    oBranchcurves.Add oILine2
    oBranchcurves.Add oILine3
    oBranchcurves.Add oHLine1
    oBranchcurves.Add oVLine

    Dim objHBcurves As IngrGeom3D.ComplexString3d
    oStPoint.Set dLineStrPoints1(3), dLineStrPoints1(4), dLineStrPoints1(5)
    Set objHBcurves = PlaceTrCString(oStPoint, oBranchcurves)
    
    Dim oTrayBottom As IngrGeom3D.Plane3d
    Set oTrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -dFacetoCenter, 0, -HD, oProjVector.x, oProjVector.y, oProjVector.z)
    Call oTrayBottom.AddBoundary(objHBcurves)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTrayBottom
    Set oTrayBottom = Nothing
    Set objHBcurves = Nothing
    Set oBranchcurves = Nothing
    Set oStPoint = Nothing
    
    'Insert your code for output 3(U Shape Inclined Tangent)
    Dim oInclinedLineStr As IngrGeom3D.LineString3d
    oPort2.Set dFacetoCenter * Cos(dAngle), dFacetoCenter * Sin(dAngle), 0
    
    dLineStrPoints(0) = oPort2.x - HW * Sin(dAngle)
    dLineStrPoints(1) = oPort2.y + HW * Cos(dAngle)
    dLineStrPoints(2) = oPort2.z + HD
    
    dLineStrPoints(3) = oPort2.x - HW * Sin(dAngle)
    dLineStrPoints(4) = oPort2.y + HW * Cos(dAngle)
    dLineStrPoints(5) = oPort2.z - HD
    
    dLineStrPoints(6) = oPort2.x + HW * Sin(dAngle)
    dLineStrPoints(7) = oPort2.y - HW * Cos(dAngle)
    dLineStrPoints(8) = oPort2.z - HD
    
    dLineStrPoints(9) = oPort2.x + HW * Sin(dAngle)
    dLineStrPoints(10) = oPort2.y - HW * Cos(dAngle)
    dLineStrPoints(11) = oPort2.z + HD
    
    Set oInclinedLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    oProjVector.Set -Cos(dAngle), -Sin(dAngle), 0
    Set ObjInclTangent = PlaceProjection(m_OutputColl, oInclinedLineStr, oProjVector, parTangentLength, False)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInclTangent
    Set ObjInclTangent = Nothing
    Set oProjVector = Nothing
    'Remove cable tray line string U-shape
    Set objLineString = oInclinedLineStr
    Set oInclinedLineStr = Nothing
    objLineString.Remove
    Set oGeomFactory = Nothing
    
    'Place Nozzle 1
    Dim oDir As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort As GSCADNozzleEntities.IJCableTrayPortOcc
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector

    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
    
    'Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oDir.Set Cos(dAngle), Sin(dAngle), 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

    'Set the output
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
