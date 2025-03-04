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
'   -----------    -----   ------------------
'   24.JUN.2008     MA     CR-144224 Created the symbol
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
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
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
    
    Dim dCentoEndofTangent1 As Double
    Dim dAngle As Double
    dAngle = 75 * (PI / 180)
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth As Double
    Dim HalfWidth As Double
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
    
    parTangentLength = oTrayPart.TangentLength
    parTangentLength = (parBendRadius) * Tan(dAngle / 2)
    dCentoEndofTangent1 = HalfDepth * Tan(dAngle / 2)
    
    'Retrieved Tangent Length is along the outer edge. dTangentLength is computed along the
    'cable tray axis.
    Dim dTangentLength As Double
    dTangentLength = parTangentLength - 2 * dCentoEndofTangent1
    
    oPort1.Set -(dCentoEndofTangent1 + dTangentLength), 0, 0
    oPort2.Set (dCentoEndofTangent1 + dTangentLength) * Cos(dAngle), 0, (dCentoEndofTangent1 + dTangentLength) * Sin(dAngle)
            
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength, 0) Then dTangentLength = 0.0001
    
    'Insert your code for output 1(Horizontal Tangent)
    Dim dLineStrPoints(0 To 11)  As Double
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
    dLineStrPoints(0) = oPort1.x
    dLineStrPoints(1) = oPort1.y - HalfWidth
    dLineStrPoints(2) = oPort1.z + HalfDepth
    
    dLineStrPoints(3) = oPort1.x
    dLineStrPoints(4) = oPort1.y - HalfWidth
    dLineStrPoints(5) = oPort1.z - HalfDepth
    
    dLineStrPoints(6) = oPort1.x
    dLineStrPoints(7) = oPort1.y + HalfWidth
    dLineStrPoints(8) = oPort1.z - HalfDepth
    
    dLineStrPoints(9) = oPort1.x
    dLineStrPoints(10) = oPort1.y + HalfWidth
    dLineStrPoints(11) = oPort1.z + HalfDepth
    oProjVector.Set 1, 0, 0
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, dTangentLength, False)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing
    Set oLineString = Nothing

    'Insert your code for output 2(Vertical Bend)
    'Create the bend bottom
    Dim oSweepLineSegment As IngrGeom3D.Line3d
    Set oSweepLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
            -dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
            -dCentoEndofTangent1, HalfWidth, -HalfDepth)
    
    Dim oTraceLineSegment As IngrGeom3D.Line3d
    Set oTraceLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
           -dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
           dCentoEndofTangent1, -HalfWidth, -HalfDepth)
    Dim oTraceCurveColl  As Collection
    Set oTraceCurveColl = New Collection
    oTraceCurveColl.Add oTraceLineSegment
            
    Set oTraceLineSegment = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dCentoEndofTangent1, -HalfWidth, -HalfDepth, _
        dCentoEndofTangent1 * (1 + 2 * Cos(dAngle)), -HalfWidth, _
                   -(HalfDepth - 2 * dCentoEndofTangent1 * Sin(dAngle)))
            
    oTraceCurveColl.Add oTraceLineSegment

    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set -dCentoEndofTangent1, -HalfWidth, -HalfDepth
    
    Dim objTraceCurve  As IngrGeom3D.ComplexString3d
    Set objTraceCurve = PlaceTrCString(oStPoint, oTraceCurveColl)
        
    Dim objBendBottom As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
                              
    Set objBendBottom = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
              objTraceCurve, oSweepLineSegment, 0, 0, stnorm, ednorm, False)
              
    'Set the output
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
    
    'Create Bend Front Surface.
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
    
    dPoints(9) = dCentoEndofTangent1 * (1 + 2 * Cos(dAngle))
    dPoints(10) = dPoints(1)
    dPoints(11) = -(HalfDepth - 2 * dCentoEndofTangent1 * Sin(dAngle))
    
    Dim objBendFrontSurface As IngrGeom3D.Plane3d
    Set objBendFrontSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBendFrontSurface
    Set objBendFrontSurface = Nothing
        
    'Create Bend Rear Surface.
    dPoints(0) = -dCentoEndofTangent1
    dPoints(1) = HalfWidth
    dPoints(2) = HalfDepth
    
    dPoints(3) = dPoints(0)
    dPoints(4) = dPoints(1)
    dPoints(5) = -HalfDepth
    
    dPoints(6) = dCentoEndofTangent1
    dPoints(7) = dPoints(1)
    dPoints(8) = dPoints(5)
    
    dPoints(9) = dCentoEndofTangent1 * (1 + 2 * Cos(dAngle))
    dPoints(10) = dPoints(1)
    dPoints(11) = -(HalfDepth - 2 * dCentoEndofTangent1 * Sin(dAngle))
    
    Dim objBendRearSurface As IngrGeom3D.Plane3d
    Set objBendRearSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBendRearSurface
    Set objBendRearSurface = Nothing
         
    'Insert your code for output 3(Vertical Tangent)
    dLineStrPoints(0) = (dCentoEndofTangent1) * Cos(dAngle) - HalfDepth * Cos(PI / 2 - dAngle)
    dLineStrPoints(1) = oPort2.y - HalfWidth
    dLineStrPoints(2) = (dCentoEndofTangent1) * Sin(dAngle) + HalfDepth * Sin(PI / 2 - dAngle)

    dLineStrPoints(3) = (dCentoEndofTangent1) * Cos(dAngle) + HalfDepth * Cos(PI / 2 - dAngle)
    dLineStrPoints(4) = oPort2.y - HalfWidth
    dLineStrPoints(5) = (dCentoEndofTangent1) * Sin(dAngle) - HalfDepth * Sin(PI / 2 - dAngle)

    dLineStrPoints(6) = (dCentoEndofTangent1) * Cos(dAngle) + HalfDepth * Cos(PI / 2 - dAngle)
    dLineStrPoints(7) = oPort2.y + HalfWidth
    dLineStrPoints(8) = (dCentoEndofTangent1) * Sin(dAngle) - HalfDepth * Sin(PI / 2 - dAngle)

    dLineStrPoints(9) = (dCentoEndofTangent1) * Cos(dAngle) - HalfDepth * Cos(PI / 2 - dAngle)
    dLineStrPoints(10) = oPort2.y + HalfWidth
    dLineStrPoints(11) = (dCentoEndofTangent1) * Sin(dAngle) + HalfDepth * Sin(PI / 2 - dAngle)

    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    oProjVector.Set Cos(dAngle), 0, Sin(dAngle)
    
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, dTangentLength, False)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing
    Set oProjVector = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing
    
    'Place Nozzle 1
    Dim oDir As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort  As GSCADNozzleEntities.IJCableTrayPortOcc
    
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
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort1 = Nothing
    Set oPortLocation = Nothing
    
    'Place Nozzle 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oDir.Set Cos(dAngle), 0, Sin(dAngle)
    oRadialOrient.Set -Sin(dAngle), 0, Cos(dAngle)
    Set oPortLocation = New AutoMath.DPosition
    
    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

    'Set the output
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
