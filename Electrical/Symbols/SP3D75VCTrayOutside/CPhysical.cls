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
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is 75 Degree Vertical CableTray Outside symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------   -----     ------------------
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
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double

    iOutput = 0

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
        
    Dim dCentoEndofTangent1 As Double
    Dim dAngle As Double
    dAngle = 75 * (PI / 180)
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth As Double
    Dim HalfWidth As Double
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
 
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    Dim dTangentLength As Double
    
    parTangentLength = oTrayPart.TangentLength
    parBendRadius = oTrayPart.BendRadius
    dTangentLength = (parBendRadius) * Tan(dAngle / 2)
    dCentoEndofTangent1 = HalfDepth * Tan(dAngle / 2)

    'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength, 0) Then dTangentLength = 0.0001
 
    'Insert your code for output 1(Vertical Tangent)
    Dim dLineStrPoints(0 To 11)  As Double
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oPort1 As AutoMath.DPosition  'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Set oPort2 = New AutoMath.DPosition

    oPort1.Set 0, 0, -(dCentoEndofTangent1 + dTangentLength)
    
    dLineStrPoints(0) = oPort1.x - HalfDepth
    dLineStrPoints(1) = oPort1.y - HalfWidth
    dLineStrPoints(2) = oPort1.z
    
    dLineStrPoints(3) = oPort1.x + HalfDepth
    dLineStrPoints(4) = oPort1.y - HalfWidth
    dLineStrPoints(5) = oPort1.z
    
    dLineStrPoints(6) = oPort1.x + HalfDepth
    dLineStrPoints(7) = oPort1.y + HalfWidth
    dLineStrPoints(8) = oPort1.z
    
    dLineStrPoints(9) = oPort1.x - HalfDepth
    dLineStrPoints(10) = oPort1.y + HalfWidth
    dLineStrPoints(11) = oPort1.z

    oProjVector.Set 0, 0, 1
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    Dim objVertiTangent As Object
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, dTangentLength, False)
    'Remove cable tray line string U-shape
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing

    'Insert your code for output 2(Vertical Bend)
    'Create Bend Rear Surface.
    Dim dPoints(0 To 11) As Double

    dPoints(0) = -HalfDepth
    dPoints(1) = HalfWidth
    dPoints(2) = -dCentoEndofTangent1
    
    dPoints(3) = HalfDepth
    dPoints(4) = dPoints(1)
    dPoints(5) = dPoints(2)
    
    dPoints(6) = -(HalfDepth - 2 * dCentoEndofTangent1 * Sin(dAngle))
    dPoints(7) = dPoints(1)
    dPoints(8) = dCentoEndofTangent1 * (1 + 2 * Cos(dAngle))
    
    dPoints(9) = dPoints(0)
    dPoints(10) = dPoints(1)
    dPoints(11) = dCentoEndofTangent1
            
    Dim objBendRearSurface As IngrGeom3D.Plane3d
    Set objBendRearSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBendRearSurface
    Set objBendRearSurface = Nothing

    'Create Bend Front Surface.
    dPoints(0) = -HalfDepth
    dPoints(1) = -HalfWidth
    dPoints(2) = -dCentoEndofTangent1
    
    dPoints(3) = HalfDepth
    dPoints(4) = dPoints(1)
    dPoints(5) = dPoints(2)
    
    dPoints(6) = -(HalfDepth - 2 * dCentoEndofTangent1 * Sin(dAngle))
    dPoints(7) = dPoints(1)
    dPoints(8) = dCentoEndofTangent1 * (1 + 2 * Cos(dAngle))
    
    dPoints(9) = dPoints(0)
    dPoints(10) = dPoints(1)
    dPoints(11) = dCentoEndofTangent1

    Dim objBendFrontSurface As IngrGeom3D.Plane3d
    Set objBendFrontSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBendFrontSurface
    Set objBendFrontSurface = Nothing
        
    'Insert your code for output 3(Inclined Tangent)
    oPort2.Set (dCentoEndofTangent1 + dTangentLength) * Cos(PI / 2 - dAngle), _
                0, (dCentoEndofTangent1 + dTangentLength) * Sin(PI / 2 - dAngle)
                
    dLineStrPoints(0) = (dCentoEndofTangent1) * Cos(PI / 2 - dAngle) - HalfDepth * Cos(dAngle)
    dLineStrPoints(1) = oPort2.y - HalfWidth
    dLineStrPoints(2) = (dCentoEndofTangent1) * Sin(PI / 2 - dAngle) + HalfDepth * Sin(dAngle)

    dLineStrPoints(3) = (dCentoEndofTangent1) * Cos(PI / 2 - dAngle) + HalfDepth * Cos(dAngle)
    dLineStrPoints(4) = oPort2.y - HalfWidth
    dLineStrPoints(5) = (dCentoEndofTangent1) * Sin(PI / 2 - dAngle) - HalfDepth * Sin(dAngle)

    dLineStrPoints(6) = (dCentoEndofTangent1) * Cos(PI / 2 - dAngle) + HalfDepth * Cos(dAngle)
    dLineStrPoints(7) = oPort2.y + HalfWidth
    dLineStrPoints(8) = (dCentoEndofTangent1) * Sin(PI / 2 - dAngle) - HalfDepth * Sin(dAngle)

    dLineStrPoints(9) = (dCentoEndofTangent1) * Cos(PI / 2 - dAngle) - HalfDepth * Cos(dAngle)
    dLineStrPoints(10) = oPort2.y + HalfWidth
    dLineStrPoints(11) = (dCentoEndofTangent1) * Sin(PI / 2 - dAngle) + HalfDepth * Sin(dAngle)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPoints)
    oProjVector.Set Cos(PI / 2 - dAngle), 0, Sin(PI / 2 - dAngle)
    
    Dim objHoriTangent As Object
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, oProjVector, dTangentLength, False)
    'Remove cable tray line string U-shape
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing
    Set oProjVector = Nothing
    Set oGeomFactory = Nothing
    
    'Place Nozzle 1
    Dim oDir As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort  As GSCADNozzleEntities.IJCableTrayPortOcc
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    oDir.Set 0, 0, -1
    oRadialOrient.Set -1, 0, 0
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
    
    oDir.Set Cos(PI / 2 - dAngle), 0, Sin(PI / 2 - dAngle)
    oRadialOrient.Set -Cos(dAngle), 0, Sin(dAngle)
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
