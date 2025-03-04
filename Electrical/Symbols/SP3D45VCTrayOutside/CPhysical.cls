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
'   Author:         BBA /MS
'   Creation Date:  Tuesday, Nov 2 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is 45 Degree Vertical CableTray Outside symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  13.Jun.2006     MA   CR-114880 Implemented Part data basis for values 55,57,59 and 61.
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
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double

    iOutput = 0

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parTangentLength1 = arrayOfInputs(2)
'    parTangentLength2 = arrayOfInputs(3)
    
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
        
    Dim dCentoEndofTangent1 As Double
    Dim Angle As Double
    Angle = PI / 4
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
    Dim HalfDepth              As Double
    Dim HalfWidth              As Double
    HalfDepth = parActualDepth / 2
    HalfWidth = parActualWidth / 2
 
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -55):45 degree Vertical Curved Downward Bend, Symmetrical, specified
'                                   by Bend Radius and Tangent Length
'   (Part data Basis value -57):45 degree Vertical Curved Downward Bend, Asymmetrical, specified
'                                   by Bend Radius, Tangent Length 1 and Tangent Length 2
'   (Part data Basis value -59):45 degree Vertical Segmented Downward Bend, Symmetrical, specified
'                                   by Tangent Length
'   (Part data Basis value -61):45 degree Vertical Segmented Downward Bend, Asymmetrical, specified
'                                   by Tangent Length 1 and Tangent Length 2
    
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
 
        Case Is <= 1, 55
            parBendRadius = oTrayPart.BendRadius
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dCentoEndofTangent1 = (parBendRadius + HalfDepth) * Tan(Angle / 2)
 
        Case 57
            parBendRadius = oTrayPart.BendRadius
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
            dCentoEndofTangent1 = (parBendRadius + HalfDepth) * Tan(Angle / 2)
 
        Case 59
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dCentoEndofTangent1 = HalfDepth / Tan(3 * PI / 8)
        
        Case 61
            dTangentLength1 = arrayOfInputs(2)
            dTangentLength2 = arrayOfInputs(3)
            dCentoEndofTangent1 = HalfDepth / Tan(3 * PI / 8)
        
        Case Else
            GoTo ErrorLabel:
            
    End Select
 
'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
 
 ' Insert your code for output 1(Vertical Tangent)
    Dim LineStrPoints(0 To 11)  As Double
    Dim ProjVector        As AutoMath.DVector
    Set ProjVector = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim oPort1 As AutoMath.DPosition  'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    Dim oPort2 As AutoMath.DPosition  'Port 2 center point
    Set oPort2 = New AutoMath.DPosition

    oPort1.Set 0, 0, -(dCentoEndofTangent1 + dTangentLength1)
    
    LineStrPoints(0) = oPort1.x - HalfDepth
    LineStrPoints(1) = oPort1.y - HalfWidth
    LineStrPoints(2) = oPort1.z
    
    LineStrPoints(3) = oPort1.x + HalfDepth
    LineStrPoints(4) = oPort1.y - HalfWidth
    LineStrPoints(5) = oPort1.z
    
    LineStrPoints(6) = oPort1.x + HalfDepth
    LineStrPoints(7) = oPort1.y + HalfWidth
    LineStrPoints(8) = oPort1.z
    
    LineStrPoints(9) = oPort1.x - HalfDepth
    LineStrPoints(10) = oPort1.y + HalfWidth
    LineStrPoints(11) = oPort1.z

    ProjVector.Set 0, 0, 1
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    Dim objVertiTangent As Object
    Set objVertiTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength1, False)
    'Remove cable tray line string U-shape
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiTangent
    Set objVertiTangent = Nothing

' Insert your code for output 2(Vertical Bend)
    Dim LineStrCP As AutoMath.DPosition
    Set LineStrCP = New AutoMath.DPosition
    Dim CP As New AutoMath.DPosition
    Dim RevVector        As AutoMath.DVector
    Set RevVector = New AutoMath.DVector
    LineStrCP.Set 0, 0, -dCentoEndofTangent1
    
    Select Case lCableTrayPartDataBasis

        Case Is <= 1, 55, 57
            LineStrPoints(0) = LineStrCP.x - HalfDepth
            LineStrPoints(1) = LineStrCP.y - HalfWidth
            LineStrPoints(2) = LineStrCP.z
            
            LineStrPoints(3) = LineStrCP.x + HalfDepth
            LineStrPoints(4) = LineStrCP.y - HalfWidth
            LineStrPoints(5) = LineStrCP.z
            
            LineStrPoints(6) = LineStrCP.x + HalfDepth
            LineStrPoints(7) = LineStrCP.y + HalfWidth
            LineStrPoints(8) = LineStrCP.z
            
            LineStrPoints(9) = LineStrCP.x - HalfDepth
            LineStrPoints(10) = LineStrCP.y + HalfWidth
            LineStrPoints(11) = LineStrCP.z

            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
            RevVector.Set 0, 1, 0

            CP.Set (parBendRadius + HalfDepth), 0, -dCentoEndofTangent1
        
            Dim objVertiCurve As Object
            Set objVertiCurve = PlaceRevolution(m_OutputColl, oLineString, RevVector, CP, Angle, False)
            'Remove cable tray line string U-shape
            Set objLineString = oLineString
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
        ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertiCurve
            Set objVertiCurve = Nothing
            Set CP = Nothing
            Set LineStrCP = Nothing
            Set RevVector = Nothing

        Case 59, 61
         ' Create Bend Rear Surface.
            Dim dPoints(0 To 11) As Double
        
            dPoints(0) = -HalfDepth
            dPoints(1) = HalfWidth
            dPoints(2) = -dCentoEndofTangent1
            
            dPoints(3) = HalfDepth
            dPoints(4) = dPoints(1)
            dPoints(5) = dPoints(2)
            
            dPoints(6) = -HalfDepth + 2 * dCentoEndofTangent1 * Cos(PI / 4)
            dPoints(7) = dPoints(1)
            dPoints(8) = dCentoEndofTangent1 * (1 + 2 * Cos(PI / 4))
            
            dPoints(9) = dPoints(0)
            dPoints(10) = dPoints(1)
            dPoints(11) = dCentoEndofTangent1
            
            Dim objBendRearSurface As IngrGeom3D.Plane3d
            Set objBendRearSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                    4, dPoints)
        
        ' Set the output
            m_OutputColl.AddOutput "BendRearSurface", objBendRearSurface
            Set objBendRearSurface = Nothing
        
        ' Create Bend Front Surface.
        
            dPoints(0) = -HalfDepth
            dPoints(1) = -HalfWidth
            dPoints(2) = -dCentoEndofTangent1
            
            dPoints(3) = HalfDepth
            dPoints(4) = dPoints(1)
            dPoints(5) = dPoints(2)
            
            dPoints(6) = -HalfDepth + 2 * dCentoEndofTangent1 * Cos(PI / 4)
            dPoints(7) = dPoints(1)
            dPoints(8) = dCentoEndofTangent1 * (1 + 2 * Cos(PI / 4))
            
            dPoints(9) = dPoints(0)
            dPoints(10) = dPoints(1)
            dPoints(11) = dCentoEndofTangent1
        
            Dim objBendFrontSurface As IngrGeom3D.Plane3d
            Set objBendFrontSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                    4, dPoints)
        
        ' Set the output
            m_OutputColl.AddOutput "BendFrontSurface", objBendFrontSurface
            Set objBendFrontSurface = Nothing
        
        Case Else
            GoTo ErrorLabel:

    End Select

' Insert your code for output 3(Inclined Tangent)
    oPort2.Set (dCentoEndofTangent1 + dTangentLength2) * Cos(PI / 2 - Angle), _
                0, (dCentoEndofTangent1 + dTangentLength2) * Sin(PI / 2 - Angle)
                
    LineStrPoints(0) = (dCentoEndofTangent1) * Cos(PI / 2 - Angle) - HalfDepth * Cos(Angle)
    LineStrPoints(1) = oPort2.y - HalfWidth
    LineStrPoints(2) = (dCentoEndofTangent1) * Sin(PI / 2 - Angle) + HalfDepth * Sin(Angle)

    LineStrPoints(3) = (dCentoEndofTangent1) * Cos(PI / 2 - Angle) + HalfDepth * Cos(Angle)
    LineStrPoints(4) = oPort2.y - HalfWidth
    LineStrPoints(5) = (dCentoEndofTangent1) * Sin(PI / 2 - Angle) - HalfDepth * Sin(Angle)

    LineStrPoints(6) = (dCentoEndofTangent1) * Cos(PI / 2 - Angle) + HalfDepth * Cos(Angle)
    LineStrPoints(7) = oPort2.y + HalfWidth
    LineStrPoints(8) = (dCentoEndofTangent1) * Sin(PI / 2 - Angle) - HalfDepth * Sin(Angle)

    LineStrPoints(9) = (dCentoEndofTangent1) * Cos(PI / 2 - Angle) - HalfDepth * Cos(Angle)
    LineStrPoints(10) = oPort2.y + HalfWidth
    LineStrPoints(11) = (dCentoEndofTangent1) * Sin(PI / 2 - Angle) + HalfDepth * Sin(Angle)
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
    ProjVector.Set Cos(PI / 2 - Angle), 0, Sin(PI / 2 - Angle)
    
    Dim objHoriTangent As Object
    Set objHoriTangent = PlaceProjection(m_OutputColl, oLineString, ProjVector, dTangentLength2, False)
    'Remove cable tray line string U-shape
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHoriTangent
    Set objHoriTangent = Nothing
    Set ProjVector = Nothing

' Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    oDir.Set 0, 0, -1
    oRadialOrient.Set -1, 0, 0
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
    
    oDir.Set Cos(PI / 2 - Angle), 0, Sin(PI / 2 - Angle)
    oRadialOrient.Set -Cos(Angle), 0, Sin(Angle)
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
