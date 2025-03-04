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
'   Author:          svsmylav
'   Creation Date:  Friday, Mar 1 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is prepared using six outputs. Output 1) ObjHoriTangent & 3)ObjInclTangent are created using
' PlaceProjection'function. Output 2 is ObjElbow created by using 'PlaceRevolution' function.
' Outputs 4 & 5 are Cable Tray Ports created by using 'CreateCableTrayPort' function.
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
'  11.June.2007     KKC                DM-121367  Implemented part data basis for the Square Outside Corners
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
    Dim ObjProj As IJDObject
    Dim ObjProj1 As IJDObject
    Dim ObjInclTangent  As Object
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
    Dim parBendRadius As Double
    Dim parTangentLength As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim oPort2 As New AutoMath.DPosition 'Port 2 center point
    Dim ProjVector As New AutoMath.DVector
    Dim LineStrCP As New AutoMath.DPosition
    Dim lCableTrayPartDataBasis As Long
    Dim parCTDimension As Double
' Inputs
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
    
    If lCableTrayPartDataBasis = 5 Then
         parCTDimension = arrayOfInputs(2)
    End If
    iOutput = 0
    parBendRadius = oTrayPart.BendRadius
    parTangentLength = oTrayPart.TangentLength
    'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(parTangentLength, 0) Then parTangentLength = 0.000001
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)

' Insert your code for output 1(U Shape Horizontal Tangent)
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim HD              As Double
    Dim HW              As Double
    Dim LineStrPoints(0 To 11)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim parAngle As Double
    parAngle = PI / 4
    HD = parActualDepth / 2
    HW = parActualWidth / 2
    If lCableTrayPartDataBasis <= 1 Then
        Dim parFacetoCenter As Double
        parFacetoCenter = (parBendRadius + HW) * Tan(parAngle / 2) + parTangentLength
        oPort1.Set -parFacetoCenter, 0, 0
    ElseIf lCableTrayPartDataBasis = 5 Then
        Dim dFacetoCenter As Double
        dFacetoCenter = parCTDimension - parActualWidth / (2 * Tan(PI / 2.667)) ' Angle 67.5 (135/2) is represented as PI/2.66
        oPort1.Set -dFacetoCenter, 0, 0
    End If
    
    
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
    If lCableTrayPartDataBasis <= 1 Then
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
        Set oElbowLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPoints)
        ProjVector.Set 0, 0, 1
        CP.Set -parFacetoCenter + parTangentLength, parBendRadius + HW, 0
        Set ObjElbow = PlaceRevolution(m_OutputColl, oElbowLineStr, ProjVector, CP, parAngle, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
        Set ObjElbow = Nothing
        '   Remove cable tray line string U-shape
        Set objLineString = oElbowLineStr
        Set oElbowLineStr = Nothing
        objLineString.Remove
        
    ElseIf lCableTrayPartDataBasis = 5 Then
    ' Insert your code for output 2(U Shape Plane apart form tangents)
        Dim oProjLineStr As IngrGeom3D.LineString3d
        LineStrCP.Set -dFacetoCenter + parTangentLength, 0, -HD
        Dim LineStrPoints1(0 To 17) As Double
        
        LineStrPoints1(0) = LineStrCP.x
        LineStrPoints1(1) = LineStrCP.y + HW
        LineStrPoints1(2) = LineStrCP.z
        
        LineStrPoints1(3) = LineStrCP.x
        LineStrPoints1(4) = LineStrCP.y - HW
        LineStrPoints1(5) = LineStrCP.z
        
        LineStrPoints1(6) = HW / Tan(PI / 2.66)
        LineStrPoints1(7) = -HW
        LineStrPoints1(8) = LineStrCP.z
        
        LineStrPoints1(9) = dFacetoCenter * Cos(parAngle) + HW * Sin(parAngle) - parTangentLength * Sin(parAngle)
        LineStrPoints1(10) = dFacetoCenter * Cos(parAngle) - HW * Cos(parAngle) - parTangentLength * Cos(parAngle)
        LineStrPoints1(11) = LineStrCP.z
        
        LineStrPoints1(12) = dFacetoCenter * Cos(parAngle) - HW * Sin(parAngle) - parTangentLength * Sin(parAngle)
        LineStrPoints1(13) = dFacetoCenter * Cos(parAngle) + HW * Cos(parAngle) - parTangentLength * Sin(parAngle)
        LineStrPoints1(14) = LineStrCP.z
        
        LineStrPoints1(15) = -HW / Tan(PI / 2.66)
        LineStrPoints1(16) = HW
        LineStrPoints1(17) = LineStrCP.z
        
        Dim oHLine As IngrGeom3D.Line3d
        Set oHLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5), LineStrPoints1(6), LineStrPoints1(7), LineStrPoints1(8))
        
        Dim oILine1 As IngrGeom3D.Line3d
        Set oILine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(6), LineStrPoints1(7), LineStrPoints1(8), LineStrPoints1(9), LineStrPoints1(10), LineStrPoints1(11))
        
        Dim oILine2 As IngrGeom3D.Line3d
        Set oILine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(9), LineStrPoints1(10), LineStrPoints1(11), LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14))
        
        Dim oILine3  As IngrGeom3D.Line3d
        Set oILine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14), LineStrPoints1(15), LineStrPoints1(16), LineStrPoints1(17))
        
        Dim oHLine1  As IngrGeom3D.Line3d
        Set oHLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(15), LineStrPoints1(16), LineStrPoints1(17), LineStrPoints1(0), LineStrPoints1(1), LineStrPoints1(2))
        
        Dim oVLine   As IngrGeom3D.Line3d
        Set oVLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, LineStrPoints1(0), LineStrPoints1(1), LineStrPoints1(2), LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5))
        
        Dim oLiCollection           As Collection
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
        
        Set ObjProj = PlaceProjection(m_OutputColl, objBp1scurve, ProjVector, parActualDepth, True)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjProj
        Set ObjProj = Nothing
        Set oLiCollection = Nothing
        
        Dim oLiCollection1 As Collection
        Set oLiCollection1 = New Collection
        oLiCollection1.Add oILine3
        oLiCollection1.Add oHLine1
        Set oStPoint = New AutoMath.DPosition
        
        oStPoint.Set LineStrPoints1(12), LineStrPoints1(13), LineStrPoints1(14)
        Set objBp1scurve = New IngrGeom3D.ComplexString3d
        Set objBp1scurve = PlaceTrCString(oStPoint, oLiCollection1)
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
        obranchcurves.Add oHLine1
        obranchcurves.Add oVLine
    
        Dim objHBcurves         As IngrGeom3D.ComplexString3d
        oStPoint.Set LineStrPoints1(3), LineStrPoints1(4), LineStrPoints1(5)
        Set objHBcurves = PlaceTrCString(oStPoint, obranchcurves)
        
        Dim TrayBottom As IngrGeom3D.Plane3d
        Set TrayBottom = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -dFacetoCenter, 0, -HD, ProjVector.x, ProjVector.y, ProjVector.z)
        Call TrayBottom.AddBoundary(objHBcurves)
    ' Set the output
        m_OutputColl.AddOutput "Proj_", TrayBottom
        Set TrayBottom = Nothing
        Set objHBcurves = Nothing
        Set obranchcurves = Nothing
    
   End If
    
' Insert your code for output 3(U Shape Inclined Tangent)
    Dim oInclinedLineStr As IngrGeom3D.LineString3d
    If lCableTrayPartDataBasis <= 1 Then
        oPort2.Set parFacetoCenter * Cos(parAngle), parFacetoCenter * Sin(parAngle), 0
    ElseIf lCableTrayPartDataBasis = 5 Then
        oPort2.Set dFacetoCenter * Cos(parAngle), dFacetoCenter * Sin(parAngle), 0
    End If
    
    
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
