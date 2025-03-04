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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         dkl
'   Creation Date:  Wednesday, Jul 19 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Pressure Regulator (Source: Jordan Valve, Mark 60 Self-Operated Pressure Regulators, www.jordanvalve.com/literature/dimdrawing_mk50_60.pdf)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI       As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
  
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_terminate()
    Set m_oGeomHelper = Nothing
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

    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
        
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2) ' A
    parInstrumentHeight = arrayOfInputs(3)  ' B
    parInstrumentHeight1 = arrayOfInputs(4)  ' C
    parInstrumentDiameter = arrayOfInputs(5)  ' D

    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
   
'The Bell
    Dim objBell As Object
    Dim LineStrPoints(0 To 32) As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim axisVect As AutoMath.DVector
    Dim cenPoint As AutoMath.DPosition
        
' Assumptions :
'  1. The height of the nut at the top (dNutHeight) is .04 * parInstrumentHeight
'  2. The height of the flanged portion of the Bell (dFlangeHeight) is .09 * parInstrumentHeight
'  3. The height of the middle portion of the Bell (dMiddleHeight) is .45 * parInstrumentHeight
'  4. The height of the Top portion of the Bell (dTopHeight) is .12 * parInstrumentHeight
'  5. The Length of the neck of the Bell (dNeckLength) is .09 * parInstrumentHeight
'  6. The Length of the Middle portion of the Bell (dMiddleLength) is .35 * parInstrumentHeight
'  7. The Length of the Top portion of the Bell (dTopLength) is .07 * parInstrumentHeight
    
    Dim dNeckHeight As Double
    Dim dFlangeHeight As Double
    Dim dMiddleHeight As Double
    Dim dTopHeight As Double
    Dim dNeckLength As Double
    Dim dMiddleLength As Double
    Dim dTopLength As Double
    Dim dNutHeight As Double
    
    dNutHeight = 0.04 * parInstrumentHeight
    dFlangeHeight = 0.09 * parInstrumentHeight
    dMiddleHeight = 0.45 * parInstrumentHeight
    dTopHeight = 0.12 * parInstrumentHeight
    dNeckLength = 0.09 * parInstrumentHeight
    dMiddleLength = 0.35 * parInstrumentHeight
    dTopLength = 0.07 * parInstrumentHeight
    dNeckHeight = parInstrumentHeight - dNutHeight - dFlangeHeight - dMiddleHeight - dTopHeight
    If CmpDblLessThan(dNeckHeight, flangeDiam) Then
    Const Factor = 0.00254
    dNeckHeight = flangeDiam + Factor
    End If
'1
    LineStrPoints(0) = -dNeckLength / 2
    LineStrPoints(1) = 0
    LineStrPoints(2) = 0
'2
    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = dNeckHeight
    LineStrPoints(5) = LineStrPoints(2)
'3
    LineStrPoints(6) = -parInstrumentDiameter / 2
    LineStrPoints(7) = LineStrPoints(4)
    LineStrPoints(8) = LineStrPoints(2)
'4
    LineStrPoints(9) = LineStrPoints(6)
    LineStrPoints(10) = dNeckHeight + dFlangeHeight
    LineStrPoints(11) = LineStrPoints(2)
'5
    LineStrPoints(12) = -dMiddleLength / 2
    LineStrPoints(13) = LineStrPoints(10)
    LineStrPoints(14) = LineStrPoints(2)
'6
    LineStrPoints(15) = LineStrPoints(12)
    LineStrPoints(16) = LineStrPoints(10) + dMiddleHeight
    LineStrPoints(17) = LineStrPoints(2)
'7
    LineStrPoints(18) = -dTopLength / 2
    LineStrPoints(19) = LineStrPoints(16)
    LineStrPoints(20) = LineStrPoints(2)
'8
    LineStrPoints(21) = LineStrPoints(18)
    LineStrPoints(22) = LineStrPoints(16) + dTopHeight
    LineStrPoints(23) = LineStrPoints(2)
'9
    LineStrPoints(24) = 0
    LineStrPoints(25) = LineStrPoints(22)
    LineStrPoints(26) = LineStrPoints(2)
'10
    LineStrPoints(27) = 0
    LineStrPoints(28) = 0
    LineStrPoints(29) = LineStrPoints(2)
'11
    LineStrPoints(30) = LineStrPoints(0)
    LineStrPoints(31) = LineStrPoints(1)
    LineStrPoints(32) = LineStrPoints(2)
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 11, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 0, 1, 0
    Set cenPoint = New AutoMath.DPosition
    cenPoint.Set 0.00001, 0, 0
    Set objBell = PlaceRevolution(m_OutputColl, oLineString, axisVect, cenPoint, _
                                2 * PI, True)

'   Set the outputoLineString
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBell
    Set objBell = Nothing
    Set cenPoint = Nothing

' Insert your code for output 2(Nut)
    Dim ObjNut   As Object
    Dim LineStrPoints1(0 To 20)  As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim iCount As Integer
    Dim dAngle As Double
    Dim dRad As Double
    dRad = 1.2 * dTopLength / 2 ' Assumption : drad = 1.2 * dTopLength / 2
    dAngle = PI / 3
    stPoint.Set 0, dNeckHeight + dFlangeHeight + dMiddleHeight + dTopHeight, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints1(3 * iCount - 3) = stPoint.x + (dRad) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints1(3 * iCount - 2) = stPoint.y
        LineStrPoints1(3 * iCount - 1) = stPoint.z + (dRad) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints1)
    axisVect.Set 0, 1, 0

    Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)

' Set the output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
    Set ObjNut = Nothing
    Set oLineString = Nothing
    Set geomFactory = Nothing

' Insert your code for output 3(Bottom Cylinder)
    Dim dCylLength As Double
    dCylLength = 0.09 * parFacetoFace ' Assumption: Length of cyl is 0.09 * parFacetoFace
    stPoint.Set -dCylLength, 0, 0 ' The cylinder is not placed symetrically along the bell portion.
    enPoint.Set 0, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * parInstrumentHeight1
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim Port1Length As Double
    Port1Length = parFacetoFace - flangeThick
    If CmpDblLessThan(Port1Length, flangeThick) Then Port1Length = flangeThick
    oDir.Set -1, 0, 0
    Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing

' Place Nozzle 2
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
 
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


