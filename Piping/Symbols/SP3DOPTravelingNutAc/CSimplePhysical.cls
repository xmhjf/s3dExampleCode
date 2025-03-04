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
'   Creation Date:  Wednesday, Aug 1 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     Traveling Nut Actuator(Source: Val-Matic Valve & Mfg. Corp., American BFY,Class 150B & 250BMark at www.valmatic.com/butterflyvalves.jsp)
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

    Dim parOperatorHeight As Double
    Dim parOperatorDepth As Double
                
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2) ' G
    parOperatorDepth = arrayOfInputs(3)  ' K
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
'    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
   Dim cenPoint As AutoMath.DPosition
   Set cenPoint = New AutoMath.DPosition
   cenPoint.Set 0, parOperatorHeight, 0
'The Instrument Box
    Dim objInstBox As Object
    Dim LineStrPoints(0 To 20) As Double
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim axisVect As AutoMath.DVector
            
' Assumptions :
' The Instrument Box is a half octagon projected along the Y axis to dWidth.
'  1. The Width of the Instrument Box (dWidth) is .25 * parOperatorHeight
'  2. The Depth of the Instrument Box (dDepth) is 1.2 * parOperatorDepth
'  Calculations : dSide(side of the octagon) = dDepth / (1 + sqr(2))
    
    Dim dWidth As Double
    Dim dDepth As Double
    Dim dSide As Double
        
    dWidth = 0.25 * parOperatorHeight
    dDepth = 1.2 * parOperatorDepth
    dSide = dDepth / (1 + Sqr(2))
       
'1
    LineStrPoints(0) = cenPoint.x - dSide / 2
    LineStrPoints(1) = cenPoint.y - dWidth / 2
    LineStrPoints(2) = cenPoint.z + dDepth / 2
'2
    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = cenPoint.z - dDepth / 2
'3
    LineStrPoints(6) = cenPoint.x
    LineStrPoints(7) = LineStrPoints(1)
    LineStrPoints(8) = LineStrPoints(5)
'4
    LineStrPoints(9) = cenPoint.x + dDepth / 2 - dSide / 2
    LineStrPoints(10) = LineStrPoints(1)
    LineStrPoints(11) = cenPoint.z - dSide / 2
'5
    LineStrPoints(12) = LineStrPoints(9)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = cenPoint.z + dSide / 2
'6
    LineStrPoints(15) = cenPoint.x
    LineStrPoints(16) = LineStrPoints(1)
    LineStrPoints(17) = LineStrPoints(2)
'7
    LineStrPoints(18) = LineStrPoints(0)
    LineStrPoints(19) = LineStrPoints(1)
    LineStrPoints(20) = LineStrPoints(2)
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 0, 1, 0
    Set objInstBox = PlaceProjection(m_OutputColl, oLineString, axisVect, dWidth, True)
    
'   Set the outputoLineString
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInstBox
    Set objInstBox = Nothing
    
' Insert your code for output 2(Cylinder1)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim dCylLength1 As Double
    dCylLength1 = 0.1 * parOperatorDepth ' Assumption: Length of cyl is 0.1 * parOperatorDepth
    stPoint.Set cenPoint.x - dSide / 4, cenPoint.y, cenPoint.z - dDepth / 2
    enPoint.Set cenPoint.x - dSide / 4, cenPoint.y, cenPoint.z - dDepth / 2 - dCylLength1
 
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 0.75 * dSide / 2
    
' Insert your code for output 3(Cylinder2)
    Dim dCylLength2 As Double
    dCylLength2 = 0.14 * parOperatorDepth ' Assumption: Length of cyl is 0.14 * parOperatorDepth
    stPoint.Set cenPoint.x - dSide / 4, cenPoint.y, cenPoint.z - dDepth / 2 - dCylLength1
    enPoint.Set cenPoint.x - dSide / 4, cenPoint.y, cenPoint.z - dDepth / 2 - dCylLength1 - dCylLength2
 
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 0.25 * dSide / 2
   
' Insert your code for output 4(Square Nut at the top)
    Dim ObjSqNut   As Object
    Dim LineStrPoints1(0 To 14)  As Double
    Dim iCount As Integer
    Dim dAngle As Double
    Dim dRad As Double
    Dim dNutHeight As Double
    dRad = 0.0508 / (2 * Sqr(2)) 'Assumption : Side of square is 2"(0.0508m), dRad(half of the diagonal) = Side/(2*Sqr(2))
    dNutHeight = dRad
    dAngle = PI / 2
    stPoint.Set cenPoint.x - dSide / 4, cenPoint.y, cenPoint.z - dDepth / 2 - dCylLength1 - dCylLength2
'   First point is considered on at an angle (PI/4) from positive Y-axis and positive X-axis
    For iCount = 1 To 5
        LineStrPoints1(3 * iCount - 3) = stPoint.x + (dRad) * Cos(PI / 4 + iCount * dAngle)
        LineStrPoints1(3 * iCount - 2) = stPoint.y + (dRad) * Sin(PI / 4 + iCount * dAngle)
        LineStrPoints1(3 * iCount - 1) = stPoint.z
    Next iCount
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints1)
    axisVect.Set 0, 0, -1

    Set ObjSqNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)

' Set the output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSqNut
    Set ObjSqNut = Nothing
    Set oLineString = Nothing
    Set geomFactory = Nothing
        
' Insert your code for output 5(Stem)
    stPoint.Set cenPoint.x, cenPoint.y - parOperatorHeight, cenPoint.z
    enPoint.Set cenPoint.x, cenPoint.y - dWidth / 2, cenPoint.z
 
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 0.0508 ' Assumption: Dia is 2"
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


