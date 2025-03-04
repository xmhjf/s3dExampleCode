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
'   Creation Date:  Wednesday, Jul 18 2006
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     Temperature Switch (Source: Nason, Temperature Switch Model TD, http://www.nasonptc.com/pdfs/nasoncat_swi2.pdf)
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

    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight1 = arrayOfInputs(2)
    parInstrumentHeight2 = arrayOfInputs(3)
    parInstrumentDiameter = arrayOfInputs(4)
    parInstrumentDiameter1 = arrayOfInputs(5)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0

' Insert your code for output 1(Bottom Cylinder)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
        
    stPoint.Set 0, 0, 0
    enPoint.Set parInstrumentHeight1, 0, 0
                                        
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parInstrumentDiameter1

' Output 2 (The heaxagonal nut at the port)
    Dim ObjNut   As Object
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim iCount   As Integer
    Dim dAngle   As Double
    Dim dNutHeight As Double
    Dim dRad As Double
    Dim dThreadedHeight As Double
    Dim k As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    
    k = 0.85    ' Assumption : Nut Height = K * Pipe Dia
    dNutHeight = k * pipeDiam
    dRad = pipeDiam 'Assumption - Nominal Dia of nut is equal to pipedia"
    dThreadedHeight = 2 * dNutHeight  'Assumption - Height of Threaded portion is 2 * dNutHeight
    
    Dim axisVect   As AutoMath.DVector
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
        
    dAngle = PI / 3
    stPoint.Set parInstrumentHeight1 + dThreadedHeight, 0, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x
        LineStrPoints(3 * iCount - 2) = stPoint.y + (dRad) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 1) = stPoint.z + (dRad) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 1, 0, 0

    Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)

' Set the output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
    Set ObjNut = Nothing
    Set oLineString = Nothing
    Set geomFactory = Nothing
    Set axisVect = Nothing

' Insert your code for output 3(Top Cylinder)
    
    stPoint.Set parInstrumentHeight1 + dThreadedHeight + dNutHeight, 0, 0
    enPoint.Set parInstrumentHeight1 + dThreadedHeight + parInstrumentHeight2, 0, 0
                                        
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parInstrumentDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1 (The Port)
    Dim dNozzleLength As Double
    Dim objNozzle  As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set parInstrumentHeight1 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    dNozzleLength = dThreadedHeight
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


