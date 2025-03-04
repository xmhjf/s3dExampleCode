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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav 
'   Creation Date:  Friday, Sep 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIV,
'   PDMS Category Name=IEEF). Physical aspect consist of four nozzles and one revolved linestring.
'   Insulation aspect consist of single entity which is a cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object

    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentDiameter2 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 29) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)            'D2 is equal to D3 + D4 + D5, so NOT used.
    parInstrumentLength = arrayOfInputs(3)      'D3
    parInstrumentLength1 = arrayOfInputs(4)     'D4
    parInstrumentLength2 = arrayOfInputs(5)     'D5
    parInstrumentDiameter = arrayOfInputs(6)    'D6
    parInstrumentDiameter1 = arrayOfInputs(7)   'D7
    parInstrumentDiameter2 = arrayOfInputs(8)   'D8
    parPortRotation1 = arrayOfInputs(9)                 'D9
    parPortRotation2 = arrayOfInputs(10)               'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0
    
'   Origin of the symbol is assumed at the intersection of axis joining port1 to port2,
'   and axis joining port3 to port4.

' Insert your code for output 2(Body of Venturi)
'   Point 1
    LineStringPoints(0) = -parInstrumentLength1 / 2 - parInstrumentLength
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = -parInstrumentLength1 / 2 - parInstrumentLength
    LineStringPoints(4) = 0
    LineStringPoints(5) = parInstrumentDiameter / 2
    
'   Assumption: Ratio of Coverging cone length to straight portion length on Port1 side is
'   0.7:0.3, diverging cone to straight portion length on Port1 side Port2 is the same.

'   Point 3
    LineStringPoints(6) = -parInstrumentLength1 / 2 - 0.3 * parInstrumentLength
    LineStringPoints(7) = 0
    LineStringPoints(8) = parInstrumentDiameter1 / 2
    
'   Point 4
    LineStringPoints(9) = -parInstrumentLength1 / 2
    LineStringPoints(10) = 0
    LineStringPoints(11) = parInstrumentDiameter1 / 2
    
'   Point 5
    LineStringPoints(12) = -parInstrumentLength1 / 2
    LineStringPoints(13) = 0
    LineStringPoints(14) = parInstrumentDiameter2 / 2
    
'   Point 6
    LineStringPoints(15) = parInstrumentLength1 / 2
    LineStringPoints(16) = 0
    LineStringPoints(17) = parInstrumentDiameter2 / 2
    
'   Point 7
    LineStringPoints(18) = parInstrumentLength1 / 2
    LineStringPoints(19) = 0
    LineStringPoints(20) = parInstrumentDiameter1 / 2
    
'   Point 8
    LineStringPoints(21) = parInstrumentLength1 / 2 + 0.3 * parInstrumentLength2
    LineStringPoints(22) = 0
    LineStringPoints(23) = parInstrumentDiameter1 / 2
    
'   Point 9
    LineStringPoints(24) = parInstrumentLength1 / 2 + parInstrumentLength2
    LineStringPoints(25) = 0
    LineStringPoints(26) = parInstrumentDiameter / 2
    
'   Point 10
    LineStringPoints(27) = parInstrumentLength1 / 2 + parInstrumentLength2
    LineStringPoints(28) = 0
    LineStringPoints(29) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, 0, 0
    Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
      
' Place Nozzle 1
'   Nozzle1 and Nozzle2  are considered at the base of cones
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parInstrumentLength1 / 2 - parInstrumentLength - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parInstrumentLength1 / 2 + parInstrumentLength2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim radnozzPlaceDist As Double
    radnozzPlaceDist = Sqr(Abs((parInstrumentDiameter2 / 2) ^ 2 - (pipeDiam / 2 * pipeDiam / 2)))
    oPlacePoint.Set 0, _
                    (radnozzPlaceDist + sptOffset - depth) * Cos(parPortRotation1), _
                    (radnozzPlaceDist + sptOffset - depth) * Sin(parPortRotation1)
    oDir.Set 0, Cos(parPortRotation1), Sin(parPortRotation1)
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    radnozzPlaceDist = Sqr(Abs((parInstrumentDiameter2 / 2) ^ 2 - (pipeDiam / 2 * pipeDiam / 2)))
    oPlacePoint.Set 0, _
                    (radnozzPlaceDist + sptOffset - depth) * Cos(parPortRotation2), _
                    (radnozzPlaceDist + sptOffset - depth) * Sin(parPortRotation2)
    oDir.Set 0, Cos(parPortRotation2), Sin(parPortRotation2)
    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
 
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
 
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
