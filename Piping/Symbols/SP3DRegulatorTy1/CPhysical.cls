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
'   Author:          BG
'   Creation Date:  Tuesday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG1,
'   PDMS Category Name=IVRF). Physical aspect consist of Two cones one converging other divrging
'   and a vertical Cylinder with varying diameter, two nozzles one facing -ive direction direction of
'   X-axis and the other in +ive direction of X-axis.
'   Insulation aspect consist of one Cylinder along X-axis and another cylinder along Y-axis
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
'''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim sptOffset2       As Double
    Dim flangeThick2 As Double
    Dim depth2           As Double
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjConvergCone As Object
    Dim ObjDivergCone As Object
    Dim ObjVertCylins As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFaceToFace = arrayOfInputs(3)        'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight = arrayOfInputs(5)      'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentDiameter = arrayOfInputs(8)     'D8
    parInstrumentDiameter1 = arrayOfInputs(9)        'D9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
        
' Insert your code for output 1(Converging Cone)
    Dim centerTopPt As AutoMath.DPosition
    Dim centerBasePt As AutoMath.DPosition
    Set centerTopPt = New AutoMath.DPosition
    Set centerBasePt = New AutoMath.DPosition
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset1, depth1
'Origin of the symbol is assumed at the intersection of vertical cylinder axis and nozzles  axes
    centerBasePt.Set -parFace1toCenter + flangeThick1, 0, 0
    centerTopPt.Set 0, 0, 0
    Set ObjConvergCone = PlaceCone(m_OutputColl, centerBasePt, centerTopPt, pipeDiam1 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergCone
    Set ObjConvergCone = Nothing
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2
 ' Insert your code for output 2(Diverging Cone)
    centerBasePt.Set parFaceToFace - parFace1toCenter - flangeThick2, 0, 0
    centerTopPt.Set 0, 0, 0
    Set ObjDivergCone = PlaceCone(m_OutputColl, centerBasePt, centerTopPt, pipeDiam2 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergCone
    Set ObjDivergCone = Nothing
    Set centerTopPt = Nothing
    Set centerBasePt = Nothing
    
' Insert your code for output 3(Vertical Cylinders)
    Dim CylDia As Double
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 29) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector
    'Bottom Cylinder Diameter is taken as equal to Pipe 1 Diameter as per EDEN
    CylDia = pipeDiam1
    'Point1
    dblLinePts(0) = 0
    dblLinePts(1) = -parCylHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = CylDia / 2
    dblLinePts(4) = -parCylHeight
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = CylDia / 2
    dblLinePts(7) = 0
    dblLinePts(8) = 0
    'Point4
    'Cylinder, immediately above the cone, diameter is assumed as 0.5*times the parInstrumentDiameter value
    dblLinePts(9) = parInstrumentDiameter / 4
    dblLinePts(10) = 0
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = parInstrumentDiameter / 4
    dblLinePts(13) = parInstrumentHeight - parInstrumentHeight2 - parInstrumentHeight1    '(D5-D7-D6)
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = parInstrumentDiameter1 / 2
    dblLinePts(16) = parInstrumentHeight - parInstrumentHeight2 - parInstrumentHeight1   '(D5-D7-D6)
    dblLinePts(17) = 0
    'Point7
    dblLinePts(18) = parInstrumentDiameter1 / 2
    dblLinePts(19) = parInstrumentHeight - parInstrumentHeight1                                 '(D5-D6)
    dblLinePts(20) = 0
    'Point8
    dblLinePts(21) = parInstrumentDiameter / 2
    dblLinePts(22) = parInstrumentHeight - parInstrumentHeight1                                '(D5-D6)
    dblLinePts(23) = 0
    'Point9
    dblLinePts(24) = parInstrumentDiameter / 2
    dblLinePts(25) = parInstrumentHeight
    dblLinePts(26) = 0
    'Point10
    dblLinePts(27) = 0
    dblLinePts(28) = parInstrumentHeight
    dblLinePts(29) = 0
    
    Set objLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblLinePts)
    revVector.Set 0, 1, 0
    centerPt.Set 0, 0, 0
    Set ObjVertCylins = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylins
    Set ObjVertCylins = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    tmpLineStr.Remove
    Set objLineStr = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1 for Output 4 (Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 5 (Nozzle2)
    oPlacePoint.Set parFaceToFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
