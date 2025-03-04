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
'   Creation Date:  Thursday, Oct 3 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IEAM5,
'   PDMS Category Name=IVCQ). Physical aspect consist of a Stem, 5 Cylinders, 1 hand wheel, 2 cones
'   and 2 Nozzles. Insulation aspect consist of a single cylinder for the valve body.
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
    
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjStemAndCylin As Object
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjCylinder As Object
    Dim ObjMotor1 As Object
    Dim ObjMotor2 As Object
    Dim ObjStem1  As Object
    Dim ObjHandWheel As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parHandWheelOffset As Double
    Dim ActuatorDiameter As Double
    Dim parActuatorDiameter1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parMotor1Diameter As Double
    Dim parMotor1Length As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Diameter As Double
    Dim parHandWheelOffset1 As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelLength As Double
    Dim parMotor2Offset As Double
    Dim parMotor2Length As Double
    Dim parInsulationThickness As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)       'D4
    parActuatorHeight = arrayOfInputs(5)        'D5
    parActuatorHeight1 = arrayOfInputs(6)       'D6
    parActuatorHeight2 = arrayOfInputs(7)        'D7
    parHandWheelOffset = arrayOfInputs(8)       'D8
    parActuatorDiameter = arrayOfInputs(9)      'D9
    parActuatorLength = arrayOfInputs(10)            'D10
    parActuatorDiameter1 = arrayOfInputs(11)            'D11
    parMotor1Diameter = arrayOfInputs(12)            'D12
    parMotor1Length = arrayOfInputs(13)            'D13
    parMotor1Offset = arrayOfInputs(14)            'D14
    parMotor2Diameter = arrayOfInputs(15)            'D15
    parHandWheelOffset1 = arrayOfInputs(16)             'D16
    parHandWheelDiameter = arrayOfInputs(17)            'D17
    parHandWheelLength = arrayOfInputs(18)            'D18
    parMotor2Offset = arrayOfInputs(19)            'D19
    parMotor2Length = arrayOfInputs(20)            'D20
    parInsulationThickness = arrayOfInputs(21)
    
    iOutput = 0
    Dim stemDia As Double
    'variable handwheelCSDia refers to Hand wheel cross section diameter
    Dim handwheelCSDia As Double
    'Origin of the Symbol is at the point of intersection of Nozzle 1 axis and Nozzle 2 axis
    'Stem Diameter and hand wheel Diameter are assumed as 0.05
    stemDia = 0.05
    handwheelCSDia = 0.05
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
  ' Insert your code for output 1(Body Cone1 of Valve)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 2(Body Cone2 of Valve)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
 ' Insert your code for output 4(Stem and Cylinder)
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 17) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector
    'Point1
    dblLinePts(0) = 0
    dblLinePts(1) = 0
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = stemDia / 2
    dblLinePts(4) = 0
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = stemDia / 2
    dblLinePts(7) = parActuatorHeight1
    dblLinePts(8) = 0
    'Point4
    dblLinePts(9) = parActuatorDiameter / 2
    dblLinePts(10) = parActuatorHeight1
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = parActuatorDiameter / 2
    dblLinePts(13) = parActuatorHeight
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = 0
    dblLinePts(16) = parActuatorHeight
    dblLinePts(17) = 0
    
    Set objLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dblLinePts)
    revVector.Set 0, 1, 0
    centerPt.Set 0, 0, 0
    Set ObjStemAndCylin = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemAndCylin
    Set ObjStemAndCylin = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    tmpLineStr.Remove
    Set objLineStr = Nothing
    
 ' Insert your code for output 5(Cylinder)
    stPoint.Set 0, parActuatorHeight1 + parActuatorHeight2 + parHandWheelOffset, 0
    enPoint.Set -parActuatorLength, stPoint.y, 0
    Set ObjCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    
 ' Insert your code for output 6(Motor 1)
    stPoint.Set -parMotor1Offset, parActuatorHeight1 + parActuatorHeight2 + parHandWheelOffset, 0
    enPoint.Set -parMotor1Offset, stPoint.y, parMotor1Length
    Set ObjMotor1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1
    Set ObjMotor1 = Nothing
    
 ' Insert your code for output 7(Motor 2)
    stPoint.Set parMotor2Offset, parActuatorHeight1 + parActuatorHeight2 + parHandWheelOffset, 0
    enPoint.Set parMotor2Offset, stPoint.y, parMotor2Length
    Set ObjMotor2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor2Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2
    Set ObjMotor2 = Nothing
    
 ' Insert your code for output 8(Stem1)
    stPoint.Set 0, parActuatorHeight1 + parActuatorHeight2, parHandWheelOffset1
    enPoint.Set parHandWheelLength, parActuatorHeight1 + parActuatorHeight2, parHandWheelOffset1
    Set ObjStem1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem1
    Set ObjStem1 = Nothing
    Set geomFactory = Nothing
    
' Insert your code for output 9(Hand Wheel)
    Dim majorRadius As Double
    Dim minorRadius As Double
    majorRadius = (parHandWheelDiameter - handwheelCSDia) / 2
    minorRadius = handwheelCSDia / 2
    revVector.Set 1, 0, 0
    centerPt.Set parHandWheelLength - handwheelCSDia / 2, parActuatorHeight1 + parActuatorHeight2, _
                        parHandWheelOffset1
    Set ObjHandWheel = PlaceTorus(m_OutputColl, centerPt, revVector, majorRadius, minorRadius)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    
' Place Nozzle 1 for Output 10(Nozzle1)
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

' Place Nozzle 2 for output 11 (Nozzle2)
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
