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
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVCS).
'   Physical aspect consist of 2 Cones, 2 Stems, 3 Cylinders, a hand wheel ,a Bottom Cylinder
'   and Two Nozzles. Insulation aspect consist of 2 cylinders for the valve body.
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

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam     As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjStem As Object
    Dim ObjCentralMotor As Object
    Dim ObjMotor1 As Object
    Dim ObjMotor2 As Object
    Dim ObjStem1 As Object
    Dim ObjHandWheel As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl  As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorOffset As Double
    Dim parMotor2Diameter As Double
    Dim parMotor1Diameter As Double
    Dim parActuatorLength1 As Double
    Dim parMotor3Diameter As Double
    Dim parHandWheelLength As Double
    Dim parHandWheelDiameter As Double
    Dim parActuatorOffset1 As Double
    Dim parMotor3EndtoCenter As Double
    Dim parHandWheelOffset As Double
    Dim parHandWheelOffset1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFace2toCenter = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)       'D4
    parActuatorHeight = arrayOfInputs(5)        'D5
    parActuatorLength = arrayOfInputs(6)        'D6
    parActuatorOffset = arrayOfInputs(7)        'D7
    parMotor2Diameter = arrayOfInputs(8)        'D8
    parMotor1Diameter = arrayOfInputs(9)        'D9
    parActuatorLength1 = arrayOfInputs(10)       'D10
    parMotor3Diameter = arrayOfInputs(11)      'D11
    parHandWheelLength = arrayOfInputs(12)            'D12
    parHandWheelDiameter = arrayOfInputs(13)        'D13
    parActuatorOffset1 = arrayOfInputs(14)          'D14
    parMotor3EndtoCenter = arrayOfInputs(15)            'D15
    parHandWheelOffset = arrayOfInputs(16)          'D17
    parHandWheelOffset1 = arrayOfInputs(17)          'D18
    parInsulationThickness = arrayOfInputs(18)
    
    iOutput = 0
    
    Dim stemDia As Double
    Dim handwheelCSDia  As Double
    'Origin of the Symbol is at the point of intersection of Nozzle 1 axis and Stem axis
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
    stPoint.Set 0, -parFace2toCenter + flangeThick, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
'    Insert your code for output 3(Valve Center Cylinder)
    Dim CylDia As Double
    Dim ValCenCylBottLength As Double
'   Valve Center Cylinder Diameter is taken to be 0.7 times the Nozzle2 diameter
    CylDia = 0.7 * pipeDiam
'   Valve Center Cylinder Bottom Length is taken to be 0.25 times the parFace2toCenter
    ValCenCylBottLength = parFace2toCenter * 0.25
    stPoint.Set 0, parCylHeight, 0
    enPoint.Set 0, -ValCenCylBottLength, 0
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
 ' Insert your code for output 4(Stem)
    stPoint.Set 0, parCylHeight, 0
    enPoint.Set 0, parActuatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 5(Central Motor)
    stPoint.Set 0, parActuatorHeight, -parMotor1Diameter / 2
    enPoint.Set 0, parActuatorHeight, parMotor3EndtoCenter - parActuatorOffset1
    Set ObjCentralMotor = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor3Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCentralMotor
    Set ObjCentralMotor = Nothing
    
 ' Insert your code for output 6(Motor2)
    stPoint.Set 0, parActuatorHeight - parActuatorOffset, 0
    enPoint.Set -parActuatorLength, parActuatorHeight - parActuatorOffset, 0
    Set ObjMotor2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor2Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2
    Set ObjMotor2 = Nothing
    
 ' Insert your code for output 7(Motor1)
    stPoint.Set 0, parActuatorHeight, 0
    enPoint.Set parActuatorLength1, parActuatorHeight, 0
    Set ObjMotor1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1
    Set ObjMotor1 = Nothing
    
 ' Insert your code for output 8(Stem1)
    stPoint.Set parHandWheelOffset, parActuatorHeight + parHandWheelOffset1, -parMotor1Diameter / 2
    enPoint.Set parHandWheelOffset, parActuatorHeight + parHandWheelOffset1, -parHandWheelLength
    Set ObjStem1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem1
    Set ObjStem1 = Nothing
    
 ' Insert your code for output 9(HandWheel)
    Dim centerPoint As AutoMath.DPosition
    Dim axisVector As AutoMath.DVector
    Dim majorRadius As Double
    Dim minorRadius As Double
    Set centerPoint = New AutoMath.DPosition
    Set axisVector = New AutoMath.DVector
    
    centerPoint.Set parHandWheelOffset, parActuatorHeight + parHandWheelOffset1, _
                                -parHandWheelLength + handwheelCSDia / 2
    axisVector.Set 0, 0, 1
    majorRadius = (parHandWheelDiameter - handwheelCSDia) / 2
    minorRadius = handwheelCSDia / 2
    Set ObjHandWheel = PlaceTorus(m_OutputColl, centerPoint, axisVector, majorRadius, minorRadius)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
    Set centerPoint = Nothing
    Set axisVector = Nothing
 
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
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, 0
    oDir.Set 0, -1, 0
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
