VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          BG
'   Creation Date:  Friday, September 19 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IEAM2,
'   PDMS Category Name=IVCN). Physical aspect consist of Valve Body using 2 cones, 2 Stems,
'   3 Cylinders,a hand wheel ,a Bottom Cylinder and Two Nozzles. Insulation aspect consist of insulation for
'   the valve body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Operation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
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
    
    'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parPortIndex2 As Long
    Dim parNpd2 As Double
    Dim parEndPreparation2 As Long
    Dim parScheduleThickness2 As Long
    Dim parEndStandard2 As Long
    Dim parPressureRating2 As Long
    Dim parFlowDirection2 As DistribFlow
    Dim parId2 As String
    Dim parNPDUnitType2 As String

    Dim parOperationHeight  As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
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
    parNPD = arrayOfInputs(19)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(20)
    parScheduleThickness = arrayOfInputs(21)
    parEndStandard = arrayOfInputs(22)
    parPressureRating = arrayOfInputs(23)
    parFlowDirection = arrayOfInputs(24)
    parPortIndex1 = arrayOfInputs(25)
    parNpd1 = arrayOfInputs(26)
    parEndPreparation1 = arrayOfInputs(27)
    parScheduleThickness1 = arrayOfInputs(28)
    parEndStandard1 = arrayOfInputs(29)
    parPressureRating1 = arrayOfInputs(30)
    parFlowDirection1 = arrayOfInputs(31)
    parPortIndex2 = arrayOfInputs(32)
    parNpd2 = arrayOfInputs(33)
    parEndPreparation2 = arrayOfInputs(34)
    parScheduleThickness2 = arrayOfInputs(35)
    parEndStandard2 = arrayOfInputs(36)
    parPressureRating2 = arrayOfInputs(37)
    parFlowDirection2 = arrayOfInputs(38)
    parOperationHeight = arrayOfInputs(39)
    
    parId1 = arrayOfInputs(40)
    parId2 = arrayOfInputs(41)
    parNPDUnitType = arrayOfInputs(42)
    parNPDUnitType1 = arrayOfInputs(43)
    parNPDUnitType2 = arrayOfInputs(44)
    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output (Relief Valve Operation aspect)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    Dim ObjOpEnvelope As Object
    stPoint.Set -parActuatorLength, parActuatorHeight + parMotor3Diameter / 2, -parMotor1Diameter / 2
    enPoint.Set parActuatorLength1, stPoint.y + parOperationHeight, parMotor3EndtoCenter - parActuatorOffset1
    Set ObjOpEnvelope = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpEnvelope
    Set ObjOpEnvelope = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


