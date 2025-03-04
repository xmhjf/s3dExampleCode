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
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IEAM3,
'   PDMS Category Name=IVCO). The symbol consists of Thirteen outputs, Twelve physical outputs
'   and One Insulation Output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
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
    Dim parCylHeight  As Double
    Dim parMotor4Offset As Double
    Dim parMotor4Length As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parHandWheelDiameter As Double
    Dim parMotor1Diameter As Double
    Dim parMotor4Diameter As Double
    Dim parMotor3Length As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Length As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parMotor3toCenter As Double
    Dim parActuatorHeight3 As Double
    Dim parInsulationThickness As Double
    
'   Dynamic Port Input parameters variable declaration
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

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'D2
    parFacetoFace = arrayOfInputs(3)     'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parMotor4Offset = arrayOfInputs(5)   'D5
    parMotor4Length = arrayOfInputs(6)   'D6
    parActuatorHeight = arrayOfInputs(7)     'D7
    parActuatorHeight1 = arrayOfInputs(8)    'D8
    parActuatorHeight2 = arrayOfInputs(9)     'D9
    parHandWheelDiameter = arrayOfInputs(10)     'D10
    parMotor1Diameter = arrayOfInputs(11)    'D11
    parMotor4Diameter = arrayOfInputs(12)    'D12
    parMotor3Length = arrayOfInputs(13)  'D13
    parMotor1Offset = arrayOfInputs(14)  'D14
    parMotor2Offset = arrayOfInputs(15)  'D15
    parMotor1Length = arrayOfInputs(16)  'D16
    parActuatorWidth = arrayOfInputs(17)     'D17
    parActuatorWidth1 = arrayOfInputs(18)    'D18
    parMotor3toCenter = arrayOfInputs(19)    'D19
    parActuatorHeight3 = arrayOfInputs(20)    'D20
    parInsulationThickness = arrayOfInputs(21)
    parNPD = arrayOfInputs(22)
    parEndPreparation = arrayOfInputs(23)
    parScheduleThickness = arrayOfInputs(24)
    parEndStandard = arrayOfInputs(25)
    parPressureRating = arrayOfInputs(26)
    parFlowDirection = arrayOfInputs(27)
    parPortIndex1 = arrayOfInputs(28)
    parNpd1 = arrayOfInputs(29)
    parEndPreparation1 = arrayOfInputs(30)
    parScheduleThickness1 = arrayOfInputs(31)
    parEndStandard1 = arrayOfInputs(32)
    parPressureRating1 = arrayOfInputs(33)
    parFlowDirection1 = arrayOfInputs(34)
    parPortIndex2 = arrayOfInputs(35)
    parNpd2 = arrayOfInputs(36)
    parEndPreparation2 = arrayOfInputs(37)
    parScheduleThickness2 = arrayOfInputs(38)
    parEndStandard2 = arrayOfInputs(39)
    parPressureRating2 = arrayOfInputs(40)
    parFlowDirection2 = arrayOfInputs(41)
    parOperationHeight = arrayOfInputs(42)
    parId1 = arrayOfInputs(43)
    parId2 = arrayOfInputs(44)
    parNPDUnitType = arrayOfInputs(45)
    parNPDUnitType1 = arrayOfInputs(46)
    parNPDUnitType2 = arrayOfInputs(47)
  
    iOutput = 0
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output 1(Relief Valve Operation aspect)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    Dim ObjOpEnvelope As Object
    stPoint.Set -parMotor4Length, parActuatorHeight2, -parHandWheelDiameter / 2
    enPoint.Set parHandWheelDiameter / 2, parActuatorHeight2 + parOperationHeight, parMotor1Length
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



