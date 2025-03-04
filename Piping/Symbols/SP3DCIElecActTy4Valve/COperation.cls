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
'   Creation Date:  Friday, September 19 - 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDS Eden Name=IEAM4,PDMS Category Name=IVCP)
'   The symbol consists of Eleven outputs; Ten physical outputs and One Insulation Output
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
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parMotor1Diameter As Double
    Dim parMotor2Diameter As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Length As Double
    Dim parMotor1Length As Double
    Dim parActuatorOffset As Double
    Dim parActuatorHeight1 As Double
    Dim parHandWheelDiameter As Double
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
  
    Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFacetoFace = arrayOfInputs(3)      'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parActuatorHeight = arrayOfInputs(5)      'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorLength = arrayOfInputs(7)      'D7
    parActuatorLength1 = arrayOfInputs(8)     'D8
    parMotor1Diameter = arrayOfInputs(9)      'D9
    parMotor2Diameter = arrayOfInputs(10)     'D10
    parMotor2Offset = arrayOfInputs(11)   'D11
    parMotor1Offset = arrayOfInputs(12)   'D12
    parMotor2Length = arrayOfInputs(13)   'D13
    parMotor1Length = arrayOfInputs(14)   'D14
    parActuatorOffset = arrayOfInputs(15)     'D15
    parActuatorHeight1 = arrayOfInputs(16)     'D16
    parHandWheelDiameter = arrayOfInputs(17)      'D17
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
    stPoint.Set -parActuatorLength, parActuatorHeight1, -parHandWheelDiameter / 2
    enPoint.Set parActuatorLength1, parOperationHeight + parActuatorHeight1, parMotor2Length + parActuatorOffset
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
