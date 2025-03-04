VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          SSP
'   Creation Date:  Friday, Jan 3 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVCV).
'   Physical aspect consist of a Stem, 5 Cylinders, 1 hand wheel, 2 cones
'   and 2 Nozzles. Insulation aspect consist of a 2 Cylinders for the valve body.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub
Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1Ins As Object
    Dim ObjBodyCone2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter  As Double
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
    
    Dim parOperationHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFace2toCenter = arrayOfInputs(3)        'D3
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
    parNPD = arrayOfInputs(22)                       'Input parameters for Dynamic Nozzles
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
    parOperationHeight = arrayOfInputs(42)     'D21
    'Text Inputs
    parId1 = arrayOfInputs(43)
    parId2 = arrayOfInputs(44)
    parNPDUnitType = arrayOfInputs(45)
    parNPDUnitType1 = arrayOfInputs(46)
    parNPDUnitType2 = arrayOfInputs(47)
    
    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 9(Body Cone 1 Insulation)
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                            
    InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then InsulationDia = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 10(Body Cone 2 Insulation)
    InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then InsulationDia = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, parCylHeight + parInsulationThickness, 0
    Set ObjBodyCone2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output    parCylHeight,
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2Ins
    Set ObjBodyCone2Ins = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub




