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
'   Author:          SSP
'   Creation Date:  Friday, Sep 5 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVHK)
'   The symbol consists of 11 outputs: 8 Physical outputs , 2 Insulation Outputs and one operation Output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   05.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
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
    
    Dim ObjOpEnvelope As Object

    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parHandWheelDiameter As Double
    Dim parPositionerHeight As Double
    Dim parPositionerOffset As Double
    Dim parPositionerOffset1 As Double
    Dim parPositionerOffset2 As Double
    Dim parValveOutertoHandWheel As Double
    Dim parValveOutertoCenter As Double
    Dim parPositionerLength As Double
    Dim parRotation As Double
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
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFacetoFace = arrayOfInputs(3)     'D3
    parCylHeight = arrayOfInputs(4)     'D4
    parActuatorHeight = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)  'D6
    parHandWheelOffset = arrayOfInputs(7)   'D8
    parHandWheelDiameter = arrayOfInputs(8) 'D9
    parPositionerHeight = arrayOfInputs(9)      'D10
    parPositionerOffset = arrayOfInputs(10)      'D11
    parPositionerOffset1 = arrayOfInputs(11)     'D12
    parPositionerOffset2 = arrayOfInputs(12)     'D14
    parValveOutertoHandWheel = arrayOfInputs(13) 'D16
    parValveOutertoCenter = arrayOfInputs(14)       'D19
    parPositionerLength = arrayOfInputs(15)     'D20
    parRotation = arrayOfInputs(16)             'Rotation of Operator
    parInsulationThickness = arrayOfInputs(17)
    'Input parameters for Dynamic Nozzles
    parNPD = arrayOfInputs(18)
    parEndPreparation = arrayOfInputs(19)
    parScheduleThickness = arrayOfInputs(20)
    parEndStandard = arrayOfInputs(21)
    parPressureRating = arrayOfInputs(22)
    parFlowDirection = arrayOfInputs(23)
    parPortIndex1 = arrayOfInputs(24)
    parNpd1 = arrayOfInputs(25)
    parEndPreparation1 = arrayOfInputs(26)
    parScheduleThickness1 = arrayOfInputs(27)
    parEndStandard1 = arrayOfInputs(28)
    parPressureRating1 = arrayOfInputs(29)
    parFlowDirection1 = arrayOfInputs(30)
    parPortIndex2 = arrayOfInputs(31)
    parNpd2 = arrayOfInputs(32)
    parEndPreparation2 = arrayOfInputs(33)
    parScheduleThickness2 = arrayOfInputs(34)
    parEndStandard2 = arrayOfInputs(35)
    parPressureRating2 = arrayOfInputs(36)
    parFlowDirection2 = arrayOfInputs(37)
    parOperationHeight = arrayOfInputs(38)     'D7
    parOperationDiameter = arrayOfInputs(39) 'D17
    parOperationDiameter1 = arrayOfInputs(40) 'D18
    
    parId1 = arrayOfInputs(41)
    parId2 = arrayOfInputs(42)
    parNPDUnitType = arrayOfInputs(43)
    parNPDUnitType1 = arrayOfInputs(44)
    parNPDUnitType2 = arrayOfInputs(45)

    iOutput = 0

 ' Insert your code for output (Envelop for Operation )
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim dOpDiameter As Double
    
    stPoint.Set 0, parValveOutertoCenter + parActuatorHeight, 0
    enPoint.Set 0, parValveOutertoCenter + parActuatorHeight + parOperationHeight, 0
    
    If parOperationDiameter >= parOperationDiameter1 Then
        dOpDiameter = parOperationDiameter
    Else
        dOpDiameter = parOperationDiameter1
    End If
    
    Set ObjOpEnvelope = PlaceCylinder(m_OutputColl, stPoint, enPoint, dOpDiameter, True)
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
