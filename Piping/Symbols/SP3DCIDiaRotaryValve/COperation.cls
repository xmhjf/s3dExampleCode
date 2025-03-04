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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          KKC
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on (SDR1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   08.Nov.07     RRK/KKC    DI-130618: Changed the parameter names of parPositionerOrientation and
'                            parHandWheelOrientation to parPositionerRotAngle and parHandWheelStepRotAngle respectively
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
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjValCylinder As Object
    Dim ObjStem As Object
    Dim ObjActuatorBox As Object
    Dim ObjActuatorCone As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double             'V1
    Dim parFacetoFace As Double                'V2
    Dim parCylHeight As Double                 'V3
    Dim parActuatorWidth As Double             'A2
    Dim parActuatorDiameter As Double          'A3
    Dim parActuatorHeight As Double            'A4
    Dim parActuatorHeight1 As Double           'A5
    Dim parActuatorHeight2 As Double           'A6
    Dim parActuatorLength As Double            'A9
    Dim parActuatorLength1 As Double           'A10
    Dim parActuatorWidth1 As Double            'A7
    Dim parActuatorWidth2 As Double            'A8
    Dim parActuatorLength2 As Double           'A11
    Dim parHandWheelOffset As Double           'H4
    Dim parHandWhOutertoCenter As Double           'H2
    Dim parHandWheelDiameter As Double         'H3
    Dim parHandWheelOffset1 As Double           'H5
    Dim parPositionerHeight As Double          'P2
    Dim parPositionerWidth As Double           'P4
    Dim parPositionerOffset As Double          'P3
    Dim parPositionerOffset1 As Double         'P6
    Dim parPositionerOffset2 As Double         'P5
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parOperationHeight As Double
    Dim parPositionerRotAngle As Double
    Dim parHandWheelStepRotAngle As Double
    
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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorWidth = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight = arrayOfInputs(7)
    parActuatorHeight1 = arrayOfInputs(8)
    parActuatorHeight2 = arrayOfInputs(9)
    parActuatorLength = arrayOfInputs(10)
    parActuatorLength1 = arrayOfInputs(11)
    parActuatorWidth1 = arrayOfInputs(12)
    parActuatorWidth2 = arrayOfInputs(13)
    parActuatorLength2 = arrayOfInputs(14)
    parHandWheelOffset = arrayOfInputs(15)
    parHandWheelOffset1 = arrayOfInputs(16)
    parHandWhOutertoCenter = arrayOfInputs(17)
    parHandWheelDiameter = arrayOfInputs(18)
    parPositionerHeight = arrayOfInputs(19)
    parPositionerWidth = arrayOfInputs(20)
    parPositionerOffset = arrayOfInputs(21)
    parPositionerOffset1 = arrayOfInputs(22)
    parPositionerOffset2 = arrayOfInputs(23)
    parInsulationThickness = arrayOfInputs(24)
    parRotation = arrayOfInputs(25)
    parNPD = arrayOfInputs(26)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(27)
    parScheduleThickness = arrayOfInputs(28)
    parEndStandard = arrayOfInputs(29)
    parPressureRating = arrayOfInputs(30)
    parFlowDirection = arrayOfInputs(31)
    parPortIndex1 = arrayOfInputs(32)
    parNpd1 = arrayOfInputs(33)
    parEndPreparation1 = arrayOfInputs(34)
    parScheduleThickness1 = arrayOfInputs(35)
    parEndStandard1 = arrayOfInputs(36)
    parPressureRating1 = arrayOfInputs(37)
    parFlowDirection1 = arrayOfInputs(38)
    parPortIndex2 = arrayOfInputs(39)
    parNpd2 = arrayOfInputs(40)
    parEndPreparation2 = arrayOfInputs(41)
    parScheduleThickness2 = arrayOfInputs(42)
    parEndStandard2 = arrayOfInputs(43)
    parPressureRating2 = arrayOfInputs(44)
    parFlowDirection2 = arrayOfInputs(45)
    parOperationDiameter = arrayOfInputs(46)
    parOperationDiameter1 = arrayOfInputs(47)
    parOperationHeight = arrayOfInputs(48)
    parPositionerRotAngle = arrayOfInputs(49)
    parHandWheelStepRotAngle = arrayOfInputs(50)
'   Text Inputs
    parId1 = arrayOfInputs(51)
    parId2 = arrayOfInputs(52)
    parNPDUnitType = arrayOfInputs(53)
    parNPDUnitType1 = arrayOfInputs(54)
    parNPDUnitType2 = arrayOfInputs(55)
    
    
    iOutput = 0
    
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim dOpDiameter As Double
    
    oStPoint.Set 0, 0, parActuatorWidth
    oEnPoint.Set 0, 0, parActuatorWidth + parOperationHeight
    'The Greater of the Operation Diameter and the Operation Diameter 1 is
    'taken as the Diameter of the Operation Envelope
    If CmpDblGreaterthanOrEqualTo(parOperationDiameter, parOperationDiameter1) Then
        dOpDiameter = parOperationDiameter
    Else
        dOpDiameter = parOperationDiameter1
    End If
    Dim ObjOpEnvelope As Object
    Set ObjOpEnvelope = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dOpDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpEnvelope
    Set ObjOpEnvelope = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
