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
'   The symbol is prepared based on (SPR1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy    who          change description
'   -----------   -----         ------------------
'   8-Nov-07      svsmylav/KKC  DI-130618:Changed the parameter names of parPositionerOrientation and
'                               parHandWheelOrientation to parPositionerRotAngle and parHandWheelStepRotAngle respectively
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
    Dim ObjValveCone1 As Object
    Dim ObjValveCone2 As Object
    Dim ObjStemRod As Object
    Dim ObjValveCylinder As Object
    Dim ObjActuatorRect As Object
    Dim ObjActuatorCyl As Object

    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorWidth2 As Double
    Dim parHandWhOutertoCenter As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parHandWheelOffset1 As Double
    Dim parPositionerLength As Double
    Dim parPositionerHeight As Double
    Dim parPositionerWidth As Double
    Dim parPositionerOffset As Double
    Dim parPositionerOffset1 As Double
    Dim parRotation As Double
    Dim parInsulationThickness As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parOperationHeight As Double
    Dim parPositionerRotAngle As Double
    Dim parHandWheelStepRotAngle As Double
    
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


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'V1
    parFacetoFace = arrayOfInputs(3)   'V2
    parCylHeight = arrayOfInputs(4) 'V3
    parActuatorLength = arrayOfInputs(5) 'A11
    parActuatorLength1 = arrayOfInputs(6) 'A10
    parActuatorLength2 = arrayOfInputs(7) 'A9
    parActuatorDiameter = arrayOfInputs(8) 'A3
    parActuatorHeight = arrayOfInputs(9)   'A2
    parActuatorHeight1 = arrayOfInputs(10) 'A7
    parActuatorHeight2 = arrayOfInputs(11) 'A8
    parActuatorWidth = arrayOfInputs(12) 'A4
    parActuatorWidth1 = arrayOfInputs(13) 'A5
    parActuatorWidth2 = arrayOfInputs(14) 'A6
    parHandWhOutertoCenter = arrayOfInputs(15) 'H2
    parHandWheelDiameter = arrayOfInputs(16) 'H3
    parHandWheelOffset = arrayOfInputs(17) 'H4
    parHandWheelOffset1 = arrayOfInputs(18) 'H5
    parPositionerLength = arrayOfInputs(19) 'P6
    parPositionerHeight = arrayOfInputs(20) 'P4
    parPositionerWidth = arrayOfInputs(21) 'P2
    parPositionerOffset = arrayOfInputs(22) 'P5
    parPositionerOffset1 = arrayOfInputs(23) 'P3
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
    
    oStPoint.Set 0, 0, parActuatorHeight
    oEnPoint.Set 0, 0, parActuatorHeight + parOperationHeight
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


