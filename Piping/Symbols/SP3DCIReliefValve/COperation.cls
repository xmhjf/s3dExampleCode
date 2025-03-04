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
'   The symbol is prepared based on (SRV1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
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
    
    Dim iOutput As Integer
    Dim ObjOpEnvelope As Object
    Dim ObjNozz1BodyIns As Object
    Dim ObjNozz2BodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorWidth As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationHeight As Double
    Dim parOperationLength As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double
    
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
    parFace1toCenter = arrayOfInputs(2)     'R1
    parFace2toCenter = arrayOfInputs(3)     'R2
    parInstrumentHeight = arrayOfInputs(4)      'R3
    parInstrumentLength = arrayOfInputs(5)      'R5
    parInstrumentWidth = arrayOfInputs(6)       'R6
    parInstrumentWidth1 = arrayOfInputs(7)      'R7
    parInstrumentHeight1 = arrayOfInputs(8)    'R4
    parActuatorHeight = arrayOfInputs(9)    'A2
    parActuatorHeight1 = arrayOfInputs(10)   'A3
    parActuatorLength = arrayOfInputs(11)    'A4
    parActuatorWidth = arrayOfInputs(12)     'A5
    parInstrumentDiameter1 = arrayOfInputs(13)  'R8
    parInsulationThickness = arrayOfInputs(14)
    parRotation = arrayOfInputs(15) 'A1
    parNPD = arrayOfInputs(16)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(17)
    parScheduleThickness = arrayOfInputs(18)
    parEndStandard = arrayOfInputs(19)
    parPressureRating = arrayOfInputs(20)
    parFlowDirection = arrayOfInputs(21)
    parPortIndex1 = arrayOfInputs(22)
    parNpd1 = arrayOfInputs(23)
    parEndPreparation1 = arrayOfInputs(24)
    parScheduleThickness1 = arrayOfInputs(25)
    parEndStandard1 = arrayOfInputs(26)
    parPressureRating1 = arrayOfInputs(27)
    parFlowDirection1 = arrayOfInputs(28)
    parPortIndex2 = arrayOfInputs(29)
    parNpd2 = arrayOfInputs(30)
    parEndPreparation2 = arrayOfInputs(31)
    parScheduleThickness2 = arrayOfInputs(32)
    parEndStandard2 = arrayOfInputs(33)
    parPressureRating2 = arrayOfInputs(34)
    parFlowDirection2 = arrayOfInputs(35)
    parOperationHeight = arrayOfInputs(36) 'I1
    parOperationLength = arrayOfInputs(37) 'I2
    parOperationWidth = arrayOfInputs(38)  'I3
    parOperationWidth1 = arrayOfInputs(39)  'I4
' Text inputs
    parId1 = arrayOfInputs(40)
    parId2 = arrayOfInputs(41)
    parNPDUnitType = arrayOfInputs(42)
    parNPDUnitType1 = arrayOfInputs(43)
    parNPDUnitType2 = arrayOfInputs(44)
    
    iOutput = 0
    
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
'The origin is considered to be at intersection of nozzle axis
 ' Insert your code for output 1(Operation Envelope)
    
    oStPoint.Set -parFace1toCenter, -parFace2toCenter, -parOperationWidth1
    oEnPoint.Set parOperationLength, parOperationHeight, parOperationWidth
    
    Set ObjOpEnvelope = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
 
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

