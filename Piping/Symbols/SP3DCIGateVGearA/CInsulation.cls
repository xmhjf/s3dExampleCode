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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Wednusday, 21 Febraury 2007
'   Description:    Insulated Valve, construct box FacetoFace x (Insulation + Valve Height)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
''  21.Feb.2007      RRK             TR-113129 Added Insulation Aspect
'   07-Mar-07      RRK           TR-116058: Added 'If' condition to suppport new optional input without failing the prior to V7 instances
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjInsulatedBox As Object
    
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim dValveBottomtoCL As Double

    Dim parValveHeight          As Double
    Dim parHandwheelDiameter    As Double
    Dim parHandwheelAngle       As Double
    Dim parFacetoFace As Double
    
    Dim Npd                     As Double
    Dim EndPreparation          As Long
    Dim ScheduleThickness       As Long
    Dim EndStandard             As Long
    Dim PressureRating          As Long
    Dim FlowDirection           As DistribFlow
    
    Dim PortIndex1              As Long
    Dim Npd1                    As Double
    Dim EndPreparation1         As Long
    Dim ScheduleThickness1      As Long
    Dim EndStandard1            As Long
    Dim PressureRating1         As Long
    Dim FlowDirection1          As DistribFlow

    Dim PortIndex2              As Long
    Dim Npd2                    As Double
    Dim EndPreparation2         As Long
    Dim ScheduleThickness2      As Long
    Dim EndStandard2            As Long
    Dim PressureRating2         As Long
    Dim FlowDirection2          As DistribFlow
    Dim parInsulationThickness As Double
    
    Dim Id1                     As String
    Dim Id2                     As String

    Dim NpdUnitType             As String
    Dim NpdUnitType1            As String
    Dim NpdUnitType2            As String

    Dim TerminationClass1       As Long
    Dim TerminationSubClass1    As Long
    Dim SchedulePractice1       As Long
    Dim EndPractice1            As Long
    Dim RatingPractice1         As Long
    Dim TerminationClass2       As Long
    Dim TerminationSubClass2    As Long
    Dim SchedulePractice2       As Long
    Dim EndPractice2            As Long
    Dim RatingPractice2         As Long
    
    'Numeric Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parValveHeight = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    Npd = arrayOfInputs(6)
    EndPreparation = arrayOfInputs(7)
    ScheduleThickness = arrayOfInputs(8)
    EndStandard = arrayOfInputs(9)
    PressureRating = arrayOfInputs(10)
    FlowDirection = arrayOfInputs(11)
    PortIndex1 = arrayOfInputs(12)
    Npd1 = arrayOfInputs(13)
    EndPreparation1 = arrayOfInputs(14)
    ScheduleThickness1 = arrayOfInputs(15)
    EndStandard1 = arrayOfInputs(16)
    PressureRating1 = arrayOfInputs(17)
    FlowDirection1 = arrayOfInputs(18)
    PortIndex2 = arrayOfInputs(19)
    Npd2 = arrayOfInputs(20)
    EndPreparation2 = arrayOfInputs(21)
    ScheduleThickness2 = arrayOfInputs(22)
    EndStandard2 = arrayOfInputs(23)
    PressureRating2 = arrayOfInputs(24)
    FlowDirection2 = arrayOfInputs(25)
    Id1 = arrayOfInputs(26)
    Id2 = arrayOfInputs(27)
    NpdUnitType = arrayOfInputs(28)
    NpdUnitType1 = arrayOfInputs(29)
    NpdUnitType2 = arrayOfInputs(30)

    If UBound(arrayOfInputs) >= 31 Then 'Not to fail the prior to V7 instances by accessing index 31
        parInsulationThickness = arrayOfInputs(31)
    End If
    
    iOutput = 0
    dValveBottomtoCL = parFacetoFace * 0.85
   
 ' Insert your code for output 14(InsulatedBox)
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition
    oStPoint.Set -parFacetoFace / 2, -(dValveBottomtoCL + parInsulationThickness), _
                                    (dValveBottomtoCL + parInsulationThickness)
    oEnPoint.Set parFacetoFace / 2, (parValveHeight - dValveBottomtoCL / 2 + parInsulationThickness), _
                                    -(dValveBottomtoCL + parInsulationThickness)
    
    Set ObjInsulatedBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBox
    Set ObjInsulatedBox = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub



