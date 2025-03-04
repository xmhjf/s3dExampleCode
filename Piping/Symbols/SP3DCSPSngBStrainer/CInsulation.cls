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
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         Sundar(svsmylav)
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S3A5  Single Basket Strainer Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   11.Apr.2005     Sundar(svsmylav)        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                           Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedCylinder1 As Object
    Dim ObjInsulatedCylinder2 As Object
    
    Dim parFacetoCenter As Double
    Dim parCylinderDiameter As Double
    Dim parFacetoFace As Double
    Dim parCylOffset As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'A
    parCylinderDiameter = arrayOfInputs(3)  'B
    parFacetoFace = arrayOfInputs(4)    'C
    parCylOffset = arrayOfInputs(5)     'D
    parInsulationThickness = arrayOfInputs(6)
    parNPD = arrayOfInputs(7)
    parEndPreparation = arrayOfInputs(8)
    parScheduleThickness = arrayOfInputs(9)
    parEndStandard = arrayOfInputs(10)
    parPressureRating = arrayOfInputs(11)
    parFlowDirection = arrayOfInputs(12)
    parPortIndex1 = arrayOfInputs(13)
    parNpd1 = arrayOfInputs(14)
    parEndPreparation1 = arrayOfInputs(15)
    parScheduleThickness1 = arrayOfInputs(16)
    parEndStandard1 = arrayOfInputs(17)
    parPressureRating1 = arrayOfInputs(18)
    parFlowDirection1 = arrayOfInputs(19)
    parPortIndex2 = arrayOfInputs(20)
    parNpd2 = arrayOfInputs(21)
    parEndPreparation2 = arrayOfInputs(22)
    parScheduleThickness2 = arrayOfInputs(23)
    parEndStandard2 = arrayOfInputs(24)
    parPressureRating2 = arrayOfInputs(25)
    parFlowDirection2 = arrayOfInputs(26)
    parPortIndex3 = arrayOfInputs(27)
    parNpd3 = arrayOfInputs(28)
    parEndPreparation3 = arrayOfInputs(29)
    parScheduleThickness3 = arrayOfInputs(30)
    parEndStandard3 = arrayOfInputs(31)
    parPressureRating3 = arrayOfInputs(32)
    parFlowDirection3 = arrayOfInputs(33)
    parId1 = arrayOfInputs(34)
    parId2 = arrayOfInputs(35)
    parId3 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)
    parNPDUnitType3 = arrayOfInputs(40)
    
    iOutput = 0

' Insert your code for output 5(Insulation for Nozzle1 flange)
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    Dim dInsulationDiameter As Double
    dInsulationDiameter = flangeDiam(parPortIndex1) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex1) < pipeDiam(parPortIndex1) Then dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -parFacetoFace / 2 + flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        enPoint.Set -parFacetoFace / 2 + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0, 0
    End If

    Dim ObjNozzle1FlIns As Object
    Set ObjNozzle1FlIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1FlIns
    Set ObjNozzle1FlIns = Nothing
    
' Insert your code for output 6(Insulation for Nozzle1 body)
    stPoint.Set -parFacetoFace / 2 + flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    Dim ObjNozzle1BodyIns As Object
    Set ObjNozzle1BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1BodyIns
    Set ObjNozzle1BodyIns = Nothing
    
' Insert your code for output 7(Insulation for Nozzle2 flange)
    dInsulationDiameter = flangeDiam(parPortIndex2) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex2) < pipeDiam(parPortIndex2) Then dInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2 - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        enPoint.Set parFacetoFace / 2 - flangeThick(parPortIndex2) - parInsulationThickness - 0.0001, 0, 0
    End If

    Dim ObjNozzle2FlIns As Object
    Set ObjNozzle2FlIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2FlIns
    Set ObjNozzle2FlIns = Nothing
    
' Insert your code for output 8(Insulation for Nozzle2 body)
    stPoint.Set parFacetoFace / 2 - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    Dim ObjNozzle2BodyIns As Object
    Set ObjNozzle2BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2BodyIns
    Set ObjNozzle2BodyIns = Nothing
    
' Insert your code for output 9(Insulation for Nozzle3 flange)
    dInsulationDiameter = flangeDiam(parPortIndex3) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex3) < pipeDiam(parPortIndex3) Then dInsulationDiameter = pipeDiam(parPortIndex3) + 2 * parInsulationThickness
    stPoint.Set 0, parFacetoCenter, 0
    enPoint.Set 0, parFacetoCenter - flangeThick(parPortIndex3) - parInsulationThickness, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex3) = 0 And parInsulationThickness = 0 Then
        enPoint.Set 0, parFacetoCenter - flangeThick(parPortIndex3) - parInsulationThickness - 0.0001, 0
    End If
    Dim ObjNozzle3FlIns As Object
    Set ObjNozzle3FlIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3FlIns
    Set ObjNozzle3FlIns = Nothing
    
' Insert your code for output 10(Insulation for Nozzle3 body)
    stPoint.Set 0, parFacetoCenter - flangeThick(parPortIndex3) - parInsulationThickness, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex3) + 2 * parInsulationThickness
    Dim ObjNozzle3BodyIns As Object
    Set ObjNozzle3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3BodyIns
    Set ObjNozzle3BodyIns = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
