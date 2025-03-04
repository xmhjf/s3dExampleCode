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
'     This is PDS on-the-fly G51  Y Strainer Symbol.
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
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    Dim ObjInsPort3 As Object
    Dim ObjInsBody3 As Object
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object
    
    Dim parFacetoFace As Double
    Dim parNozzleOffset As Double
    Dim parFacetoCenter As Double
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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)    'A
    parNozzleOffset = arrayOfInputs(3)  'B
    parFacetoCenter = arrayOfInputs(4)  'C
    parInsulationThickness = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
    parPortIndex3 = arrayOfInputs(26)
    parNpd3 = arrayOfInputs(27)
    parEndPreparation3 = arrayOfInputs(28)
    parScheduleThickness3 = arrayOfInputs(29)
    parEndStandard3 = arrayOfInputs(30)
    parPressureRating3 = arrayOfInputs(31)
    parFlowDirection3 = arrayOfInputs(32)
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parId3 = arrayOfInputs(35)
    parNPDUnitType = arrayOfInputs(36)
    parNPDUnitType1 = arrayOfInputs(37)
    parNPDUnitType2 = arrayOfInputs(38)
    parNPDUnitType3 = arrayOfInputs(39)
    
    Dim dBend_Angle As Double
    
    dBend_Angle = 1

    iOutput = 0

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

' Insert your code for output 1(Insulation for Port1)
    stPoint.Set -parFacetoCenter, 0, 0
    Dim dInsulationDiameter As Double
    enPoint.Set -parFacetoCenter + flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        enPoint.Set -parFacetoCenter + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0, 0
    End If
    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then dInsulationDiameter = flangeDiam(parPortIndex1) + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 2( Insulation for Body)
    Dim dF2C  As Double
    dF2C = parFacetoFace - parFacetoCenter
    stPoint.Set -(parFacetoCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set dF2C - flangeThick(parPortIndex2), 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 3(Insulation for Port2)
    stPoint.Set dF2C, 0, 0
    enPoint.Set dF2C - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        enPoint.Set dF2C - flangeThick(parPortIndex2) - parInsulationThickness - 0.0001, 0, 0
    End If
    dInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then dInsulationDiameter = flangeDiam(parPortIndex2) + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 4(Insulation for Port3)
    Dim dF3C  As Double
    dF3C = parNozzleOffset / Sin(dBend_Angle)
     
    stPoint.Set dF3C * Cos(PI / 4), _
                    -dF3C * Sin(PI / 4), 0
    Dim Port3CPoint   As New AutoMath.DPosition
    Port3CPoint.Set stPoint.x, stPoint.y, 0
    enPoint.Set Port3CPoint.x - (flangeThick(parPortIndex3) + parInsulationThickness) * Cos(PI / 4), _
                Port3CPoint.y - (flangeThick(parPortIndex3) + parInsulationThickness) * Sin(PI / 4), 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex3) = 0 And parInsulationThickness = 0 Then
        enPoint.Set Port3CPoint.x - (flangeThick(parPortIndex3) + parInsulationThickness + 0.0001) * Cos(PI / 4), _
                    Port3CPoint.y - (flangeThick(parPortIndex3) + parInsulationThickness + 0.0001) * Sin(PI / 4), 0
    End If
    dInsulationDiameter = pipeDiam(parPortIndex3) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then dInsulationDiameter = flangeDiam(parPortIndex3) + 2 * parInsulationThickness
    Set ObjInsPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort3
    Set ObjInsPort3 = Nothing
    
' Insert your code for output 5(Insulation for Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set (dF3C - flangeThick(parPortIndex3)) * Cos(PI / 4), _
                -(dF3C - flangeThick(parPortIndex3)) * Sin(PI / 4), 0
    dInsulationDiameter = pipeDiam(parPortIndex3) + 2 * parInsulationThickness
    Set ObjInsBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
                
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
     
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
