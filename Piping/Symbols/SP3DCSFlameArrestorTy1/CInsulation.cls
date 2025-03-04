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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Five Insulation aspect outputs are created. ObjInsFABody is created using PlaceBox
' function and the other insulation outputs are created using PlaceCylinder function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   11.Apr.2005     svsmylav        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                           Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjInsFABody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parFlArrestorBodyLength1 As Double
    Dim parFlArrestorBodyLength2 As Double
    Dim parFlArrestorBodyHeight1 As Double
    Dim parFlArrestorBodyHeight2 As Double
    Dim parFlArrestorBodyWidth1 As Double
    Dim parFlArrestorBodyWidth2 As Double
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
    
    Dim Point1   As New AutoMath.DPosition
    Dim Point2   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'P2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parOffset1 = arrayOfInputs(4)       'P4
    parOffset2 = arrayOfInputs(5)       'P5
    parFlArrestorBodyLength1 = arrayOfInputs(6)   'P6
    parFlArrestorBodyLength2 = arrayOfInputs(7)     'P7
    parFlArrestorBodyHeight1 = arrayOfInputs(8)     'P8
    parFlArrestorBodyHeight2 = arrayOfInputs(9)     'P9
    parFlArrestorBodyWidth1 = arrayOfInputs(10)     'P10
    parFlArrestorBodyWidth2 = arrayOfInputs(11)     'P11
    parInsulationThickness = arrayOfInputs(12)
    parNPD = arrayOfInputs(13)
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parPortIndex1 = arrayOfInputs(19)
    parNpd1 = arrayOfInputs(20)
    parEndPreparation1 = arrayOfInputs(21)
    parScheduleThickness1 = arrayOfInputs(22)
    parEndStandard1 = arrayOfInputs(23)
    parPressureRating1 = arrayOfInputs(24)
    parFlowDirection1 = arrayOfInputs(25)
    parPortIndex2 = arrayOfInputs(26)
    parNpd2 = arrayOfInputs(27)
    parEndPreparation2 = arrayOfInputs(28)
    parScheduleThickness2 = arrayOfInputs(29)
    parEndStandard2 = arrayOfInputs(30)
    parPressureRating2 = arrayOfInputs(31)
    parFlowDirection2 = arrayOfInputs(32)
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parNPDUnitType = arrayOfInputs(35)
    parNPDUnitType1 = arrayOfInputs(36)
    parNPDUnitType2 = arrayOfInputs(37)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Flame Arrestor Body)

    Point1.Set -parFlArrestorBodyLength1 - parInsulationThickness, _
                        -parOffset1 - parFlArrestorBodyHeight1 - parInsulationThickness, _
                        parFlArrestorBodyWidth2 + parInsulationThickness
    Point2.Set parFlArrestorBodyLength2 + parInsulationThickness, _
                        parFlArrestorBodyHeight2 - parOffset2 + parInsulationThickness, _
                        -parFlArrestorBodyWidth1 - parInsulationThickness
    Set ObjInsFABody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFABody
    Set ObjInsFABody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    Point1.Set -parFace1toCenter, -parOffset1, 0
    Point2.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness, -parOffset1, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        Point2.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, -parOffset1, 0
    End If
    parInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then parInsulationDiameter = flangeDiam(parPortIndex1) + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for body Port1 Side)
    Point1.Set -parFace1toCenter + flangeThick(parPortIndex1), -parOffset1, 0
    Point2.Set -parFlArrestorBodyLength1, -parOffset1, 0
    parInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 4(Insulation Port2 Side)
    Point1.Set parFace2toCenter, parOffset2, 0
    Point2.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness, parOffset2, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        Point2.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness + 0.0001, parOffset2, 0
    End If
    parInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then parInsulationDiameter = flangeDiam(parPortIndex2) + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 5(Insulation for body Port2 Side)
    Point1.Set parFace2toCenter - flangeThick(parPortIndex2), parOffset2, 0
    Point2.Set parFlArrestorBodyLength2, parOffset2, 0
    parInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
