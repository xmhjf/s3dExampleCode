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
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 8 2004
'   Description:
'     This is PDS on-the-fly I10A  With Upstream Valve Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   11.Apr.2005     svsmylav        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                             Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'   08.SEP.2006     KKC             DI-95670  Replace names with initials in all revision history sheets and symbols
'   25.Sep.2006     KKC             TR-102869 Modified symbol code to use both Face to Face
'                                             and Face to Center Dimensions
'  2.Jul.2007       KKC                     TR-123363: Corrected the symbol not to fail the V6.1 Instances
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const MIN_LENGTH = 0.0001
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
    Dim ObjInsValveBody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    
    Dim parCylinderLength As Double
    Dim parFacetoFace As Double
    Dim parWidth As Double
    Dim parHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim lPartDataBasis As Long

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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2) 'A
    parWidth = arrayOfInputs(3)      'B
    parHeight = arrayOfInputs(4)     'C
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
'    parFacetoCenter = arrayOfInputs(26)
'    parId1 = arrayOfInputs(27)
'    parId2 = arrayOfInputs(28)
'    parNPDUnitType = arrayOfInputs(29)
'    parNPDUnitType1 = arrayOfInputs(30)
'    parNPDUnitType2 = arrayOfInputs(31)
'   Check to ensure both V6.1 and V7 symbol definitions work
    If UBound(arrayOfInputs) >= 31 Then
        'In V7, Index 27 onwards are textual inputs
        parId1 = arrayOfInputs(27)
        parId2 = arrayOfInputs(28)
        parNPDUnitType = arrayOfInputs(29)
        parNPDUnitType1 = arrayOfInputs(30)
        parNPDUnitType2 = arrayOfInputs(31)
    Else
        'In V6.1,Index 26 onwards are textual inputs
        parId1 = arrayOfInputs(26)
        parId2 = arrayOfInputs(27)
        parNPDUnitType = arrayOfInputs(28)
        parNPDUnitType1 = arrayOfInputs(29)
        parNPDUnitType2 = arrayOfInputs(30)
    End If
    
    iOutput = 0
     
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
     
    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then
       parFacetoFace = arrayOfInputs(2)
    ElseIf lPartDataBasis = 10 Then
       parFacetoCenter = arrayOfInputs(26)
       parFacetoFace = 2 * parFacetoCenter
    End If

    Dim dBody_length As Double
    Dim dNoz_len  As Double

    dBody_length = parFacetoFace * 0.6
    dNoz_len = (parFacetoFace - dBody_length) / 2#
    If CmpDblLessThan(dNoz_len, MIN_LENGTH) Then dNoz_len = MIN_LENGTH

' Insert your code for output 1(Insulation Port1 Side)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    Point1.Set -parFacetoFace / 2, 0, 0
    Point2.Set Point1.x + flangeThick(parPortIndex1) + parInsulationThickness, Point1.y, Point1.z
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        Point2.Set Point1.x + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, Point1.y, Point1.z
    End If
    Dim dInsulationDiameter  As Double
    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then dInsulationDiameter = flangeDiam(parPortIndex1) + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, Point1, Point2, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for body Port1 Side)
    Point1.Set -parFacetoFace / 2 + flangeThick(parPortIndex1), 0, 0
    Point2.Set -parFacetoFace / 2 + dNoz_len, 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, Point1, Point2, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 4(Insulation Port2 Side)
    Point1.Set parFacetoFace / 2, 0, 0
    Point2.Set Point1.x - flangeThick(parPortIndex1) - parInsulationThickness, Point1.y, Point1.z
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        Point2.Set Point1.x - flangeThick(parPortIndex1) - parInsulationThickness - 0.0001, Point1.y, Point1.z
    End If
    dInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then dInsulationDiameter = flangeDiam(parPortIndex2) + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, Point1, Point2, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 5(Insulation for body Port2 Side)
    Point1.Set parFacetoFace / 2 - flangeThick(parPortIndex1), 0, 0
    Point2.Set parFacetoFace / 2 - dNoz_len, 0, 0
    dInsulationDiameter = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, Point1, Point2, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing

' Insert your code for output 1(Insulation for Valve Body)
    Point1.Set -dBody_length / 2 - parInsulationThickness, -parHeight / 2 _
            - parInsulationThickness, -parWidth / 2 - parInsulationThickness
    Point2.Set dBody_length / 2 + parInsulationThickness, parHeight / 2 _
            + parInsulationThickness, parWidth / 2 + parInsulationThickness

    Set ObjInsValveBody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsValveBody
    Set ObjInsValveBody = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
