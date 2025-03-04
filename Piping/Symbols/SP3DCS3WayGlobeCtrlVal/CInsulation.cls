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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = GLO3W,
'   SN = V13 at page no 239 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    ReturnMax3 = MaxValue
End Function
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

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
    Dim ObjInsHorBody As Object
    Dim ObjInsVertBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ellipseMajorRadius      As Double
    Dim EllipseMinorRadius      As Double
    Dim parInsulationDiameter   As Double

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

    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter  = arrayOfInputs(2)
    parFace2toCenter  = arrayOfInputs(3)
    parFace3toCenter  = arrayOfInputs(4)
    parInsulationThickness  = arrayOfInputs(5)
    parHandwheelAngle  = arrayOfInputs(6)
    parNpd  = arrayOfInputs(7)
    parEndPreparation  = arrayOfInputs(8)
    parScheduleThickness  = arrayOfInputs(9)
    parEndStandard  = arrayOfInputs(10)
    parPressureRating  = arrayOfInputs(11)
    parFlowDirection  = arrayOfInputs(12)
    parPortIndex1  = arrayOfInputs(13)
    parNpd1  = arrayOfInputs(14)
    parEndPreparation1  = arrayOfInputs(15)
    parScheduleThickness1  = arrayOfInputs(16)
    parEndStandard1  = arrayOfInputs(17)
    parPressureRating1  = arrayOfInputs(18)
    parFlowDirection1  = arrayOfInputs(19)
    parPortIndex2  = arrayOfInputs(20)
    parNpd2  = arrayOfInputs(21)
    parEndPreparation2  = arrayOfInputs(22)
    parScheduleThickness2  = arrayOfInputs(23)
    parEndStandard2  = arrayOfInputs(24)
    parPressureRating2  = arrayOfInputs(25)
    parFlowDirection2  = arrayOfInputs(26)
    parPortIndex3  = arrayOfInputs(27)
    parNpd3  = arrayOfInputs(28)
    parEndPreparation3  = arrayOfInputs(29)
    parScheduleThickness3  = arrayOfInputs(30)
    parEndStandard3  = arrayOfInputs(31)
    parPressureRating3  = arrayOfInputs(32)
    parFlowDirection3  = arrayOfInputs(33)
    parId1  = arrayOfInputs(34)
    parId2  = arrayOfInputs(35)
    parId3  = arrayOfInputs(36)
    parNpdUnitType  = arrayOfInputs(37)
    parNpdUnitType1  = arrayOfInputs(38)
    parNpdUnitType2  = arrayOfInputs(39)
    parNpdUnitType3  = arrayOfInputs(40)

    iOutput = 0

' Insert your code for output 1(Insulation for Valve Body - Horizontal)
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

    parInsulationDiameter = ReturnMax4(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1), _
                                                pipeDiam(parPortIndex2), flangeDiam(parPortIndex2)) + _
                            2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjInsHorBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHorBody
    Set ObjInsHorBody = Nothing
    
' Insert your code for output 2(Insulation for Valve Body - Vertical)
'   Extend the insulation along negative Y axis by ellipse minor radius plus parInsulationThickness
'   Consider Ellipse Majoj Axis is 0.6 times the larger face to center distance /2 and
'   Minor axis as 0.4 times the same center distance.
    ellipseMajorRadius = (0.6 * ReturnMax3(parFace1toCenter, parFace2toCenter, parFace3toCenter)) / 2
    EllipseMinorRadius = 0.4 * ellipseMajorRadius
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, -EllipseMinorRadius - parInsulationThickness, 0
    parInsulationDiameter = pipeDiam(parPortIndex3) + 2 * parInsulationThickness
    
    If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then
        parInsulationDiameter = flangeDiam(parPortIndex3) + 2 * parInsulationThickness
    ElseIf parInsulationDiameter < (ellipseMajorRadius + parInsulationThickness) * 2 Then
        parInsulationDiameter = (ellipseMajorRadius + parInsulationThickness) * 2
    End If
    
    Set ObjInsVertBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVertBody
    Set ObjInsVertBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
