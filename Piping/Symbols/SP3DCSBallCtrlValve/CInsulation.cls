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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Tuesday, Jan 8 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = I13AT or I13AS,
'   I13AN = I3AZ at page no 360 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy           who              change description
'   -----------         -----              ------------------
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav               Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols
'   25.Sep.2006     KKC                    TR-102869 Modified symbol code to use both Face to Face
'                                          and Face to Center Dimensions
'   2.Jul.2007      svsmylav               TR-123363: Added checks for V6.1 and V7 textual inputs.
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
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjBallInsulation As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim InsulationDia As Double
    Dim parFacetoCenter As Double

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

    Dim parHandwheelAngle   As Double
    Dim lPartDataBasis As Long
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parPortIndex2 = arrayOfInputs(18)
    parNpd2 = arrayOfInputs(19)
    parEndPreparation2 = arrayOfInputs(20)
    parScheduleThickness2 = arrayOfInputs(21)
    parEndStandard2 = arrayOfInputs(22)
    parPressureRating2 = arrayOfInputs(23)
    parFlowDirection2 = arrayOfInputs(24)
'    parFacetoCenter = arrayOfInputs(25)
'    parId1 = arrayOfInputs(26)
'    parId2 = arrayOfInputs(27)
'    parNPDUnitType = arrayOfInputs(28)
'    parNPDUnitType1 = arrayOfInputs(29)
'    parNPDUnitType2 = arrayOfInputs(30)

'   Check to ensure both V6.1 and V7 symbol definitions work
    If UBound(arrayOfInputs) >= 30 Then
        'In V7, Index 26 onwards are textual inputs
        parId1 = arrayOfInputs(26)
        parId2 = arrayOfInputs(27)
        parNPDUnitType = arrayOfInputs(28)
        parNPDUnitType1 = arrayOfInputs(29)
        parNPDUnitType2 = arrayOfInputs(30)
    Else
        'In V6.1,Index 25 onwards are textual inputs
        parId1 = arrayOfInputs(25)
        parId2 = arrayOfInputs(26)
        parNPDUnitType = arrayOfInputs(27)
        parNPDUnitType1 = arrayOfInputs(28)
        parNPDUnitType2 = arrayOfInputs(29)
    End If

    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then
       parFacetoFace = arrayOfInputs(2)
    ElseIf lPartDataBasis = 10 Then
       parFacetoCenter = arrayOfInputs(25)
       parFacetoFace = 2 * parFacetoCenter
    End If
    
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

 ' Insert your code for output 6(As a Cylinder)
    Dim InsulationDia1 As Double
    Dim InsulationDia2 As Double
    Dim StartInsulationPos       As New AutoMath.DPosition
    Dim EndInsulationPos         As New AutoMath.DPosition
    StartInsulationPos.Set -parFacetoFace / 2, 0, 0
    EndInsulationPos.Set parFacetoFace / 2, 0, 0

    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        InsulationDia1 = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    Else
        InsulationDia1 = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    End If

    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        InsulationDia2 = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    Else
        InsulationDia2 = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    If InsulationDia1 > InsulationDia2 Then
        InsulationDia = InsulationDia1
    Else
        InsulationDia = InsulationDia2
    End If
    Set ObjBallInsulation = PlaceCylinder(m_OutputColl, StartInsulationPos, EndInsulationPos, _
                                    InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBallInsulation
    Set ObjBallInsulation = Nothing
    Set StartInsulationPos = Nothing
    Set EndInsulationPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
