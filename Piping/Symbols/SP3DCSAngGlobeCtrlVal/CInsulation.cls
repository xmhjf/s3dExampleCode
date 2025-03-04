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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 24 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = I3AT or I3AS,
'   SN = I3AZ at page no 354 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object
    Dim parFacetoCenter As Double
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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parNPD = arrayOfInputs(4)
    parEndPreparation = arrayOfInputs(5)
    parScheduleThickness = arrayOfInputs(6)
    parEndStandard = arrayOfInputs(7)
    parPressureRating = arrayOfInputs(8)
    parFlowDirection = arrayOfInputs(9)
    parPortIndex1 = arrayOfInputs(10)
    parNpd1 = arrayOfInputs(11)
    parEndPreparation1 = arrayOfInputs(12)
    parScheduleThickness1 = arrayOfInputs(13)
    parEndStandard1 = arrayOfInputs(14)
    parPressureRating1 = arrayOfInputs(15)
    parFlowDirection1 = arrayOfInputs(16)
    parPortIndex2 = arrayOfInputs(17)
    parNpd2 = arrayOfInputs(18)
    parEndPreparation2 = arrayOfInputs(19)
    parScheduleThickness2 = arrayOfInputs(20)
    parEndStandard2 = arrayOfInputs(21)
    parPressureRating2 = arrayOfInputs(22)
    parFlowDirection2 = arrayOfInputs(23)
    parId1 = arrayOfInputs(24)
    parId2 = arrayOfInputs(25)
    parNPDUnitType = arrayOfInputs(26)
    parNPDUnitType1 = arrayOfInputs(27)
    parNPDUnitType2 = arrayOfInputs(28)
    
    iOutput = 0
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

 ' Insert your code for output 3(Insulation Cylinder)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDiameter As Double
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        InsulationDiameter = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    Else
        InsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    
' Insert your code for output 3(Insulation Cylinder)
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        InsulationDiameter = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    Else
        InsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parFacetoCenter, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
