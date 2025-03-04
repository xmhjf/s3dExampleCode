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
'   Author:          SSP
'   Creation Date:  Tuesday, Nov 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIV,
'   PDMS Category Name=IEEF). Physical aspect consist of four nozzles and one revolved linestring.
'   Insulation aspect consist of single entity which is a cylinder.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjInsBody As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentDiameter2 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    
    Dim parNpd                     As Double
    Dim parEndPreparation          As Long
    Dim parScheduleThickness       As Long
    Dim parEndStandard             As Long
    Dim parPressureRating          As Long
    Dim parFlowDirection           As DistribFlow
    
    Dim parPortIndex1              As String
    Dim parNpd1                    As Double
    Dim parEndPreparation1         As Long
    Dim parScheduleThickness1      As Long
    Dim parEndStandard1            As Long
    Dim parPressureRating1         As Long
    Dim parFlowDirection1          As DistribFlow
    Dim parPortIndex2              As String
    Dim parNpd2                    As Double
    Dim parEndPreparation2         As Long
    Dim parScheduleThickness2      As Long
    Dim parEndStandard2            As Long
    Dim parPressureRating2         As Long
    Dim parFlowDirection2          As DistribFlow
    Dim parPortIndex3              As String
    Dim parNpd3                    As Double
    Dim parEndPreparation3         As Long
    Dim parScheduleThickness3      As Long
    Dim parEndStandard3            As Long
    Dim parPressureRating3         As Long
    Dim parFlowDirection3          As DistribFlow
    Dim parPortIndex4              As String
    Dim parNpd4                    As Double
    Dim parEndPreparation4         As Long
    Dim parScheduleThickness4      As Long
    Dim parEndStandard4            As Long
    Dim parPressureRating4         As Long
    Dim parFlowDirection4          As DistribFlow
    Dim parId1                     As String
    Dim parId2                     As String
    Dim parId3                     As String
    Dim parId4                     As String
    Dim parNPDUnitType             As String
    Dim parNpdUnitType1            As String
    Dim parNpdUnitType2            As String
    Dim parNpdUnitType3            As String
    Dim parNpdUnitType4            As String
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)            'D2 is equal to D3 + D4 + D5, so NOT used.
    parInstrumentLength = arrayOfInputs(3)      'D3
    parInstrumentLength1 = arrayOfInputs(4)     'D4
    parInstrumentLength2 = arrayOfInputs(5)     'D5
    parInstrumentDiameter = arrayOfInputs(6)    'D6
    parInstrumentDiameter1 = arrayOfInputs(7)   'D7
    parInstrumentDiameter2 = arrayOfInputs(8)   'D8
    parPortRotation1 = arrayOfInputs(9)                 'D9
    parPortRotation2 = arrayOfInputs(10)               'D10
    parNpd = arrayOfInputs(11)
    parEndPreparation = arrayOfInputs(12)
    parScheduleThickness = arrayOfInputs(13)
    parEndStandard = arrayOfInputs(14)
    parPressureRating = arrayOfInputs(15)
    parFlowDirection = arrayOfInputs(16)
    parPortIndex1 = arrayOfInputs(17)
    parNpd1 = arrayOfInputs(18)
    parEndPreparation1 = arrayOfInputs(19)
    parScheduleThickness1 = arrayOfInputs(20)
    parEndStandard1 = arrayOfInputs(21)
    parPressureRating1 = arrayOfInputs(22)
    parFlowDirection1 = arrayOfInputs(23)
    parPortIndex2 = arrayOfInputs(24)
    parNpd2 = arrayOfInputs(25)
    parEndPreparation2 = arrayOfInputs(26)
    parScheduleThickness2 = arrayOfInputs(27)
    parEndStandard2 = arrayOfInputs(28)
    parPressureRating2 = arrayOfInputs(29)
    parFlowDirection2 = arrayOfInputs(30)
    parPortIndex3 = arrayOfInputs(31)
    parNpd3 = arrayOfInputs(32)
    parEndPreparation3 = arrayOfInputs(33)
    parScheduleThickness3 = arrayOfInputs(34)
    parEndStandard3 = arrayOfInputs(35)
    parPressureRating3 = arrayOfInputs(36)
    parFlowDirection3 = arrayOfInputs(37)
    parPortIndex4 = arrayOfInputs(38)
    parNpd4 = arrayOfInputs(39)
    parEndPreparation4 = arrayOfInputs(40)
    parScheduleThickness4 = arrayOfInputs(41)
    parEndStandard4 = arrayOfInputs(42)
    parPressureRating4 = arrayOfInputs(43)
    parFlowDirection4 = arrayOfInputs(44)
    
    parInsulationThickness = arrayOfInputs(45)
    
    'Text Inputs
    parId1 = arrayOfInputs(46)
    parId2 = arrayOfInputs(47)
    parId3 = arrayOfInputs(48)
    parId4 = arrayOfInputs(49)
    parNPDUnitType = arrayOfInputs(50)
    parNpdUnitType1 = arrayOfInputs(51)
    parNpdUnitType2 = arrayOfInputs(52)
    parNpdUnitType3 = arrayOfInputs(53)
    parNpdUnitType4 = arrayOfInputs(54)

    iOutput = 0

' Insert your code for output 1(Insulation for Body)
    stPoint.Set -parInstrumentLength1 / 2 - parInstrumentLength - parInsulationThickness, 0, 0
    enPoint.Set parInstrumentLength1 / 2 + parInstrumentLength2 + parInsulationThickness, 0, 0
    parInsulationDiameter = parInstrumentDiameter2 + 2 * parInsulationThickness
    If parInstrumentDiameter > parInstrumentDiameter2 Then
        parInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    End If
    Set ObjInsBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    Set ObjInsBody = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
