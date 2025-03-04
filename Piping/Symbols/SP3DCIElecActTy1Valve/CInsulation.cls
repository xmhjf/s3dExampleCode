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
'   Author:          svsmylav
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IEAM1,
'   PDMS Category Name=IVCM). The symbol consists of Twelve outputs, Eleven physical
'   outputs and One Insulation Output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function

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
    Dim ObjValveBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight  As Double
    Dim parActuatorHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorOffset As Double
    Dim parMotor1Diameter As Double
    Dim parMotor2Diameter As Double
    Dim parActuatorLength1 As Double
    Dim parMotor3Diameter As Double
    Dim parActuatorHeight1 As Double
    Dim parHandWheelDiameter As Double
    Dim parActuatorOffset1 As Double
    Dim parMotor3EndtoCenter As Double
    Dim parInsulationThickness As Double
    
'   Dynamic Port Input parameters variable declaration
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
    
 
Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'D2
    parFacetoFace = arrayOfInputs(3)     'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parActuatorHeight = arrayOfInputs(5)     'D5
    parActuatorLength = arrayOfInputs(6)     'D6
    parActuatorOffset = arrayOfInputs(7)     'D7
    parMotor1Diameter = arrayOfInputs(8)     'D8
    parMotor2Diameter = arrayOfInputs(9)     'D9
    parActuatorLength1 = arrayOfInputs(10)   'D10
    parMotor3Diameter = arrayOfInputs(11)    'D11
    parActuatorHeight1 = arrayOfInputs(12)    'D12
    parHandWheelDiameter = arrayOfInputs(13)     'D13
    parActuatorOffset1 = arrayOfInputs(14)   'D14
    parMotor3EndtoCenter = arrayOfInputs(15)     'D15
    parInsulationThickness = arrayOfInputs(16)
    parNPD = arrayOfInputs(17)
    parEndPreparation = arrayOfInputs(18)
    parScheduleThickness = arrayOfInputs(19)
    parEndStandard = arrayOfInputs(20)
    parPressureRating = arrayOfInputs(21)
    parFlowDirection = arrayOfInputs(22)
    parPortIndex1 = arrayOfInputs(23)
    parNpd1 = arrayOfInputs(24)
    parEndPreparation1 = arrayOfInputs(25)
    parScheduleThickness1 = arrayOfInputs(26)
    parEndStandard1 = arrayOfInputs(27)
    parPressureRating1 = arrayOfInputs(28)
    parFlowDirection1 = arrayOfInputs(29)
    parPortIndex2 = arrayOfInputs(30)
    parNpd2 = arrayOfInputs(31)
    parEndPreparation2 = arrayOfInputs(32)
    parScheduleThickness2 = arrayOfInputs(33)
    parEndStandard2 = arrayOfInputs(34)
    parPressureRating2 = arrayOfInputs(35)
    parFlowDirection2 = arrayOfInputs(36)
    parOperationHeight = arrayOfInputs(37)
    parId1 = arrayOfInputs(38)
    parId2 = arrayOfInputs(39)
    parNPDUnitType = arrayOfInputs(40)
    parNPDUnitType1 = arrayOfInputs(41)
    parNPDUnitType2 = arrayOfInputs(42)
    
    iOutput = 0

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' Insert your code for output 12(Insulation for Valve Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double

    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    CylDia = parFace1toCenter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
'   maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam(parPortIndex1) / 2, pipeDiam(parPortIndex2) / 2, _
                        flangeDiam(parPortIndex1) / 2, flangeDiam(parPortIndex2) / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
