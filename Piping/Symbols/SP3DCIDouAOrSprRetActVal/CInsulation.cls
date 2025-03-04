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
'   Creation Date:  Monday 13 January, 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA3A,PDMS
'   Category Name=IVCK). Physical Aspect consist of Two cones, a bottom Cylinder, a stem, Two boxes ,
'   a cylinder and Two Nozzles. Insulation aspect consist of a single cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
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
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    Dim iOutput     As Double
    Dim ObjValveBodyIns  As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
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
    
    Dim parOperationHeight As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)       'D4
    parActuatorHeight = arrayOfInputs(5)        'D5
    parActuatorHeight1 = arrayOfInputs(6)       'D6
    parActuatorLength = arrayOfInputs(7)         'D7
    parActuatorLength1 = arrayOfInputs(8)        'D8
    parActuatorWidth = arrayOfInputs(9)        'D9
    parActuatorWidth1 = arrayOfInputs(10)       'D10
    parActuatorHeight2 = arrayOfInputs(11)       'D11
    parActuatorWidth2 = arrayOfInputs(12)       'D12
    parActuatorWidth3 = arrayOfInputs(13)      'D13
    parActuatorLength2 = arrayOfInputs(14)       'D14
    parActuatorLength3 = arrayOfInputs(15)       'D15
    parActuatorOffset = arrayOfInputs(16)       'D16
    parActuatorCylDiameter = arrayOfInputs(17)  'D17
    parActuatorCylLength = arrayOfInputs(18)    'D18
    parActuatorHeight3 = arrayOfInputs(19)      'D19
    parInsulationThickness = arrayOfInputs(20)
    parNPD = arrayOfInputs(21)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(22)
    parScheduleThickness = arrayOfInputs(23)
    parEndStandard = arrayOfInputs(24)
    parPressureRating = arrayOfInputs(25)
    parFlowDirection = arrayOfInputs(26)
    parPortIndex1 = arrayOfInputs(27)
    parNpd1 = arrayOfInputs(28)
    parEndPreparation1 = arrayOfInputs(29)
    parScheduleThickness1 = arrayOfInputs(30)
    parEndStandard1 = arrayOfInputs(31)
    parPressureRating1 = arrayOfInputs(32)
    parFlowDirection1 = arrayOfInputs(33)
    parPortIndex2 = arrayOfInputs(34)
    parNpd2 = arrayOfInputs(35)
    parEndPreparation2 = arrayOfInputs(36)
    parScheduleThickness2 = arrayOfInputs(37)
    parEndStandard2 = arrayOfInputs(38)
    parPressureRating2 = arrayOfInputs(39)
    parFlowDirection2 = arrayOfInputs(40)
    parOperationHeight = arrayOfInputs(41)     'D20
    
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parNPDUnitType = arrayOfInputs(44)
    parNPDUnitType1 = arrayOfInputs(45)
    parNPDUnitType2 = arrayOfInputs(46)
    
    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 8(Valve Body Insulation)
    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    CylDia = parFace1toCenter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
         
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
'maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam(parPortIndex1) / 2, pipeDiam(parPortIndex2) / 2, _
                    flangeDiam(parPortIndex1) / 2, flangeDiam(parPortIndex2) / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
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

