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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on (SP1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   8-Nov-07        KKC/svsmylav     DI-130618: Changed the parameter names of parPositionerOrientation and parHandWheelOrientation
'                                          to parPositionerRotAngle and parHandWheelRotAngle respectively
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
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    If CmpDblLessThan(MaxValue, E) Then MaxValue = E
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
    Dim ObjValveBodyIns   As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parActuatorCyl1Diameter As Double
    Dim parActuatorCyl1Length As Double
    Dim parHandWheelLength As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parPositionerLength As Double
    Dim parPositionerHeight As Double
    Dim parPostionerOffset As Double
    Dim parPostionerOffset1 As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parPositionerRotAngle As Double
    Dim parHandWheelRotAngle As Double
    
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

    Dim StemDia As Double
    Dim parOperationHeight As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'V1
    parFacetoFace = arrayOfInputs(3)          'V2
    parCylHeight = arrayOfInputs(4)           'V3
    parActuatorHeight = arrayOfInputs(5)      'A6
    parActuatorHeight1 = arrayOfInputs(6)     'A5
    parActuatorWidth = arrayOfInputs(7)       'A7
    parActuatorWidth1 = arrayOfInputs(8)      'A11
    parActuatorLength = arrayOfInputs(9)      'A3
    parActuatorLength1 = arrayOfInputs(10)    'A4
    parActuatorHeight2 = arrayOfInputs(11)    'A9
    parActuatorLength2 = arrayOfInputs(12)    'A10
    parActuatorWidth2 = arrayOfInputs(13)     'A8
    parActuatorOffset = arrayOfInputs(14)     'C1
    parActuatorCylDiameter = arrayOfInputs(15) 'C11
    parActuatorCylLength = arrayOfInputs(16)      'C12
    parActuatorHeight3 = arrayOfInputs(17)     'A2
    parActuatorCyl1Diameter = arrayOfInputs(18)    'C21
    parActuatorCyl1Length = arrayOfInputs(19)      'C22
    parHandWheelLength = arrayOfInputs(20)         'H2
    parHandWheelDiameter = arrayOfInputs(21)         'H3
    parHandWheelOffset = arrayOfInputs(22)          'H4
    parPositionerLength = arrayOfInputs(23)         'P2
    parPositionerHeight = arrayOfInputs(24)         'P4
    parPostionerOffset = arrayOfInputs(25)         'P3
    parPostionerOffset1 = arrayOfInputs(26)         'P5
    parInsulationThickness = arrayOfInputs(27)
    parRotation = arrayOfInputs(28)
    parNPD = arrayOfInputs(29)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(30)
    parScheduleThickness = arrayOfInputs(31)
    parEndStandard = arrayOfInputs(32)
    parPressureRating = arrayOfInputs(33)
    parFlowDirection = arrayOfInputs(34)
    parPortIndex1 = arrayOfInputs(35)
    parNpd1 = arrayOfInputs(36)
    parEndPreparation1 = arrayOfInputs(37)
    parScheduleThickness1 = arrayOfInputs(38)
    parEndStandard1 = arrayOfInputs(39)
    parPressureRating1 = arrayOfInputs(40)
    parFlowDirection1 = arrayOfInputs(41)
    parPortIndex2 = arrayOfInputs(42)
    parNpd2 = arrayOfInputs(43)
    parEndPreparation2 = arrayOfInputs(44)
    parScheduleThickness2 = arrayOfInputs(45)
    parEndStandard2 = arrayOfInputs(46)
    parPressureRating2 = arrayOfInputs(47)
    parFlowDirection2 = arrayOfInputs(48)
    parOperationHeight = arrayOfInputs(49)     'I1
    parPositionerRotAngle = arrayOfInputs(50)       'P1
    parHandWheelRotAngle = arrayOfInputs(51)        'H1
'   Text Inputs
    parId1 = arrayOfInputs(52)
    parId2 = arrayOfInputs(53)
    parNPDUnitType = arrayOfInputs(54)
    parNPDUnitType1 = arrayOfInputs(55)
    parNPDUnitType2 = arrayOfInputs(56)
    

    iOutput = 0

' Insert your code for output 6(Bottom Cylinder)
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
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
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
    
End Sub
