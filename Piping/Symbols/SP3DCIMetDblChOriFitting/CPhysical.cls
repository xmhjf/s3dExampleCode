VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 10 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMDC2,
'   PDMS Category Name=IFRC). The symbol consists of 12  outputs, nine physical,
'   two insulation outputs and one Operation output.
'   The 4 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   12.Sep.2003    SSP/BG                  TR-48212 Corrected Arm Cross section dimentions.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
        
    Dim iOutput     As Double
    Dim ObjLHSPipe As Object
    Dim ObjRHSPipe As Object
    Dim ObjFittingBody As Object
    Dim ObjArmCylinder1 As Object
    Dim ObjArmBody2 As Object
    
    Dim parInstrumentLength      As Double
    Dim parInstrumentLength1     As Double
    Dim parInstrumentHeight      As Double
    Dim parInstrumentHeight1     As Double
    Dim parInstrumentLength2     As Double
    Dim parInstrumentLength3     As Double
    Dim parInstrumentWidth       As Double
    Dim parInstrumentWidth1      As Double
    Dim parArmLength             As Double
    Dim parNozzleOffset          As Double
    Dim parNozzleOffset1         As Double
    Dim parArmHeight             As Double
    Dim parArmHeight1            As Double
    Dim parInstrumentLength4     As Double
    Dim parInstrumentLength5     As Double
    Dim parInsulationThickness   As Double
    
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim Angle As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 26) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
    Dim parOperationHeight As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)      'D2
    parInstrumentLength1 = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength2 = arrayOfInputs(6)     'D6
    parInstrumentLength3 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parArmLength = arrayOfInputs(10)            'D10
    parNozzleOffset = arrayOfInputs(11)         'D11
    parNozzleOffset1 = arrayOfInputs(12)        'D12
    parArmHeight = arrayOfInputs(13)            'D14
    parArmHeight1 = arrayOfInputs(14)           'D15
    parInstrumentLength4 = arrayOfInputs(15)    'D16
    parInstrumentLength5 = arrayOfInputs(16)    'D17
    parInsulationThickness = arrayOfInputs(17)
    parNPD = arrayOfInputs(18)              'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(19)
    parScheduleThickness = arrayOfInputs(20)
    parEndStandard = arrayOfInputs(21)
    parPressureRating = arrayOfInputs(22)
    parFlowDirection = arrayOfInputs(23)
    parPortIndex1 = arrayOfInputs(24)
    parNpd1 = arrayOfInputs(25)
    parEndPreparation1 = arrayOfInputs(26)
    parScheduleThickness1 = arrayOfInputs(27)
    parEndStandard1 = arrayOfInputs(28)
    parPressureRating1 = arrayOfInputs(29)
    parFlowDirection1 = arrayOfInputs(30)
    parPortIndex2 = arrayOfInputs(31)
    parNpd2 = arrayOfInputs(32)
    parEndPreparation2 = arrayOfInputs(33)
    parScheduleThickness2 = arrayOfInputs(34)
    parEndStandard2 = arrayOfInputs(35)
    parPressureRating2 = arrayOfInputs(36)
    parFlowDirection2 = arrayOfInputs(37)
    parPortIndex3 = arrayOfInputs(38)
    parNpd3 = arrayOfInputs(39)
    parEndPreparation3 = arrayOfInputs(40)
    parScheduleThickness3 = arrayOfInputs(41)
    parEndStandard3 = arrayOfInputs(42)
    parPressureRating3 = arrayOfInputs(43)
    parFlowDirection3 = arrayOfInputs(44)
    parPortIndex4 = arrayOfInputs(45)
    parNpd4 = arrayOfInputs(46)
    parEndPreparation4 = arrayOfInputs(47)
    parScheduleThickness4 = arrayOfInputs(48)
    parEndStandard4 = arrayOfInputs(49)
    parPressureRating4 = arrayOfInputs(50)
    parFlowDirection4 = arrayOfInputs(51)
    parOperationHeight = arrayOfInputs(52)     'D13
'   Comment to make test with text inputs
    parId1 = arrayOfInputs(53)
    parId2 = arrayOfInputs(54)
    parId3 = arrayOfInputs(55)
    parId4 = arrayOfInputs(56)
    parNPDUnitType = arrayOfInputs(57)
    parNPDUnitType1 = arrayOfInputs(58)
    parNPDUnitType2 = arrayOfInputs(59)
    parNPDUnitType3 = arrayOfInputs(60)
    parNPDUnitType4 = arrayOfInputs(61)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
'   Center position is assumed at the intersection of the flow axis and
'   verticle axis through the point at D11 offset from port3 center.

' Place Nozzle 1 (Nozzle1)
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parInstrumentLength - parInstrumentLength4 _
                - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex2, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNPDUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex2)

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parInstrumentLength1 + parInstrumentLength5 + _
                sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    
' Place Nozzle 3 (Nozzle3)
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNPDUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex3)

' Direction of the Nozzle
    oDir.Set 0, 0, 1
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parNozzleOffset, 0, _
                parInstrumentWidth + sptOffset(parPortIndex3) - depth(parPortIndex3)
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 4 (Nozzle4)
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNPDUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex4)

' Direction of the Nozzle
    oDir.Set 0, 0, 1
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parNozzleOffset1, 0, _
            parInstrumentWidth + sptOffset(parPortIndex4) - depth(parPortIndex4)
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Insert your code for output 5(LHSPipe)
'   Assumption: Instrument Left Hand Side/Right Hand Side flanges (At distances D2 and D3 respectively)
'   diameter is assumed to be 1.2 times and  flange thickness is 0.1 times corresponding pipe diameter.
'   Point 1
    LineStringPoints(0) = -parInstrumentLength - parInstrumentLength4
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = -parInstrumentLength - parInstrumentLength4
    LineStringPoints(4) = 0
    LineStringPoints(5) = pipeDiam(parPortIndex1) / 2
    
'   Point 3
    LineStringPoints(6) = -parInstrumentLength - 0.1 * pipeDiam(parPortIndex1)
    LineStringPoints(7) = 0
    LineStringPoints(8) = pipeDiam(parPortIndex1) / 2
    
'   Point 4
    LineStringPoints(9) = -parInstrumentLength - 0.1 * pipeDiam(parPortIndex1)
    LineStringPoints(10) = 0
    LineStringPoints(11) = 1.2 * pipeDiam(parPortIndex1) / 2
    
'   Point 5
    LineStringPoints(12) = -parInstrumentLength
    LineStringPoints(13) = 0
    LineStringPoints(14) = 1.2 * pipeDiam(parPortIndex1) / 2
    
'   Point 6
    LineStringPoints(15) = -parInstrumentLength + 0.1 * pipeDiam(parPortIndex1)
    LineStringPoints(16) = 0
    LineStringPoints(17) = 1.2 * pipeDiam(parPortIndex1) / 2
    
'   Point 7
    LineStringPoints(18) = -parInstrumentLength + 0.1 * pipeDiam(parPortIndex1)
    LineStringPoints(19) = 0
    LineStringPoints(20) = pipeDiam(parPortIndex1) / 2
    
'   Point 8
    LineStringPoints(21) = -parInstrumentLength2
    LineStringPoints(22) = 0
    LineStringPoints(23) = pipeDiam(parPortIndex1) / 2
    
'   Point 9
    LineStringPoints(24) = -parInstrumentLength2
    LineStringPoints(25) = 0
    LineStringPoints(26) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set -(parInstrumentLength + parInstrumentLength4) / 2, 0, 0
    Angle = 2 * PI
    Set ObjLHSPipe = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, Angle, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSPipe
    Set ObjLHSPipe = Nothing
    
' Insert your code for output 6(RHSPipe)
'   Point 1
    LineStringPoints(0) = parInstrumentLength1 + parInstrumentLength5
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = parInstrumentLength1 + parInstrumentLength5
    LineStringPoints(4) = 0
    LineStringPoints(5) = pipeDiam(parPortIndex2) / 2
    
'   Point 3
    LineStringPoints(6) = parInstrumentLength1 + 0.1 * pipeDiam(parPortIndex2)
    LineStringPoints(7) = 0
    LineStringPoints(8) = pipeDiam(parPortIndex2) / 2
    
'   Point 4
    LineStringPoints(9) = parInstrumentLength1 + 0.1 * pipeDiam(parPortIndex2)
    LineStringPoints(10) = 0
    LineStringPoints(11) = 1.2 * pipeDiam(parPortIndex2) / 2
    
'   Point 5
    LineStringPoints(12) = parInstrumentLength1
    LineStringPoints(13) = 0
    LineStringPoints(14) = 1.2 * pipeDiam(parPortIndex2) / 2
    
'   Point 6
    LineStringPoints(15) = parInstrumentLength1 - 0.1 * pipeDiam(parPortIndex2)
    LineStringPoints(16) = 0
    LineStringPoints(17) = 1.2 * pipeDiam(parPortIndex2) / 2
    
'   Point 7
    LineStringPoints(18) = parInstrumentLength1 - 0.1 * pipeDiam(parPortIndex2)
    LineStringPoints(19) = 0
    LineStringPoints(20) = pipeDiam(parPortIndex2) / 2
    
'   Point 8
    LineStringPoints(21) = parInstrumentLength3
    LineStringPoints(22) = 0
    LineStringPoints(23) = pipeDiam(parPortIndex2) / 2
    
'   Point 9
    LineStringPoints(24) = parInstrumentLength3
    LineStringPoints(25) = 0
    LineStringPoints(26) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set (parInstrumentLength1 + parInstrumentLength5) / 2, 0, 0
    Set ObjRHSPipe = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, Angle, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSPipe
    Set ObjRHSPipe = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    
' Insert your code for output 7(FittingBody)
    stPoint.Set -parInstrumentLength2, -parInstrumentHeight, parInstrumentWidth
    enPoint.Set parInstrumentLength3, parInstrumentHeight1, -parInstrumentWidth1
    Set ObjFittingBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFittingBody
    Set ObjFittingBody = Nothing
    
' Insert your code for output 8(ArmCylinder1)
'   Assumption: Arm diameter is 0.1 times that of parArmLength
    Dim ArmDiameter As Double
    stPoint.Set 0, parArmHeight, -parInstrumentWidth1
    enPoint.Set 0, parArmHeight, -parArmLength
    ArmDiameter = parArmLength * 0.1
    Set ObjArmCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ArmDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmCylinder1
    Set ObjArmCylinder1 = Nothing
    
' Insert your code for output 9(ArmBody2)
'   Assuming the Arm to be Square cross-section
    Dim arm2Width As Double
    Dim arm2Thickness As Double
    Dim armSize As Double
    
    armSize = 0.1 * parArmHeight1
    If armSize < 0.025 Then
        armSize = 0.025
    ElseIf armSize > 0.05 Then
        armSize = 0.05
    End If
    
    arm2Width = armSize
    arm2Thickness = armSize
    
    stPoint.Set arm2Width / 2, _
                        parArmHeight + ArmDiameter / 2 - parArmHeight1, _
                        -(parArmLength - arm2Thickness)

    enPoint.Set -arm2Width / 2, _
                        parArmHeight + ArmDiameter / 2, _
                        -parArmLength
    Set ObjArmBody2 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmBody2
    Set ObjArmBody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

