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
'   Author:          ACM
'   Creation Date:  Tuesday, Jan 7 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVCU)
'  The symbol consists of 13 outputs; Ten physical outputs, two Insulation Output and one Operaration output.
'  The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
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
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjStemBody As Object
    Dim ObjActuatorBody As Object
    Dim ObjMotor1Body As Object
    Dim ObjMotor2Body As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parMotor1Diameter As Double
    Dim parMotor2Diameter As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Length As Double
    Dim parMotor1Length As Double
    Dim parActuatorOffset As Double
    Dim parActuatorHeight1 As Double
    Dim parHandWheelDiameter As Double
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
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)      'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parActuatorHeight = arrayOfInputs(5)      'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorLength = arrayOfInputs(7)      'D7
    parActuatorLength1 = arrayOfInputs(8)     'D8
    parMotor1Diameter = arrayOfInputs(9)      'D9
    parMotor2Diameter = arrayOfInputs(10)     'D10
    parMotor2Offset = arrayOfInputs(11)   'D11
    parMotor1Offset = arrayOfInputs(12)   'D12
    parMotor2Length = arrayOfInputs(13)   'D13
    parMotor1Length = arrayOfInputs(14)   'D14
    parActuatorOffset = arrayOfInputs(15)     'D15
    parActuatorHeight1 = arrayOfInputs(16)     'D16
    parHandWheelDiameter = arrayOfInputs(17)      'D17
    parInsulationThickness = arrayOfInputs(18)
        parNPD = arrayOfInputs(19)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(20)
    parScheduleThickness = arrayOfInputs(21)
    parEndStandard = arrayOfInputs(22)
    parPressureRating = arrayOfInputs(23)
    parFlowDirection = arrayOfInputs(24)
    parPortIndex1 = arrayOfInputs(25)
    parNpd1 = arrayOfInputs(26)
    parEndPreparation1 = arrayOfInputs(27)
    parScheduleThickness1 = arrayOfInputs(28)
    parEndStandard1 = arrayOfInputs(29)
    parPressureRating1 = arrayOfInputs(30)
    parFlowDirection1 = arrayOfInputs(31)
    parPortIndex2 = arrayOfInputs(32)
    parNpd2 = arrayOfInputs(33)
    parEndPreparation2 = arrayOfInputs(34)
    parScheduleThickness2 = arrayOfInputs(35)
    parEndStandard2 = arrayOfInputs(36)
    parPressureRating2 = arrayOfInputs(37)
    parFlowDirection2 = arrayOfInputs(38)
    parOperationHeight = arrayOfInputs(39)     'D18
    
    parId1 = arrayOfInputs(40)
    parId2 = arrayOfInputs(41)
    parNPDUnitType = arrayOfInputs(42)
    parNPDUnitType1 = arrayOfInputs(43)
    parNPDUnitType2 = arrayOfInputs(44)
    
    iOutput = 0
    
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

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
    oPlacePoint.Set -parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset(parPortIndex2) + depth(parPortIndex2), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
 
' Insert your code for output 3(Body Cone1 of Valve)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim HandWheelThk As Double
    Dim StemDia As Double
    HandWheelThk = 0.05
    StemDia = 0.05
    
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 4(Body Cone2 of Valve)
    stPoint.Set 0, -parFace2toCenter + flangeThick(parPortIndex2), 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
'    Insert your code for output 5(Valve Center Cylinder)
    Dim CylDia As Double
    Dim ValCenCylBottLength As Double
'   Valve Center Cylinder Diameter is taken to be 0.7 times the Nozzle2 diameter
    CylDia = 0.7 * pipeDiam(parPortIndex1)
'   Valve Center Cylinder Bottom Length is taken to be 0.25 times the parFace2toCenter
    ValCenCylBottLength = parFace2toCenter * 0.25
    stPoint.Set 0, parCylHeight, 0
    enPoint.Set 0, -ValCenCylBottLength, 0
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
 ' Insert your code for output 6(Stem Body)
     stPoint.Set 0, parCylHeight, 0
    enPoint.Set 0, parActuatorHeight1, 0
    Set ObjStemBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemBody
    Set ObjStemBody = Nothing
    
 ' Insert your code for output 7(Actuator Body)
     stPoint.Set -parActuatorLength, parActuatorHeight, parActuatorOffset
    enPoint.Set parActuatorLength1, parActuatorHeight, parActuatorOffset
    Set ObjActuatorBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    
 ' Insert your code for output 8(Motor1 Body)
    stPoint.Set -parMotor1Offset, parActuatorHeight, parActuatorOffset
    enPoint.Set -parMotor1Offset, parActuatorHeight, parMotor1Length + parActuatorOffset
    Set ObjMotor1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1Body
    Set ObjMotor1Body = Nothing
    
 ' Insert your code for output 9(Motor2 Body)
    stPoint.Set parMotor2Offset, parActuatorHeight, parActuatorOffset
    enPoint.Set parMotor2Offset, parActuatorHeight, parMotor2Length + parActuatorOffset
    Set ObjMotor2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor2Diameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2Body
    Set ObjMotor2Body = Nothing
    
 ' Insert your code for output 10(HandWheel)
 '   Place Torus
    Dim ObjHandwheel    As IngrGeom3D.Torus3d
    Dim TorSecRad As Double
    Dim TorPriRad As Double
    Dim TorOrigin       As New AutoMath.DPosition
    Dim TorAxis         As New AutoMath.DVector
    TorSecRad = HandWheelThk / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    TorOrigin.Set 0, parActuatorHeight1 - HandWheelThk / 2, 0
    TorAxis.Set 0, 1, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

