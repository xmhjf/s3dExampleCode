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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
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
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.May.2008     KKC  CR-135970  Provide ability to rotate actuator for on-the-fly control valves
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
    Dim ObjStem1 As Object
    Dim ObjMotor1 As Object
    Dim ObjMotor2 As Object
    Dim ObjMotor3 As Object
    Dim ObjStem2 As Object
    Dim ObjHandwheel As Object
    
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

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim TorOrigin As New AutoMath.DPosition
    
    Dim StemDia As Double
    Dim HandWheelCSDia As Double
 
    Dim parOperationHeight As Double
    Dim parAngle As Double
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
    
    If UBound(arrayOfInputs) > 42 Then
        parAngle = arrayOfInputs(43)
    End If
   
    iOutput = 0
    
'   Stem diameter and Hand wheel minimum Thickness are taken as 0.05
    StemDia = 0.05
    HandWheelCSDia = 0.05

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

'   Set the output
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
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' Insert your code for output 3(Valve Cone1)
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 4(Body Cone2 of Valve)
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 5(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 '  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
' Insert your code for output 6(Stem 1)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight, 0
    Set ObjStem1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem1
    Set ObjStem1 = Nothing
    
' Insert your code for output 7(Motor 1)
    stPoint.Set -parActuatorLength, _
                parActuatorHeight - parActuatorOffset, _
                parActuatorOffset1
    enPoint.Set 0, parActuatorHeight - parActuatorOffset, parActuatorOffset1
    Set ObjMotor1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
    
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    
    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parAngle, oAxisVect
    ObjMotor1.Transform oTransMatrix
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1
    Set ObjMotor1 = Nothing
    
' Insert your code for output 8(Motor 2)
    stPoint.Set 0, parActuatorHeight, parActuatorOffset1
    enPoint.Set parActuatorLength1, parActuatorHeight, parActuatorOffset1
    Set ObjMotor2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor2Diameter, True)
 
    ObjMotor2.Transform oTransMatrix
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2
    Set ObjMotor2 = Nothing
    
' Insert your code for output 9(Motor 3)
    stPoint.Set 0, parActuatorHeight, _
                -(parMotor2Diameter / 2 - parActuatorOffset1)
    enPoint.Set 0, parActuatorHeight, _
                parMotor3EndtoCenter
    Set ObjMotor3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor3Diameter, True)
 
    ObjMotor3.Transform oTransMatrix
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor3
    Set ObjMotor3 = Nothing
    
' Insert your code for output 10(Stem 2)
    stPoint.Set 0, parActuatorHeight, 0
    enPoint.Set 0, parActuatorHeight1, 0
    Set ObjStem2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem2
    Set ObjStem2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 11(Handwheel)
    TorOrigin.Set 0, parActuatorHeight1 - HandWheelCSDia / 2, 0
    axisVect.Set 0, 1, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, axisVect, _
                        parHandWheelDiameter / 2 - HandWheelCSDia / 2, _
                        HandWheelCSDia / 2)
                        
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set axisVect = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
