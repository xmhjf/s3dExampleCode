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
'   Author:          SSP
'   Creation Date:  Friday, Jan 3 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVCT).
'   The symbol consists of 15 outputs, Thirteen physical outputs, one Insulation Output and one Operation output.
'   The nozzles are fully Parametric.
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
    Dim ObjMotor3Body As Object
    Dim ObjMotor4Body As Object
    Dim ObjHandWheel As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylHeight As Double
    Dim parMotor4Offset As Double
    Dim parMotor4Length As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parHandWheelDiameter As Double
    Dim parMotor1Diameter As Double
    Dim parMotor4Diameter As Double
    Dim parMotor3Length As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Length As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parMotor3toCenter As Double
    Dim parActuatorHeight3 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim BaseRadius As Double
    Dim TopRadius As Double
    Dim Diameter As Double
    
    Dim StemDia As Double
    Dim HandWheelCSDia As Double
    
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
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parMotor4Offset = arrayOfInputs(5)  'D5
    parMotor4Length = arrayOfInputs(6)  'D6
    parActuatorHeight = arrayOfInputs(7)    'D7
    parActuatorHeight1 = arrayOfInputs(8)   'D8
    parActuatorHeight2 = arrayOfInputs(9)   'D9
    parHandWheelDiameter = arrayOfInputs(10)    'D10
    parMotor1Diameter = arrayOfInputs(11)   'D11
    parMotor4Diameter = arrayOfInputs(12)   'D12
    parMotor3Length = arrayOfInputs(13) 'D13
    parMotor1Offset = arrayOfInputs(14) 'D14
    parMotor2Offset = arrayOfInputs(15) 'D15
    parMotor1Length = arrayOfInputs(16) 'D16
    parActuatorWidth = arrayOfInputs(17)    'D17
    parActuatorWidth1 = arrayOfInputs(18)   'D18
    parMotor3toCenter = arrayOfInputs(19)   'D19
    parActuatorHeight3 = arrayOfInputs(20)  'D20
    parInsulationThickness = arrayOfInputs(21)
    parNPD = arrayOfInputs(22)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(23)
    parScheduleThickness = arrayOfInputs(24)
    parEndStandard = arrayOfInputs(25)
    parPressureRating = arrayOfInputs(26)
    parFlowDirection = arrayOfInputs(27)
    parPortIndex1 = arrayOfInputs(28)
    parNpd1 = arrayOfInputs(29)
    parEndPreparation1 = arrayOfInputs(30)
    parScheduleThickness1 = arrayOfInputs(31)
    parEndStandard1 = arrayOfInputs(32)
    parPressureRating1 = arrayOfInputs(33)
    parFlowDirection1 = arrayOfInputs(34)
    parPortIndex2 = arrayOfInputs(35)
    parNpd2 = arrayOfInputs(36)
    parEndPreparation2 = arrayOfInputs(37)
    parScheduleThickness2 = arrayOfInputs(38)
    parEndStandard2 = arrayOfInputs(39)
    parPressureRating2 = arrayOfInputs(40)
    parFlowDirection2 = arrayOfInputs(41)
    parOperationHeight = arrayOfInputs(42)     'D21
    'Text Inputs
    parId1 = arrayOfInputs(43)
    parId2 = arrayOfInputs(44)
    parNPDUnitType = arrayOfInputs(45)
    parNPDUnitType1 = arrayOfInputs(46)
    parNPDUnitType2 = arrayOfInputs(47)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
 ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

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
    
' Insert your code for output 3(Body Cone1
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
' Insert your code for output 4(Body Cone2 of Valve)
    stPoint.Set 0, -parFace2toCenter + flangeThick(parPortIndex2), 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
' Insert your code for output 5(Valve Center Cylinder)
    Dim CylDia As Double
    Dim ValCenCylBottLength As Double
    
'   Valve Center Cylinder Diameter is taken to be 0.7 times the Nozzle2 diameter
    CylDia = 0.7 * pipeDiam(parPortIndex2)
'   Valve Center Cylinder Bottom Length is taken to be 0.25 times the parFace2toCenter
    ValCenCylBottLength = parFace2toCenter * 0.25
    stPoint.Set 0, parCylHeight, 0
    enPoint.Set 0, -ValCenCylBottLength, 0
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
'   Stem diameter and Hand wheel minimum Thickness are taken as 0.25 times cylDia
    StemDia = CylDia * 0.25
    HandWheelCSDia = CylDia * 0.25
' Insert your code for output 6(StemBody)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight2, 0
    Set ObjStemBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemBody
    Set ObjStemBody = Nothing
    
' Insert your code for output 7(ActuatorBody)
    stPoint.Set -parMotor1Offset - parMotor1Diameter / 2, _
                        parActuatorHeight3 - parMotor4Offset - parActuatorHeight, _
                        parActuatorWidth
    enPoint.Set parMotor2Offset + parMotor1Diameter / 2, _
                        parActuatorHeight3 - parMotor4Offset + parActuatorHeight1, _
                        -parActuatorWidth1
    Set ObjActuatorBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    
' Insert your code for output 8(Motor1Body)
    stPoint.Set -parMotor1Offset, parActuatorHeight3 - parMotor4Offset, parActuatorWidth
    enPoint.Set -parMotor1Offset, parActuatorHeight3 - parMotor4Offset, parMotor1Length
    Set ObjMotor1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor1Body
    Set ObjMotor1Body = Nothing
    
' Insert your code for output 9(Motor2Body)
    stPoint.Set parMotor2Offset, parActuatorHeight3 - parMotor4Offset, parActuatorWidth
    enPoint.Set parMotor2Offset, parActuatorHeight3 - parMotor4Offset, parMotor1Length
    Set ObjMotor2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor2Body
    Set ObjMotor2Body = Nothing
    
' Insert your code for output 10(Motor3Body)
    stPoint.Set -parMotor1Offset, parActuatorHeight3 - parMotor4Offset, parMotor3toCenter
    enPoint.Set -parMotor1Offset - parMotor3Length, parActuatorHeight3 - parMotor4Offset, _
                parMotor3toCenter
    Set ObjMotor3Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor1Diameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor3Body
    Set ObjMotor3Body = Nothing
    
' Insert your code for output 11(Motor4Body)
    stPoint.Set -parMotor1Offset - parMotor1Diameter / 2, parActuatorHeight3, 0
    enPoint.Set -parMotor4Length, parActuatorHeight3, 0
    Set ObjMotor4Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMotor4Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotor4Body
    Set ObjMotor4Body = Nothing
    
' Insert your code for output 12(HandWheel)
'   Place Torus
    Dim TorSecRad As Double
    Dim TorPriRad As Double
    Dim TorOrigin       As New AutoMath.DPosition
    Dim TorAxis         As New AutoMath.DVector
    TorSecRad = HandWheelCSDia / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    TorOrigin.Set 0, parActuatorHeight2 - HandWheelCSDia / 2, 0
    TorAxis.Set 0, 1, 0
    Set ObjHandWheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


