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
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IDCO)
'   The symbol consists of 10 outputs, seven physical, two insulation outputs and one Operation output.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
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
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    
    Dim iOutput     As Double
    Dim ObjChamberBody As Object
    Dim ObjArmCyl1 As Object
    Dim ObjArmBody2 As Object
    Dim ObjNozzCylIns As Object
    Dim ObjChamberBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parArmLength As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parArmHeight As Double
    Dim parArmHeight1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
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
    
    Dim parOperationHeight As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)  'D4
    parInstrumentHeight1 = arrayOfInputs(5) 'D5
    parInstrumentLength = arrayOfInputs(6)  'D6
    parInstrumentLength1 = arrayOfInputs(7) 'D7
    parInstrumentWidth = arrayOfInputs(8)   'D8
    parInstrumentWidth1 = arrayOfInputs(9)  'D9
    parArmLength = arrayOfInputs(10)    'D10
    parNozzleOffset = arrayOfInputs(11)   'D11
    parNozzleOffset1 = arrayOfInputs(12)  'D12
    parArmHeight = arrayOfInputs(13)    'D14
    parArmHeight1 = arrayOfInputs(14)   'D15
    parInsulationThickness = arrayOfInputs(15)
    parNPD = arrayOfInputs(16)                                'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(17)
    parScheduleThickness = arrayOfInputs(18)
    parEndStandard = arrayOfInputs(19)
    parPressureRating = arrayOfInputs(20)
    parFlowDirection = arrayOfInputs(21)
    parPortIndex1 = arrayOfInputs(22)
    parNpd1 = arrayOfInputs(23)
    parEndPreparation1 = arrayOfInputs(24)
    parScheduleThickness1 = arrayOfInputs(25)
    parEndStandard1 = arrayOfInputs(26)
    parPressureRating1 = arrayOfInputs(27)
    parFlowDirection1 = arrayOfInputs(28)
    parPortIndex2 = arrayOfInputs(29)
    parNpd2 = arrayOfInputs(30)
    parEndPreparation2 = arrayOfInputs(31)
    parScheduleThickness2 = arrayOfInputs(32)
    parEndStandard2 = arrayOfInputs(33)
    parPressureRating2 = arrayOfInputs(34)
    parFlowDirection2 = arrayOfInputs(35)
    parPortIndex3 = arrayOfInputs(36)
    parNpd3 = arrayOfInputs(37)
    parEndPreparation3 = arrayOfInputs(38)
    parScheduleThickness3 = arrayOfInputs(39)
    parEndStandard3 = arrayOfInputs(40)
    parPressureRating3 = arrayOfInputs(41)
    parFlowDirection3 = arrayOfInputs(42)
    parPortIndex4 = arrayOfInputs(43)
    parNpd4 = arrayOfInputs(44)
    parEndPreparation4 = arrayOfInputs(45)
    parScheduleThickness4 = arrayOfInputs(46)
    parEndStandard4 = arrayOfInputs(47)
    parPressureRating4 = arrayOfInputs(48)
    parFlowDirection4 = arrayOfInputs(49)
    parOperationHeight = arrayOfInputs(50)     'D13
    'Text Inputs
    parId1 = arrayOfInputs(51)
    parId2 = arrayOfInputs(52)
    parId3 = arrayOfInputs(53)
    parId4 = arrayOfInputs(54)
    parNPDUnitType = arrayOfInputs(55)
    parNPDUnitType1 = arrayOfInputs(56)
    parNPDUnitType2 = arrayOfInputs(57)
    parNPDUnitType3 = arrayOfInputs(58)
    parNPDUnitType4 = arrayOfInputs(58)
    
    iOutput = 0
'The orgin is considered to be along the flow axis and at parFace1toCenter from nozzle 1
 ' Insert your code for output 1(Chamber Body)
    stPoint.Set -parInstrumentLength, -parInstrumentHeight, parInstrumentWidth
    enPoint.Set parInstrumentLength1, parInstrumentHeight1, -parInstrumentWidth1
    Set ObjChamberBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjChamberBody
    Set ObjChamberBody = Nothing
    
 ' Insert your code for output 2(Arm Cylinder1)
'Assuming the Horizontal Arm Diameter to be 0.2 times the parInstrumentLength
    Dim arm1Diameter As Double
    arm1Diameter = 0.2 * parInstrumentLength
    stPoint.Set 0, parArmHeight, -parInstrumentWidth1
    enPoint.Set 0, parArmHeight, -parArmLength
    Set ObjArmCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, arm1Diameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmCyl1
    Set ObjArmCyl1 = Nothing
    
 ' Insert your code for output 3(Arm Body2)
'Assuming the Arm to be rectangular cross-section, with width being 0.2 times the parInstrumentLength, and Thickness being 0.1times the parInstrumentLength
    Dim arm2Width As Double
    Dim arm2Thickness As Double
    arm2Width = 0.6 * parInstrumentLength
    arm2Thickness = 0.1 * parInstrumentLength
    
    stPoint.Set arm2Width / 2, parArmHeight + arm1Diameter / 2 - parArmHeight1, -parArmLength + arm2Thickness
    enPoint.Set -arm2Width / 2, parArmHeight + arm1Diameter / 2, -parArmLength
    Set ObjArmBody2 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmBody2
    Set ObjArmBody2 = Nothing
    
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
    oNozzle.Length = parFace1toCenter - parInstrumentLength
    
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
     oNozzle.Length = parFace2toCenter - parInstrumentLength1

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace2toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
       
' Place Nozzle 3
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

' Direction of the Nozzle
    oDir.Set 0, 0, 1
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parNozzleOffset, 0, parInstrumentWidth + sptOffset(parPortIndex3) - depth(parPortIndex3)
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 4
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
    
' Direction of the Nozzle
    oDir.Set 0, 0, 1
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parNozzleOffset1, 0, parInstrumentWidth + sptOffset(parPortIndex4) - depth(parPortIndex4)
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
