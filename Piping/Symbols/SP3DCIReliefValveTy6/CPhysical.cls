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
'   Author:          MS
'   Creation Date:  Thursday, 26 Dec 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT6, PDMS Category Name=IVPS)
'   The symbol consists of nine outputs, eight physical and one insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
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
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjValveBody As Object
    Dim ObjCyl1Body As Object
    Dim ObjCyl2Body As Object
    Dim ObjCyl3Body As Object
    Dim ObjCyl4Body As Object
    
    Dim parFace1toCenter   As Double
    Dim parFace3toCenter   As Double
    Dim parInstrumentHeight    As Double
    Dim parInstrumentHeight1   As Double
    Dim parInstrumentLength    As Double
    Dim parInstrumentLength1   As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1    As Double
    Dim parInstrumentRadius    As Double
    Dim parInstrumentHeight2   As Double
    Dim parInstrumentHeight3   As Double
    Dim parInstrumentRadius1   As Double
    Dim parInstrumentHeight4   As Double
    Dim parInstrumentHeight5   As Double
    Dim parInstrumentRadius2   As Double
    Dim parInstrumentHeight6   As Double
    Dim parInstrumentHeight7   As Double
    Dim parInstrumentRadius3   As Double
    Dim parInstrumentHeight8   As Double
    Dim parInstrumentHeight9   As Double
    Dim parCylOffset   As Double
    Dim parCylOffset1  As Double
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace3toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentRadius = arrayOfInputs(10)     'D10
    parInstrumentHeight2 = arrayOfInputs(11)    'D11
    parInstrumentHeight3 = arrayOfInputs(12)    'D12
    parInstrumentRadius1 = arrayOfInputs(13)    'D13
    parInstrumentHeight4 = arrayOfInputs(14)    'D14
    parInstrumentHeight5 = arrayOfInputs(15)    'D15
    parInstrumentRadius2 = arrayOfInputs(16)    'D16
    parInstrumentHeight6 = arrayOfInputs(17)    'D17
    parInstrumentHeight7 = arrayOfInputs(18)    'D18
    parInstrumentRadius3 = arrayOfInputs(19)    'D19
    parInstrumentHeight8 = arrayOfInputs(20)    'D20
    parInstrumentHeight9 = arrayOfInputs(21)    'D21
    parCylOffset = arrayOfInputs(22)    'D22
    parCylOffset1 = arrayOfInputs(23)   'D23
    parInsulationThickness = arrayOfInputs(24)
    parNPD = arrayOfInputs(25)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(26)
    parScheduleThickness = arrayOfInputs(27)
    parEndStandard = arrayOfInputs(28)
    parPressureRating = arrayOfInputs(29)
    parFlowDirection = arrayOfInputs(30)
    parPortIndex1 = arrayOfInputs(31)
    parNpd1 = arrayOfInputs(32)
    parEndPreparation1 = arrayOfInputs(33)
    parScheduleThickness1 = arrayOfInputs(34)
    parEndStandard1 = arrayOfInputs(35)
    parPressureRating1 = arrayOfInputs(36)
    parFlowDirection1 = arrayOfInputs(37)
    parPortIndex2 = arrayOfInputs(38)
    parNpd2 = arrayOfInputs(39)
    parEndPreparation2 = arrayOfInputs(40)
    parScheduleThickness2 = arrayOfInputs(41)
    parEndStandard2 = arrayOfInputs(42)
    parPressureRating2 = arrayOfInputs(43)
    parFlowDirection2 = arrayOfInputs(44)
    parPortIndex3 = arrayOfInputs(45)
    parNpd3 = arrayOfInputs(46)
    parEndPreparation3 = arrayOfInputs(47)
    parScheduleThickness3 = arrayOfInputs(48)
    parEndStandard3 = arrayOfInputs(49)
    parPressureRating3 = arrayOfInputs(50)
    parFlowDirection3 = arrayOfInputs(51)
    
    parOperationLength = arrayOfInputs(52)
    parOperationHeight = arrayOfInputs(53)
    parOperationWidth = arrayOfInputs(54)
    parOperationWidth1 = arrayOfInputs(55)

    'Comment to make test with text inputs
    parId1 = arrayOfInputs(56)
    parId2 = arrayOfInputs(57)
    parId3 = arrayOfInputs(58)
    parNPDUnitType = arrayOfInputs(59)
    parNPDUnitType1 = arrayOfInputs(60)
    parNPDUnitType2 = arrayOfInputs(61)
    parNPDUnitType3 = arrayOfInputs(62)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
'   The orgin is considered to be at intersection of the nozzle axis

' Insert your code for output 1(ValveBody)
    stPoint.Set -parInstrumentLength1, -parInstrumentHeight, -parInstrumentWidth
    enPoint.Set parInstrumentLength, parInstrumentHeight1, parInstrumentWidth1
    Set ObjValveBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing

 ' Insert your code for output 2(Cyl1Body)
    stPoint.Set parCylOffset, _
                        parInstrumentHeight1 - (parInstrumentHeight2 - parInstrumentHeight3), _
                        -(parInstrumentWidth + parInstrumentRadius)
    enPoint.Set parCylOffset, _
                        parInstrumentHeight1 + parInstrumentHeight3, _
                        -(parInstrumentWidth + parInstrumentRadius)
    Set ObjCyl1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl1Body
    Set ObjCyl1Body = Nothing

 ' Insert your code for output 3(Cyl2Body)
    stPoint.Set 0, _
                    parInstrumentHeight1 - (parInstrumentHeight4 - parInstrumentHeight5), _
                    -(parInstrumentWidth + parInstrumentRadius1)
    enPoint.Set 0, _
                        parInstrumentHeight1 + parInstrumentHeight5, _
                        -(parInstrumentWidth + parInstrumentRadius1)
    Set ObjCyl2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius1, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl2Body
    Set ObjCyl2Body = Nothing

 ' Insert your code for output 4(Cyl3Body)
    stPoint.Set -parCylOffset1, _
                    parInstrumentHeight1 - (parInstrumentHeight6 - parInstrumentHeight7), _
                    -(parInstrumentWidth + parInstrumentRadius2)
    enPoint.Set -parCylOffset1, _
                        parInstrumentHeight1 + parInstrumentHeight7, _
                    -(parInstrumentWidth + parInstrumentRadius2)
    Set ObjCyl3Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl3Body
    Set ObjCyl3Body = Nothing
  
 ' Insert your code for output 5(Cyl4Body)
    stPoint.Set 0, _
                    parInstrumentHeight1 - (parInstrumentHeight8 - parInstrumentHeight9), _
                    parInstrumentWidth1 + parInstrumentRadius3
    enPoint.Set 0, _
                    parInstrumentHeight1 + parInstrumentHeight9, _
                    parInstrumentWidth1 + parInstrumentRadius3
    Set ObjCyl4Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius3, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl4Body
    Set ObjCyl4Body = Nothing
       Set stPoint = Nothing
    Set enPoint = Nothing
    
   If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
'CreateRetrieveDynamicNozzle
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
    oNozzle.Length = parFace1toCenter - parInstrumentLength1

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
    oNozzle.Length = parFace1toCenter - parInstrumentLength

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
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
    oNozzle.Length = parFace3toCenter - parInstrumentHeight

' Direction of the Nozzle

    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace3toCenter - sptOffset(parPortIndex3) + depth(parPortIndex3), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
        
    
End Sub

