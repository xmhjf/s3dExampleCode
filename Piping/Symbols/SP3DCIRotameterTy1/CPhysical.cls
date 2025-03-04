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
'   Author:          BG
'   Creation Date:  Thusday Dec 26, 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR1,
'   PDMS Category Name=IFNG).This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. The symbol consists of seven outputs, five physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Sub Class_Initialize()
''
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
    Dim ObjCyl1 As Object
    Dim ObjCyl2 As Object
    Dim ObjCyl3 As Object
    
    Dim parInstrumentHeight As Double            'D2
    Dim parInstrumentDiameter As Double          'D3
    Dim parInstrumentHeight1 As Double           'D4
    Dim parFacetoFace As Double                  'D5
    Dim parInstrumentLength As Double            'D6
    Dim parInstrumentDiameter1 As Double         'D7
    Dim parInstrumentDiameter2 As Double         'D8
    Dim parInstrumentLength1 As Double           'D9
    Dim parInstrumentHeight2 As Double           'D10
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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)
    parInstrumentDiameter = arrayOfInputs(3)
    parInstrumentHeight1 = arrayOfInputs(4)
    parFacetoFace = arrayOfInputs(5)
    parInstrumentLength = arrayOfInputs(6)
    parInstrumentDiameter1 = arrayOfInputs(7)
    parInstrumentDiameter2 = arrayOfInputs(8)
    parInstrumentLength1 = arrayOfInputs(9)
    parInstrumentHeight2 = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    parNPD = arrayOfInputs(12)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(13)
    parScheduleThickness = arrayOfInputs(14)
    parEndStandard = arrayOfInputs(15)
    parPressureRating = arrayOfInputs(16)
    parFlowDirection = arrayOfInputs(17)
    parPortIndex1 = arrayOfInputs(18)
    parNpd1 = arrayOfInputs(19)
    parEndPreparation1 = arrayOfInputs(20)
    parScheduleThickness1 = arrayOfInputs(21)
    parEndStandard1 = arrayOfInputs(22)
    parPressureRating1 = arrayOfInputs(23)
    parFlowDirection1 = arrayOfInputs(24)
    parPortIndex2 = arrayOfInputs(25)
    parNpd2 = arrayOfInputs(26)
    parEndPreparation2 = arrayOfInputs(27)
    parScheduleThickness2 = arrayOfInputs(28)
    parEndStandard2 = arrayOfInputs(29)
    parPressureRating2 = arrayOfInputs(30)
    parFlowDirection2 = arrayOfInputs(31)
    parId1 = arrayOfInputs(32)
    parId2 = arrayOfInputs(33)
    parNPDUnitType = arrayOfInputs(34)
    parNPDUnitType1 = arrayOfInputs(35)
    parNPDUnitType2 = arrayOfInputs(36)

    iOutput = 0
    
 ' Insert your code for output 1(Cyl1)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -(parInstrumentDiameter / 2 + parInstrumentLength1), 0, 0
    enPoint.Set -parInstrumentDiameter / 2, 0, 0
    
    Set ObjCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl1
    Set ObjCyl1 = Nothing
    
 ' Insert your code for output 2(Cyl2)
    stPoint.Set 0, -(parInstrumentHeight1 - (parInstrumentHeight2 - parInstrumentHeight)), _
                0
    enPoint.Set 0, parInstrumentHeight2 - parInstrumentHeight1, 0
    
    Set ObjCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl2
    Set ObjCyl2 = Nothing
    
 ' Insert your code for output 3(Cyl3)
    
    stPoint.Set parInstrumentDiameter / 2, 0, 0
    
' This checks in case D6 = D3/2 which causes a failure of placecylinder.
    If parInstrumentLength = parInstrumentDiameter / 2 Then
        enPoint.Set stPoint.x + 0.001, 0, 0
    Else
        enPoint.Set parInstrumentLength, 0, 0
    End If
    
    Set ObjCyl3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl3
    Set ObjCyl3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' Place Nozzle 1 for output 5 (Nozzle1)
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

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
    oPlacePoint.Set -(parInstrumentDiameter / 2 + parInstrumentLength1) - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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

'   In normal scenario nozzle length will be equal to parInstrumentDiameter / 2,
'   if parInstrumentDiameter / 2 is less than (or equal to) the flange thickness then nozzle
'   length will be equal to flange thickness
    If (parInstrumentDiameter / 2) <= flangeThick(parPortIndex2) Then
        oNozzle.Length = flangeThick(parPortIndex2)
    Else
        oNozzle.Length = parInstrumentDiameter / 2
    End If

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oPlacePoint.Set -parInstrumentDiameter / 2 - sptOffset(parPortIndex2) + depth(parPortIndex2), parFacetoFace, 0
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

