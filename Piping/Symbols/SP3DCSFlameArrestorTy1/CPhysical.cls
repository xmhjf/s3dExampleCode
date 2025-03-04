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
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No.: S1AC or S1AZ, Page No. D-99 PDS Piping Component Data Reference Guide.
'  Symbol is created with Eight Outputs - Five are Insulation aspect outputs.
' Other three are physical aspect outputs created as follows:
'   ObjFABody - Using 'PlaceBox' function and
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjFABody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parFlArrestorBodyLength1 As Double
    Dim parFlArrestorBodyLength2 As Double
    Dim parFlArrestorBodyHeight1 As Double
    Dim parFlArrestorBodyHeight2 As Double
    Dim parFlArrestorBodyWidth1 As Double
    Dim parFlArrestorBodyWidth2 As Double
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
    
    Dim Point1   As New AutoMath.DPosition
    Dim Point2   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'P2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parOffset1 = arrayOfInputs(4)       'P4
    parOffset2 = arrayOfInputs(5)       'P5
    parFlArrestorBodyLength1 = arrayOfInputs(6)   'P6
    parFlArrestorBodyLength2 = arrayOfInputs(7)     'P7
    parFlArrestorBodyHeight1 = arrayOfInputs(8)     'P8
    parFlArrestorBodyHeight2 = arrayOfInputs(9)     'P9
    parFlArrestorBodyWidth1 = arrayOfInputs(10)     'P10
    parFlArrestorBodyWidth2 = arrayOfInputs(11)     'P11
    parInsulationThickness = arrayOfInputs(12)
    parNPD = arrayOfInputs(13)
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parPortIndex1 = arrayOfInputs(19)
    parNpd1 = arrayOfInputs(20)
    parEndPreparation1 = arrayOfInputs(21)
    parScheduleThickness1 = arrayOfInputs(22)
    parEndStandard1 = arrayOfInputs(23)
    parPressureRating1 = arrayOfInputs(24)
    parFlowDirection1 = arrayOfInputs(25)
    parPortIndex2 = arrayOfInputs(26)
    parNpd2 = arrayOfInputs(27)
    parEndPreparation2 = arrayOfInputs(28)
    parScheduleThickness2 = arrayOfInputs(29)
    parEndStandard2 = arrayOfInputs(30)
    parPressureRating2 = arrayOfInputs(31)
    parFlowDirection2 = arrayOfInputs(32)
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parNPDUnitType = arrayOfInputs(35)
    parNPDUnitType1 = arrayOfInputs(36)
    parNPDUnitType2 = arrayOfInputs(37)

    iOutput = 0
' Insert your code for output 6(Flame Arrestor Body)
    Point1.Set -parFlArrestorBodyLength1, -parOffset1 - parFlArrestorBodyHeight1, parFlArrestorBodyWidth2
    Point2.Set parFlArrestorBodyLength2, parFlArrestorBodyHeight2 - parOffset2, -parFlArrestorBodyWidth1
    Set ObjFABody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFABody
    Set ObjFABody = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
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
    oNozzle.Length = parFace1toCenter - parFlArrestorBodyLength1

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), -parOffset1, 0
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
    oNozzle.Length = parFace2toCenter - parFlArrestorBodyLength2

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace2toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), parOffset2, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
        
    
End Sub
    
