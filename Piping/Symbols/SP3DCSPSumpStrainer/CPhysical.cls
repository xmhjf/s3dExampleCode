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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S1A7  Sump Strainer Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()

''''

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
    Dim parLength As Double
    
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
    parLength = arrayOfInputs(2) 'A
    parNPD = arrayOfInputs(3)
    parEndPreparation = arrayOfInputs(4)
    parScheduleThickness = arrayOfInputs(5)
    parEndStandard = arrayOfInputs(6)
    parPressureRating = arrayOfInputs(7)
    parFlowDirection = arrayOfInputs(8)
    parPortIndex1 = arrayOfInputs(9)
    parNpd1 = arrayOfInputs(10)
    parEndPreparation1 = arrayOfInputs(11)
    parScheduleThickness1 = arrayOfInputs(12)
    parEndStandard1 = arrayOfInputs(13)
    parPressureRating1 = arrayOfInputs(14)
    parFlowDirection1 = arrayOfInputs(15)
    parPortIndex2 = arrayOfInputs(16)
    parNpd2 = arrayOfInputs(17)
    parEndPreparation2 = arrayOfInputs(18)
    parScheduleThickness2 = arrayOfInputs(19)
    parEndStandard2 = arrayOfInputs(20)
    parPressureRating2 = arrayOfInputs(21)
    parFlowDirection2 = arrayOfInputs(22)
    parId1 = arrayOfInputs(23)
    parId2 = arrayOfInputs(24)
    parNPDUnitType = arrayOfInputs(25)
    parNPDUnitType1 = arrayOfInputs(26)
    parNPDUnitType2 = arrayOfInputs(27)
    
    iOutput = 0
     
'   Origin is taken at center of port1 and port2

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
    oPlacePoint.Set -flangeThick(parPortIndex1) / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
'   Port 2 is at flangeThick(parPortIndex1) / 2 from the origin
    oPlacePoint.Set flangeThick(parPortIndex1) / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

 ' Insert your code for output 3(Strainer)
    Dim baseCenPt As AutoMath.DPosition
    Dim topCenPt As AutoMath.DPosition
    Set baseCenPt = New AutoMath.DPosition
    Set topCenPt = New AutoMath.DPosition

    baseCenPt.Set flangeThick(parPortIndex1) / 2, 0, 0
    topCenPt.Set parLength - flangeThick(parPortIndex1) / 2, 0, 0
    Dim ObjStrainer As Object
    Set ObjStrainer = PlaceCylinder(m_OutputColl, baseCenPt, topCenPt, pipeDiam(parPortIndex1), True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainer
    Set ObjStrainer = Nothing
    Set baseCenPt = Nothing
    Set topCenPt = Nothing

    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
