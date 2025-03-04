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
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 8 2004
'   Description:
'     This is PDS on-the-fly I10A  With Upstream Valve Symbol.
'
'   Change History:
'   dd.mmm.yyyy    who                      change description
'   -----------   -----                     ------------------
'  08.SEP.2006     KKC          DI-95670  Replace names with initials in all revision history sheets and symbols
'  25.Sep.2006     KKC          TR-102869 Modified symbol code to use both Face to Face
'                                         and Face to Center Dimensions
'  2.Jul.2007       KKC                     TR-123363: Corrected the symbol not to fail the V6.1 Instances
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Const MODULE = "Physical:" 'Used for error messages
Const MIN_LENGTH = 0.0001

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
    
    Dim parFacetoFace As Double
    Dim parWidth As Double
    Dim parHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim lPartDataBasis As Long
    
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
'    parFacetoFace = arrayOfInputs(2) 'A
    parWidth = arrayOfInputs(3)      'B
    parHeight = arrayOfInputs(4)     'C
    parInsulationThickness = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
'    parFacetoCenter = arrayOfInputs(26)
'    parId1 = arrayOfInputs(27)
'    parId2 = arrayOfInputs(28)
'    parNPDUnitType = arrayOfInputs(29)
'    parNPDUnitType1 = arrayOfInputs(30)
'    parNPDUnitType2 = arrayOfInputs(31)
'   Check to ensure both V6.1 and V7 symbol definitions work
    If UBound(arrayOfInputs) >= 31 Then
        'In V7, Index 27 onwards are textual inputs
        parId1 = arrayOfInputs(27)
        parId2 = arrayOfInputs(28)
        parNPDUnitType = arrayOfInputs(29)
        parNPDUnitType1 = arrayOfInputs(30)
        parNPDUnitType2 = arrayOfInputs(31)
    Else
        'In V6.1,Index 26 onwards are textual inputs
        parId1 = arrayOfInputs(26)
        parId2 = arrayOfInputs(27)
        parNPDUnitType = arrayOfInputs(28)
        parNPDUnitType1 = arrayOfInputs(29)
        parNPDUnitType2 = arrayOfInputs(30)
    End If

    iOutput = 0
    
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then
       parFacetoFace = arrayOfInputs(2)
    ElseIf lPartDataBasis = 10 Then
       parFacetoCenter = arrayOfInputs(26)
       parFacetoFace = 2 * parFacetoCenter
    End If
    
    Dim dBody_length As Double
    Dim dNoz_len  As Double
    
    dBody_length = parFacetoFace * 0.6
    dNoz_len = (parFacetoFace - dBody_length) / 2#
    If CmpDblLessThan(dNoz_len, MIN_LENGTH) Then dNoz_len = MIN_LENGTH
    
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
    oNozzle.Length = dNoz_len

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    oNozzle.Length = dNoz_len

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFacetoFace / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
' Insert your code for output 8(Hammer Arrestor Body)
    Point1.Set -dBody_length / 2, -parHeight / 2, -parWidth / 2
    Point2.Set dBody_length / 2, parHeight / 2, parWidth / 2

    Dim ObjValveBody As Object
    Set ObjValveBody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
    
