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
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S3A7 Inverted Bucket Steam Trap symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const MIN_LENGTH = 0.0001

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
        
    Dim parFacetoFace As Double
    Dim parCylinderDiameter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzleOffset2 As Double
    Dim parNozzleOffset3  As Double
    
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozzleOffset = arrayOfInputs(2)  'A
    parCylinderDiameter = arrayOfInputs(3)  'B
    parFacetoFace = arrayOfInputs(4)    'C
    parNozzleOffset1 = arrayOfInputs(5) 'D
    parNozzleOffset2 = arrayOfInputs(6) 'E
    parNozzleOffset3 = arrayOfInputs(7) 'F
    parNPD = arrayOfInputs(8)
    parEndPreparation = arrayOfInputs(9)
    parScheduleThickness = arrayOfInputs(10)
    parEndStandard = arrayOfInputs(11)
    parPressureRating = arrayOfInputs(12)
    parFlowDirection = arrayOfInputs(13)
    parPortIndex1 = arrayOfInputs(14)
    parNpd1 = arrayOfInputs(15)
    parEndPreparation1 = arrayOfInputs(16)
    parScheduleThickness1 = arrayOfInputs(17)
    parEndStandard1 = arrayOfInputs(18)
    parPressureRating1 = arrayOfInputs(19)
    parFlowDirection1 = arrayOfInputs(20)
    parPortIndex2 = arrayOfInputs(21)
    parNpd2 = arrayOfInputs(22)
    parEndPreparation2 = arrayOfInputs(23)
    parScheduleThickness2 = arrayOfInputs(24)
    parEndStandard2 = arrayOfInputs(25)
    parPressureRating2 = arrayOfInputs(26)
    parFlowDirection2 = arrayOfInputs(27)
    parPortIndex3 = arrayOfInputs(28)
    parNpd3 = arrayOfInputs(29)
    parEndPreparation3 = arrayOfInputs(30)
    parScheduleThickness3 = arrayOfInputs(31)
    parEndStandard3 = arrayOfInputs(32)
    parPressureRating3 = arrayOfInputs(33)
    parFlowDirection3 = arrayOfInputs(34)
    parPortIndex4 = arrayOfInputs(35)
    parNpd4 = arrayOfInputs(36)
    parEndPreparation4 = arrayOfInputs(37)
    parScheduleThickness4 = arrayOfInputs(38)
    parEndStandard4 = arrayOfInputs(39)
    parPressureRating4 = arrayOfInputs(40)
    parFlowDirection4 = arrayOfInputs(41)
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parId3 = arrayOfInputs(44)
    parId4 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    parNPDUnitType3 = arrayOfInputs(49)
    parNPDUnitType4 = arrayOfInputs(50)

    iOutput = 0
    
'   Origin is taken at center point of port1 and port2

' Place Nozzle 1
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
    
    Dim dNozzleLength As Double

    dNozzleLength = parFacetoFace / 2
    If dNozzleLength < flangeThick(parPortIndex1) Then dNozzleLength = flangeThick(parPortIndex1)
    oNozzle.Length = dNozzleLength

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
    dNozzleLength = parFacetoFace / 2
    If dNozzleLength < flangeThick(parPortIndex2) Then dNozzleLength = flangeThick(parPortIndex2)
    oNozzle.Length = dNozzleLength

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
    
    Dim dCylinderLength As Double
    
    dCylinderLength = parNozzleOffset - 3 * flangeThick(parPortIndex3)

    dNozzleLength = parNozzleOffset - dCylinderLength
    If dNozzleLength < MIN_LENGTH Then dNozzleLength = MIN_LENGTH
    
    oNozzle.Length = dNozzleLength

' Direction of the Nozzle
    oDir.Set 0, 1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parCylinderDiameter / 2 - parNozzleOffset2, (parNozzleOffset - parNozzleOffset1) + _
              sptOffset(parPortIndex3) - depth(parPortIndex3), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

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
    dNozzleLength = parFacetoFace / 2
    If dNozzleLength < flangeThick(parPortIndex4) Then dNozzleLength = flangeThick(parPortIndex4)
    oNozzle.Length = dNozzleLength

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFacetoFace / 2 + sptOffset(parPortIndex4) - depth(parPortIndex4), _
                    -parNozzleOffset1 + parNozzleOffset3, 0
     
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing

 ' Insert your code for output 5(Cylinder)
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    

    BaseCenPoint.Set 0, -parNozzleOffset1, 0
    TopCenPoint.Set BaseCenPoint.x, BaseCenPoint.y + dCylinderLength, BaseCenPoint.z
    Dim ObjCylinder  As Object
    Set ObjCylinder = PlaceCylinder(m_OutputColl, BaseCenPoint, TopCenPoint, parCylinderDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
                    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing

End Sub
