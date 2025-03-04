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
'     This is PDS on-the-fly I18A  3 Way GlobeValve Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'
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
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjValveBodyLeftCone As Object
    Dim ObjValveBodyRightCone As Object
    Dim ObjValveBodyTopCone As Object
    Dim ObjSphere As Object
    Dim ObjValveBodyIns As Object
    Dim ObjValveBodyIns2 As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
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
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'A
    parFace1toCenter = arrayOfInputs(3) 'B
    parFace2toCenter = arrayOfInputs(4) 'C
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
    parPortIndex3 = arrayOfInputs(26)
    parNpd3 = arrayOfInputs(27)
    parEndPreparation3 = arrayOfInputs(28)
    parScheduleThickness3 = arrayOfInputs(29)
    parEndStandard3 = arrayOfInputs(30)
    parPressureRating3 = arrayOfInputs(31)
    parFlowDirection3 = arrayOfInputs(32)
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parId3 = arrayOfInputs(35)
    parNPDUnitType = arrayOfInputs(36)
    parNPDUnitType1 = arrayOfInputs(37)
    parNPDUnitType2 = arrayOfInputs(38)
    parNPDUnitType3 = arrayOfInputs(39)
    
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
    oNozzle.Length = flangeThick(parPortIndex3)

' Direction of the Nozzle
    oDir.Set 0, 1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, parFace2toCenter + sptOffset(parPortIndex3) - depth(parPortIndex3), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Insert your code for output 4(Left Cone of Valve Body)
    stPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjValveBodyLeftCone = PlaceCone(m_OutputColl, stPoint, enPoint, 0.5 * pipeDiam(parPortIndex1), 0.001)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyLeftCone
    Set ObjValveBodyLeftCone = Nothing
    
' Insert your code for output 5(Right Cone of Valve Body)
    stPoint.Set 0, 0, 0
    enPoint.Set parFace1toCenter - flangeThick(parPortIndex2), 0, 0
    Set ObjValveBodyRightCone = PlaceCone(m_OutputColl, stPoint, enPoint, 0.001, 0.5 * pipeDiam(parPortIndex2))
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyRightCone
    Set ObjValveBodyRightCone = Nothing
    
' Insert your code for output 5(Top Cone of Valve Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parFace2toCenter - flangeThick(parPortIndex3), 0
    Set ObjValveBodyTopCone = PlaceCone(m_OutputColl, stPoint, enPoint, 0.001, 0.5 * pipeDiam(parPortIndex3))
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyTopCone
    Set ObjValveBodyTopCone = Nothing
    
' Insert your code for output 4(Sphere)
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    
    Dim two As Double
    Dim ten As Double
    Dim factor As Double
    
    two = 2 * 25.4 / 1000
    ten = 10 * 25.4 / 1000

    If (pipeDiam(parPortIndex1) <= two) Then
        factor = 0.5
    Else
        If (pipeDiam(parPortIndex1) > two And pipeDiam(parPortIndex1) <= ten) Then
            factor = 0.55
        Else
            If (pipeDiam(parPortIndex1) > ten) Then
                factor = 0.6
            End If
        End If
    End If
    Dim sphere_radius As Double
    sphere_radius = (parFacetoCenter - flangeThick(parPortIndex1)) * factor * 0.5
    
    Set ObjSphere = PlaceSphere(m_OutputColl, oCenterPoint, sphere_radius)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSphere
    Set ObjSphere = Nothing
    
 
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

