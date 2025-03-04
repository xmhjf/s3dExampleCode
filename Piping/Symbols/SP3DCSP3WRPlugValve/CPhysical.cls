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
'     This is PDS on-the-fly I28A  3 Way Rotary Plug Valve Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
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
    Dim ObjLeftCone As Object
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    Dim ObjTopCone As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)     'A
    parFace1toCenter = arrayOfInputs(3)     'B
    parFace2toCenter = arrayOfInputs(4)     'C
    parInsulationThickness = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)                       'Input parameters for Dynamic Nozzles
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
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFacetoCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

'' Place Nozzle 2
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
    
'' Place Nozzle 3
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
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing

 ' Insert your code for output 4(Left Cone)
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    
    
    BaseCenPoint.Set -parFacetoCenter + flangeThick(parPortIndex1), 0#, 0#
    TopCenPoint.Set 0#, 0#, 0#
    Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam(parPortIndex1) / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
    
 ' Insert your code for output 5(Upper Cone)
    Dim UpperConeBasePt As New AutoMath.DPosition
    Dim UpperConeTopPt As New AutoMath.DPosition
    Dim coneDia As Double
    Dim coneHt As Double
    
'    coneDia = parFacetoCenter * 0.5
    coneDia = pipeDiam(parPortIndex1) * 0.5 'As per EDEN
    If (pipeDiam(parPortIndex1) < 2 * 25.4 / 1000) Then 'Pipe diameter is less than two inch
        coneDia = 25.4 / 1000
    End If
    coneHt = pipeDiam(parPortIndex1) * 0.5 'As per EDEN
    UpperConeBasePt.Set 0#, 0#, 0#
    UpperConeTopPt.Set 0#, coneHt, 0#
    
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperConeBasePt, UpperConeTopPt, coneDia / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    Set UpperConeBasePt = Nothing
    Set UpperConeTopPt = Nothing
    
 ' Insert your code for output 6(Lower Cone)
    Dim LowerConeBasePt As New AutoMath.DPosition
    Dim LowerConeTopPt As New AutoMath.DPosition
    
    LowerConeBasePt.Set 0#, 0#, 0#
    LowerConeTopPt.Set 0#, -coneHt, 0#
    
    Set ObjLowerCone = PlaceCone(m_OutputColl, LowerConeBasePt, LowerConeTopPt, coneDia / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowerCone
    Set ObjLowerCone = Nothing
    Set LowerConeBasePt = Nothing
    Set LowerConeTopPt = Nothing
    
 ' Insert your code for output 7(Right Cone)
    Dim RightConeBasePt As New AutoMath.DPosition
    Dim RightConeTopPt  As New AutoMath.DPosition
    
    RightConeBasePt.Set parFace1toCenter - flangeThick(parPortIndex2), 0#, 0#
    RightConeTopPt.Set 0#, 0#, 0#
    Set ObjRightCone = PlaceCone(m_OutputColl, RightConeBasePt, RightConeTopPt, pipeDiam(parPortIndex2) / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set RightConeBasePt = Nothing
    Set RightConeTopPt = Nothing
    
    
' Insert your code for output 8(Top Cone)
    Dim TopConeBasePt As New AutoMath.DPosition
    Dim TopConeTopPt  As New AutoMath.DPosition
    

    TopConeBasePt.Set 0, parFace2toCenter - flangeThick(parPortIndex3), 0#
    TopConeTopPt.Set 0#, 0#, 0#
    
    Set ObjTopCone = PlaceCone(m_OutputColl, TopConeBasePt, TopConeTopPt, pipeDiam(parPortIndex3) / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCone
    Set ObjTopCone = Nothing
    Set TopConeBasePt = Nothing
    Set TopConeTopPt = Nothing
            
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
Private Sub Class_Terminate()

    Set m_oCodeListMetadata = Nothing
    
End Sub
