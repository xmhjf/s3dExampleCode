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
'   Author:          Sundar(svsmylav)
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly I16A Custom Speciality Rotary Plug Valve symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

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
    Dim ObjLeftCone As Object
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double
    
'   Dynamic Port Input parameters variable declaration
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
    parFacetoCenter = arrayOfInputs(2)  'A
    parFace1toCenter = arrayOfInputs(3) 'B
    parInsulationThickness = arrayOfInputs(4)
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parPortIndex2 = arrayOfInputs(18)
    parNpd2 = arrayOfInputs(19)
    parEndPreparation2 = arrayOfInputs(20)
    parScheduleThickness2 = arrayOfInputs(21)
    parEndStandard2 = arrayOfInputs(22)
    parPressureRating2 = arrayOfInputs(23)
    parFlowDirection2 = arrayOfInputs(24)
    parId1 = arrayOfInputs(25)
    parId2 = arrayOfInputs(26)
    parNPDUnitType = arrayOfInputs(27)
    parNPDUnitType1 = arrayOfInputs(28)
    parNPDUnitType2 = arrayOfInputs(29)
        
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

     
' Place Nozzle 1
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
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing

 ' Insert your code for output 3(Left Cone)
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    Dim dLength_1 As Double
    dLength_1 = parFacetoCenter - flangeThick(parPortIndex1)
    BaseCenPoint.Set -dLength_1, 0, 0
    TopCenPoint.Set 0#, 0#, 0#
    Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam(parPortIndex1) / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
    
 ' Insert your code for output 4(Upper Cone)
    Dim two As Double
    Dim ten As Double
    Dim factor As Double
    
    two = 2 * 25.4 / 1000
    ten = 10 * 25.4 / 1000

    If (pipeDiam(parPortIndex1) <= two) Then
        factor = 0.4
    Else
        If (pipeDiam(parPortIndex1) > two And pipeDiam(parPortIndex1) <= ten) Then
            factor = 0.45
        Else
            If (pipeDiam(parPortIndex1) > ten) Then
                factor = 0.5
            End If
        End If
    End If
    Dim UpperConeBasePt As New AutoMath.DPosition
    Dim UpperConeTopPt As New AutoMath.DPosition
    Dim coneDia As Double
    Dim coneHt As Double
    coneDia = dLength_1 * factor
    coneHt = pipeDiam(parPortIndex1) * 0.5  'As per EDEN
'    If pipeDiam(parPortIndex2) > pipeDiam(parPortIndex1) Then
'        coneHt = pipeDiam(parPortIndex2) * 0.5
'    End If
    UpperConeBasePt.Set 0#, 0#, 0#
    UpperConeTopPt.Set 0, coneHt, 0
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperConeBasePt, UpperConeTopPt, coneDia / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    Set UpperConeBasePt = Nothing
    Set UpperConeTopPt = Nothing
    
 ' Insert your code for output 5(Lower Cone)
    Dim LowerConeBasePt As New AutoMath.DPosition
    Dim LowerConeTopPt As New AutoMath.DPosition
    LowerConeBasePt.Set 0#, 0#, 0#
    LowerConeTopPt.Set 0, -coneHt, 0
    Set ObjLowerCone = PlaceCone(m_OutputColl, LowerConeBasePt, LowerConeTopPt, coneDia / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowerCone
    Set ObjLowerCone = Nothing
    Set LowerConeBasePt = Nothing
    Set LowerConeTopPt = Nothing
    
 ' Insert your code for output 6(Right Cone)
    Dim RightConeBasePt As New AutoMath.DPosition
    Dim RightConeTopPt  As New AutoMath.DPosition
    RightConeBasePt.Set parFace1toCenter - flangeThick(parPortIndex2), 0, 0
    RightConeTopPt.Set 0#, 0#, 0#
    Set ObjRightCone = PlaceCone(m_OutputColl, RightConeBasePt, RightConeTopPt, pipeDiam(parPortIndex2) / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set RightConeBasePt = Nothing
    Set RightConeTopPt = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

