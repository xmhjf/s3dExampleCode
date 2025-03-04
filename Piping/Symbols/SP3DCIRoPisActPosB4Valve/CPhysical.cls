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
'   Author:          ACM
'   Creation Date:  Friday, Jan 3 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPR4B),If this is
'     rotated through 18deg about flow axis then we get (PDS Eden Name = IPL4C) symbol.
'   The symbol consists of Nine outputs,Eight physical outputs
'   and One Insulation Output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim ObjValveCone1 As Object
    Dim ObjValveCone2 As Object
    Dim ObjStemRod As Object
    Dim ObjValveCylinder As Object
    Dim ObjActuatorRect As Object
    Dim ObjActuatorCyl As Object

    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorWidth2 As Double
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
    parFace1toCenter = arrayOfInputs(2) ''D2
    parFacetoFace = arrayOfInputs(3) ''D3''
    parCylHeight = arrayOfInputs(4) ''D4
    parActuatorLength = arrayOfInputs(5) ''D5
    parActuatorDiameter = arrayOfInputs(6) ''D6
    parActuatorHeight = arrayOfInputs(7) ''D7
    parActuatorHeight1 = arrayOfInputs(8) ''D10
    parActuatorHeight2 = arrayOfInputs(9) ''D11
    parActuatorWidth = arrayOfInputs(10) ''D12
    parActuatorWidth1 = arrayOfInputs(11) ''D14
    parActuatorLength1 = arrayOfInputs(12) ''D25
    parActuatorLength2 = arrayOfInputs(13) ''D26
    parActuatorWidth2 = arrayOfInputs(14) ''D28
    parInsulationThickness = arrayOfInputs(15)
    
    parNPD = arrayOfInputs(16)                       'Input parameters for Dynamic Nozzles
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
    parId1 = arrayOfInputs(36)
    parId2 = arrayOfInputs(37)
    parNPDUnitType = arrayOfInputs(38)
    parNPDUnitType1 = arrayOfInputs(39)
    parNPDUnitType2 = arrayOfInputs(40)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
        
' Place Nozzle 1 (Nozzle1)
    
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
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    

 ' Insert your code for output 3(Conical Valve Body 1)
         
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
         
    stPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    enPoint.Set 0, 0, 0
    
    Set ObjValveCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone1
    Set ObjValveCone1 = Nothing
    
 ' Insert your code for output 4(Conical Valve Body 2 )
         
    stPoint.Set parFacetoFace - parFace1toCenter - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
        
    Set ObjValveCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone2
    Set ObjValveCone2 = Nothing
    
 ' Insert your code for output 5(Stem Body )
   Dim StemDia As Double
   
   StemDia = 0.05
   stPoint.Set 0, 0, 0
   enPoint.Set 0, parActuatorHeight - parActuatorHeight2, 0
 
   Set ObjStemRod = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemRod
    Set ObjStemRod = Nothing
    
 ' Insert your code for output 6(Valve Cyledrical Body)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    
    CylDia = parFace1toCenter
    
    Set ObjValveCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCylinder
    Set ObjValveCylinder = Nothing
    
 ' Insert your code for output 7(Actuator Rectangular box )
 
    stPoint.Set parActuatorLength1, parActuatorHeight - parActuatorHeight2, _
                                        parActuatorWidth1 + parActuatorWidth2
    enPoint.Set -parActuatorLength2, parActuatorHeight + parActuatorHeight1, _
                                        -(parActuatorWidth - parActuatorWidth2)
    
   Set ObjActuatorRect = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorRect
    Set ObjActuatorRect = Nothing
    
 ' Insert your code for output 8(Actuator Cylendrical Body)
   
   stPoint.Set parActuatorLength1, parActuatorHeight, parActuatorWidth2
   enPoint.Set parActuatorLength, parActuatorHeight, parActuatorWidth2
   
   Set ObjActuatorCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCyl
    Set ObjActuatorCyl = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

