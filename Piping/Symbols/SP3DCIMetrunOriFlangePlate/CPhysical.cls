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
'   Author:          SSP
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMOP1,
'   PDMS Category Name=IFRB). The symbol consists of 9 oupputs: 8 physical,
'   and 1 insulation aspect outputs.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006        KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.Aug.2007        KKC                   CR-122175 Corrected the code not to have negative nozzle length
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


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
    Dim ObjDivergingCone As Object
    Dim ObjConvergingCone As Object
    Dim ObjFlange1 As Object
    Dim ObjFlange2 As Object
    
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parOrificeFlangeClearance As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double
    Dim NozzleLength As Double
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
    parFacetoFace = arrayOfInputs(2)            'D2
    parFace1toCenter = arrayOfInputs(3)       'D3
    parOrificeFlangeClearance = arrayOfInputs(4)               'D4
    parPortRotation1 = arrayOfInputs(5)                     'D5
    parPortRotation2 = arrayOfInputs(6)                   'D6
    parInsulationThickness = arrayOfInputs(7)
    parNPD = arrayOfInputs(8)                                'Input parameters for Dynamic Nozzles
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
    'Text Inputs
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
' The orgin is considered to be along the flow axis and at parFace1toCenter plus parClearanc/2 from nozzle 1
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle1             As GSCADNozzleEntities.IJDNozzle
    Dim oNozzle2             As GSCADNozzleEntities.IJDNozzle
    Dim oNozzle3             As GSCADNozzleEntities.IJDNozzle
    Dim oNozzle4             As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    Set oNozzle1 = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle1
    Set oDistribPort = oNozzle1
      
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parOrificeFlangeClearance / 2 - parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Place Nozzle 3
    Set oNozzle3 = CreateRetrieveDynamicNozzle(parPortIndex3, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNPDUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)

    Set oLogicalDistPort = oNozzle3
    Set oDistribPort = oNozzle3
    
    Dim xCord As Double
    Dim yCord As Double
    Dim zCord As Double
    Dim flgneck1 As Double
    Dim flgthk1 As Double
    Dim flgod1 As Double
'   Assumptions:
'   flgthk1 = pipeDiam(parPortIndex3) * 1.4
'   flgneck1=flgthk1*1.5
'   flgod1 = pipeDiam(parPortIndex1) * 2 are taken
    flgthk1 = pipeDiam(parPortIndex3) * 1.4
    flgneck1 = flgthk1 * 1.5
    flgod1 = pipeDiam(parPortIndex1) * 2

    xCord = -parOrificeFlangeClearance / 2 - flgthk1 / 2
    yCord = (flgod1 / 2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Cos(parPortRotation1)
    zCord = (flgod1 / 2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Sin(parPortRotation1)
    
' Direction of the Nozzle
    oDir.Set 0, Cos(parPortRotation1), Sin(parPortRotation1)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set xCord, yCord, zCord
    oLogicalDistPort.SetCenterLocation oPlacePoint

'Nozzle1 length
    If (parFace1toCenter - flgthk1 - flgneck1) < 0 Then
        oNozzle1.Length = 0.00001
        
    Else
         oNozzle1.Length = parFace1toCenter - flgthk1 - flgneck1
    End If
    
  ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle1
    Set oNozzle1 = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle3
    Set oNozzle3 = Nothing

' Place Nozzle 2
    Set oNozzle2 = CreateRetrieveDynamicNozzle(parPortIndex2, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNPDUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle2
    Set oDistribPort = oNozzle2
    
' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFacetoFace - parFace1toCenter - parOrificeFlangeClearance / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
           
' Place Nozzle 4
    Set oNozzle4 = CreateRetrieveDynamicNozzle(parPortIndex4, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNPDUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle4
    Set oDistribPort = oNozzle4
    
    Dim flgneck2 As Double
    Dim flgthk2 As Double
    Dim flgod2 As Double
'    Assumptions:
'    flgthk2 = pipeDiam(parPortIndex4)* 1.4
'    flgneck1=flgthk2*1.5
'    flgod2 = pipeDiam(parPortIndex2) * 2 are taken
    flgthk2 = pipeDiam(parPortIndex4) * 1.4
    flgneck2 = flgthk2 * 1.5
    flgod2 = pipeDiam(parPortIndex2) * 2

    xCord = parOrificeFlangeClearance / 2 + flgthk2 / 2
    yCord = (flgod2 / 2 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Cos(parPortRotation2)
    zCord = (flgod2 / 2 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Sin(parPortRotation2)

' Direction of the Nozzle
    oDir.Set 0, Cos(parPortRotation2), Sin(parPortRotation2)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set xCord, yCord, zCord
    oLogicalDistPort.SetCenterLocation oPlacePoint
' Nozzle4 length

    If (parFacetoFace - parFace1toCenter - parOrificeFlangeClearance - flgthk2 - flgneck2) < 0 Then
       oNozzle2.Length = 0.0001
    Else
      oNozzle2.Length = parFacetoFace - parFace1toCenter - parOrificeFlangeClearance - flgthk2 - flgneck2
    End If
    

' Set the output for Nozzle2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle2
    Set oNozzle2 = Nothing

' Set the output for Nozzle4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle4
    Set oNozzle4 = Nothing

' Insert your code for output 5(DivergingCone)
    Dim coneBaseRad As Double
    Dim coneTopRad As Double

    stPoint.Set -parOrificeFlangeClearance / 2 - flgthk1, 0, 0
    enPoint.Set -parOrificeFlangeClearance / 2 - flgthk1 - flgneck1, 0, 0

'    Assumptions:
'    coneTopRad = pipeDiam(parPortIndex1)/2
'    coneBaseRad=1.6 * pipeDiam(parPortIndex1)/2 are taken
    coneBaseRad = 1.6 * pipeDiam(parPortIndex1) / 2
    coneTopRad = pipeDiam(parPortIndex1) / 2
    
    Set ObjDivergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, coneBaseRad, coneTopRad, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    
 ' Insert your code for output 6(ConvergingCone)
    stPoint.Set parOrificeFlangeClearance / 2 + flgthk2, 0, 0
    enPoint.Set parOrificeFlangeClearance / 2 + flgthk2 + flgneck2, 0, 0

'    Assumptions:
'    coneTopRad = pipeDiam(parPortIndex2)/2
'    coneBaseRad=1.6 * pipeDiam(parPortIndex2)/2 are taken
    coneBaseRad = 1.6 * pipeDiam(parPortIndex2) / 2
    coneTopRad = pipeDiam(parPortIndex2) / 2
    
    Set ObjConvergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, coneBaseRad, coneTopRad, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
 ' Insert your code for output 7(Flange1)
    stPoint.Set -parOrificeFlangeClearance / 2, 0, 0
    enPoint.Set -parOrificeFlangeClearance / 2 - flgthk1, 0, 0

    Set ObjFlange1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flgod1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange1
    Set ObjFlange1 = Nothing
    
 ' Insert your code for output 8(Flange2)
    stPoint.Set parOrificeFlangeClearance / 2, 0, 0
    enPoint.Set parOrificeFlangeClearance / 2 + flgthk2, 0, 0
    
    Set ObjFlange2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flgod2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange2
    Set ObjFlange2 = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
            
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
