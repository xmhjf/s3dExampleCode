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
'   Creation Date:  Wed,Dec 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMRSV, PDMS Category Name=IFRD)
'   The symbol consists of eleven outputs, eight physical and three insulation outputs.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   21.Jul.2003      ACM                            TR 44555 Put a check for nozzle length for all ports,such that it should'nt fail when Nozzle Length is less than FlangeThickness.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData
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
    Dim ObjCylinderBody1 As Object
    Dim ObjCylinderBody2 As Object
    Dim ObjReducingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjVenturiBodyIns As Object
    Dim ObjNozz3BodyIns As Object
    Dim ObjNozz4BodyIns As Object
    
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentDiameter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzleOffset2 As Double
    Dim parNozzleOffset3 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInstrumentLength4 As Double
    Dim parInstrumentLength5 As Double
    
    Dim parInsulationThickness As Double
    
    'Dynamic Port Input parameters variable declaration
    Dim parNpd                     As Double
    Dim parEndPreparation          As Long
    Dim parScheduleThickness       As Long
    Dim parEndStandard             As Long
    Dim parPressureRating          As Long
    Dim parFlowDirection           As DistribFlow
    Dim parPortIndex1              As Long
    Dim parNpd1                    As Double
    Dim parEndPreparation1         As Long
    Dim parScheduleThickness1      As Long
    Dim parEndStandard1            As Long
    Dim parPressureRating1         As Long
    Dim parFlowDirection1          As DistribFlow
    Dim parPortIndex2              As Long
    Dim parNpd2                    As Double
    Dim parEndPreparation2         As Long
    Dim parScheduleThickness2      As Long
    Dim parEndStandard2            As Long
    Dim parPressureRating2         As Long
    Dim parFlowDirection2          As DistribFlow
    Dim parPortIndex3              As Long
    Dim parNpd3                    As Double
    Dim parEndPreparation3         As Long
    Dim parScheduleThickness3      As Long
    Dim parEndStandard3            As Long
    Dim parPressureRating3         As Long
    Dim parFlowDirection3          As DistribFlow
    Dim parPortIndex4              As Long
    Dim parNpd4                    As Double
    Dim parEndPreparation4         As Long
    Dim parScheduleThickness4      As Long
    Dim parEndStandard4            As Long
    Dim parPressureRating4         As Long
    Dim parFlowDirection4          As DistribFlow
    Dim parId1                     As String
    Dim parId2                     As String
    Dim parId3                     As String
    Dim parId4                     As String
    Dim parNPDUnitType             As String
    Dim parNpdUnitType1            As String
    Dim parNpdUnitType2            As String
    Dim parNpdUnitType3            As String
    Dim parNpdUnitType4            As String

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim BaseRadius As Double
    Dim TopRadius As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)           'D2
    parInstrumentLength1 = arrayOfInputs(3)         'D3
    parInstrumentLength2 = arrayOfInputs(4)         'D4
    parInstrumentLength3 = arrayOfInputs(5)         'D5
    parInstrumentDiameter = arrayOfInputs(6)       'D6
    parNozzleOffset = arrayOfInputs(7)                  'D7
    parNozzleOffset1 = arrayOfInputs(8)                'D8
    parNozzleOffset2 = arrayOfInputs(9)                'D9
    parNozzleOffset3 = arrayOfInputs(10)              'D10
    parPortRotation1 = arrayOfInputs(11)                           'D11
    parPortRotation2 = arrayOfInputs(12)                         'D12
    parInstrumentLength4 = arrayOfInputs(13)       'D13
    parInstrumentLength5 = arrayOfInputs(14)       'D14
    parNpd = arrayOfInputs(15)
    parEndPreparation = arrayOfInputs(16)
    parScheduleThickness = arrayOfInputs(17)
    parEndStandard = arrayOfInputs(18)
    parPressureRating = arrayOfInputs(19)
    parFlowDirection = arrayOfInputs(20)
    parPortIndex1 = arrayOfInputs(21)
    parNpd1 = arrayOfInputs(22)
    parEndPreparation1 = arrayOfInputs(23)
    parScheduleThickness1 = arrayOfInputs(24)
    parEndStandard1 = arrayOfInputs(25)
    parPressureRating1 = arrayOfInputs(26)
    parFlowDirection1 = arrayOfInputs(27)
    parPortIndex2 = arrayOfInputs(28)
    parNpd2 = arrayOfInputs(29)
    parEndPreparation2 = arrayOfInputs(30)
    parScheduleThickness2 = arrayOfInputs(31)
    parEndStandard2 = arrayOfInputs(32)
    parPressureRating2 = arrayOfInputs(33)
    parFlowDirection2 = arrayOfInputs(34)
    parPortIndex3 = arrayOfInputs(35)
    parNpd3 = arrayOfInputs(36)
    parEndPreparation3 = arrayOfInputs(37)
    parScheduleThickness3 = arrayOfInputs(38)
    parEndStandard3 = arrayOfInputs(39)
    parPressureRating3 = arrayOfInputs(40)
    parFlowDirection3 = arrayOfInputs(41)
    parPortIndex4 = arrayOfInputs(42)
    parNpd4 = arrayOfInputs(43)
    parEndPreparation4 = arrayOfInputs(44)
    parScheduleThickness4 = arrayOfInputs(45)
    parEndStandard4 = arrayOfInputs(46)
    parPressureRating4 = arrayOfInputs(47)
    parFlowDirection4 = arrayOfInputs(48)
    
    parInsulationThickness = arrayOfInputs(49)
    
    'Text Inputs
    parId1 = arrayOfInputs(50)
    parId2 = arrayOfInputs(51)
    parId3 = arrayOfInputs(52)
    parId4 = arrayOfInputs(53)
    parNPDUnitType = arrayOfInputs(54)
    parNpdUnitType1 = arrayOfInputs(55)
    parNpdUnitType2 = arrayOfInputs(56)
    parNpdUnitType3 = arrayOfInputs(57)
    parNpdUnitType4 = arrayOfInputs(58)
    
    iOutput = 0
'The orgin is considered to be along the flow axis and mid of parInstrumentLength2
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

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNpdUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    
    If parInstrumentLength4 > flangeThick(parPortIndex1) Then
    oNozzle.Length = parInstrumentLength4
    Else
        oNozzle.Length = flangeThick(parPortIndex1)
    End If

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength - _
                    parInstrumentLength4 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex2, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNpdUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

    If parInstrumentLength5 > flangeThick(parPortIndex2) Then
    oNozzle.Length = parInstrumentLength5
    Else
        oNozzle.Length = flangeThick(parPortIndex2)
    End If

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parInstrumentLength2 / 2 + parInstrumentLength3 + parInstrumentLength5 + _
                    sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 3
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNpdUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)

    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

    If parNozzleOffset2 > flangeThick(parPortIndex3) Then
    oNozzle.Length = parNozzleOffset2
    Else
        oNozzle.Length = flangeThick(parPortIndex3)
    End If

' Direction of the Nozzle
    oDir.Set 0, Sin(parPortRotation1), Cos(parPortRotation1)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -(parNozzleOffset1 - parNozzleOffset), _
                                (parNozzleOffset2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Sin(parPortRotation1), _
                                (parNozzleOffset2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Cos(parPortRotation1)
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 4
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNpdUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle

    If parNozzleOffset3 > flangeThick(parPortIndex4) Then
    oNozzle.Length = parNozzleOffset3
    Else
        oNozzle.Length = flangeThick(parPortIndex4)
    End If

' Direction of the Nozzle
    oDir.Set 0, Sin(parPortRotation2), Cos(parPortRotation2)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, (parNozzleOffset3 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Sin(parPortRotation2), _
                                    (parNozzleOffset3 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Cos(parPortRotation2)
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
 ' Insert your code for output 5(Cylinder Body1)
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength, 0, 0
    enPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1, 0, 0
    Set ObjCylinderBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1), True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinderBody1
    Set ObjCylinderBody1 = Nothing
    
' Insert your code for output 6(Cylinder Body2)
    stPoint.Set parInstrumentLength2 / 2, 0, 0
    enPoint.Set -parInstrumentLength2 / 2, 0, 0
    Set ObjCylinderBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinderBody2
    Set ObjCylinderBody2 = Nothing

 ' Insert your code for output 7(Reducing Cone)
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1, 0, 0
    enPoint.Set -parInstrumentLength2 / 2, 0, 0
    Set ObjReducingCone = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, parInstrumentDiameter / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducingCone
    Set ObjReducingCone = Nothing

 ' Insert your code for output 8(Diverging Cone)
    stPoint.Set parInstrumentLength2 / 2 + parInstrumentLength3, 0, 0
    enPoint.Set parInstrumentLength2 / 2, 0, 0
    Set ObjDivergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, parInstrumentDiameter / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

