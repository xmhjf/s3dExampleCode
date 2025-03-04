VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          ACM
'   Creation Date:  Tuesday, May 16, 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'   This is an Exchanger Shell Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is a Dow Equipment E-161 Exchanger Shell symbol.

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim pipeDiam(1 To 5) As Double
    Dim sptOffset(1 To 5) As Double
    Dim flangeDiam(1 To 5)    As Double
    Dim depth(1 To 5) As Double
    Dim flangeThick(1 To 5) As Double
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parCPtoFace As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parNozzlePosition As Double
    Dim parNozzleOrientation As Double
    Dim parRadialPosition As Double
    Dim parOffset As Double
    Dim parNozzlePosition1 As Double
    Dim parNozzleOrientation1 As Double
    Dim parRadialPosition1 As Double
    Dim parOffset1 As Double
    Dim parNozzlePosition2 As Double
    Dim parNozzleOrientation2 As Double
    Dim parRadialPosition2 As Double
    Dim parOffset2 As Double
    Dim parNozzlePosition3 As Double
    Dim parNozzleOrientation3 As Double
    Dim parRadialPosition3 As Double
    Dim parOffset3 As Double
    Dim parNozzlePosition4 As Double
    Dim parNozzleOrientation4 As Double
    Dim parRadialPosition4 As Double
    Dim parOffset4 As Double
    
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
    
    Dim parPortIndex5              As Long
    Dim parNpd5                    As Double
    Dim parEndPreparation5         As Long
    Dim parScheduleThickness5      As Long
    Dim parEndStandard5            As Long
    Dim parPressureRating5         As Long
    Dim parFlowDirection5          As DistribFlow
    
    Dim parId1                     As String
    Dim parId2                     As String
    Dim parId3                     As String
    Dim parId4                     As String
    Dim parId5                     As String
    Dim parNPDUnitType             As String
    Dim parNpdUnitType1            As String
    Dim parNpdUnitType2            As String
    Dim parNpdUnitType3            As String
    Dim parNpdUnitType4            As String
    Dim parNpdUnitType5            As String
        
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselHeight = arrayOfInputs(3)
    parCPtoFace = arrayOfInputs(4)
    parFaceDiameter = arrayOfInputs(5)
    parFaceThickness = arrayOfInputs(6)
    parNozzlePosition = arrayOfInputs(7)
    parNozzleOrientation = arrayOfInputs(8)
    parRadialPosition = arrayOfInputs(9)
    parOffset = arrayOfInputs(10)
    parNozzlePosition1 = arrayOfInputs(11)
    parNozzleOrientation1 = arrayOfInputs(12)
    parRadialPosition1 = arrayOfInputs(13)
    parOffset1 = arrayOfInputs(14)
    parNozzlePosition2 = arrayOfInputs(15)
    parNozzleOrientation2 = arrayOfInputs(16)
    parRadialPosition2 = arrayOfInputs(17)
    parOffset2 = arrayOfInputs(18)
    parNozzlePosition3 = arrayOfInputs(19)
    parNozzleOrientation3 = arrayOfInputs(20)
    parRadialPosition3 = arrayOfInputs(21)
    parOffset3 = arrayOfInputs(22)
    parNozzlePosition4 = arrayOfInputs(23)
    parNozzleOrientation4 = arrayOfInputs(24)
    parRadialPosition4 = arrayOfInputs(25)
    parOffset4 = arrayOfInputs(26)

    parNpd = arrayOfInputs(27)
    parEndPreparation = arrayOfInputs(28)
    parScheduleThickness = arrayOfInputs(29)
    parEndStandard = arrayOfInputs(30)
    parPressureRating = arrayOfInputs(31)
    parFlowDirection = arrayOfInputs(32)
    parPortIndex1 = arrayOfInputs(33)
    parNpd1 = arrayOfInputs(34)
    parEndPreparation1 = arrayOfInputs(35)
    parScheduleThickness1 = arrayOfInputs(36)
    parEndStandard1 = arrayOfInputs(37)
    parPressureRating1 = arrayOfInputs(38)
    parFlowDirection1 = arrayOfInputs(39)
    parPortIndex2 = arrayOfInputs(40)
    parNpd2 = arrayOfInputs(41)
    parEndPreparation2 = arrayOfInputs(42)
    parScheduleThickness2 = arrayOfInputs(43)
    parEndStandard2 = arrayOfInputs(44)
    parPressureRating2 = arrayOfInputs(45)
    parFlowDirection2 = arrayOfInputs(46)
    parPortIndex3 = arrayOfInputs(47)
    parNpd3 = arrayOfInputs(48)
    parEndPreparation3 = arrayOfInputs(49)
    parScheduleThickness3 = arrayOfInputs(50)
    parEndStandard3 = arrayOfInputs(51)
    parPressureRating3 = arrayOfInputs(52)
    parFlowDirection3 = arrayOfInputs(53)
    parPortIndex4 = arrayOfInputs(54)
    parNpd4 = arrayOfInputs(55)
    parEndPreparation4 = arrayOfInputs(56)
    parScheduleThickness4 = arrayOfInputs(57)
    parEndStandard4 = arrayOfInputs(58)
    parPressureRating4 = arrayOfInputs(59)
    parFlowDirection4 = arrayOfInputs(60)
    parPortIndex5 = arrayOfInputs(61)
    parNpd5 = arrayOfInputs(62)
    parEndPreparation5 = arrayOfInputs(63)
    parScheduleThickness5 = arrayOfInputs(64)
    parEndStandard5 = arrayOfInputs(65)
    parPressureRating5 = arrayOfInputs(66)
    parFlowDirection5 = arrayOfInputs(67)
    
    parId1 = arrayOfInputs(68)
    parId2 = arrayOfInputs(69)
    parId3 = arrayOfInputs(70)
    parId4 = arrayOfInputs(71)
    parId5 = arrayOfInputs(72)
    parNPDUnitType = arrayOfInputs(73)
    parNpdUnitType1 = arrayOfInputs(74)
    parNpdUnitType2 = arrayOfInputs(75)
    parNpdUnitType3 = arrayOfInputs(76)
    parNpdUnitType4 = arrayOfInputs(77)
    parNpdUnitType5 = arrayOfInputs(78)

    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim dActualVesselDiameter As Double
    Dim dDomeHeight As Double

''Assumptions
    'The input parameter for Vessel Diameter is Internal diameter .We need to take care for Thickness of Body.
    dActualVesselDiameter = parVesselDiameter + 0.0508   ' 2 Inch
    dDomeHeight = dActualVesselDiameter / 4  '2:1 ELiptical Head (TYP)
    
' Insert your code for output 1(Shell Flange)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace
    oEnPoint.Set 0, 0, -parCPtoFace - parFaceThickness
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter
    
' Insert your code for output 2(Cylindrical Exchanger Shell Body)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace - parFaceThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dActualVesselDiameter

'Insert your code for output 3(Shell Bottom Flange)

    oStPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter
    
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
' Place Nozzle 1 (C)
    Dim oPlacePoint As New AutoMath.DPosition
    Dim oDir        As New AutoMath.DVector
    Dim objNozzle   As IJDNozzle
    Dim dNozzleLength As Double
    Dim dMeetPoint As Double
    Dim dThirdSide As Double
    Dim dAdjustableNozzLength As Double

    Dim oLogicalDistPort    As IJLogicalDistPort
    Dim oDistribPort        As IJDistribPort

    Set objNozzle = CreateRetrieveDynamicNozzlePHForEquipment(parPortIndex1, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNpdUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = objNozzle
    Set oDistribPort = objNozzle
    
    dMeetPoint = Sqr(((dActualVesselDiameter / 2) ^ 2) - ((pipeDiam(1) / 2) ^ 2))
    dNozzleLength = parRadialPosition - dMeetPoint

    If (Abs(parOffset) + pipeDiam(1) / 2) > dActualVesselDiameter / 2 Then   'Nozzle is Outside the Vessel
        dAdjustableNozzLength = flangeThick(1)
    Else 'Nozzle is touching the Vessel
        dThirdSide = Sqr(Abs((dActualVesselDiameter / 2) ^ 2) - ((parOffset) + (pipeDiam(1) / 2)) ^ 2)
        If dNozzleLength <= 0.1 Then
            dAdjustableNozzLength = flangeThick(1)
        Else
            dAdjustableNozzLength = dNozzleLength + dActualVesselDiameter / 2 - dThirdSide
        End If
    End If
    objNozzle.Length = dAdjustableNozzLength
' Direction of the Nozzle
    oDir.Set Sin(parNozzleOrientation), Cos(parNozzleOrientation), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
     oPlacePoint.Set (parRadialPosition + sptOffset(1) - depth(1)) * Sin(parNozzleOrientation) + (parOffset1 * Cos(parNozzleOrientation1)), _
                     (parRadialPosition + sptOffset(1) - depth(1)) * Cos(parNozzleOrientation) + -(parOffset1 * Sin(parNozzleOrientation1)), _
                     -parNozzlePosition
     oLogicalDistPort.SetCenterLocation oPlacePoint
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 (Nozzle D-Mounted on shell body)
    Set objNozzle = CreateRetrieveDynamicNozzlePHForEquipment(parPortIndex2, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNpdUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = objNozzle
    Set oDistribPort = objNozzle
    
    dMeetPoint = Sqr(((dActualVesselDiameter / 2) ^ 2) - ((pipeDiam(2) / 2) ^ 2))
    dNozzleLength = parRadialPosition1 - dMeetPoint

    If (Abs(parOffset1) + pipeDiam(2) / 2) > dActualVesselDiameter / 2 Then   'Nozzle is Outside the Vessel
        dAdjustableNozzLength = flangeThick(2)
    Else 'Nozzle is touching the Vessel
        dThirdSide = Sqr(Abs((dActualVesselDiameter / 2) ^ 2) - ((parOffset1) + (pipeDiam(2) / 2)) ^ 2)
        If dNozzleLength <= 0.1 Then
            dAdjustableNozzLength = flangeThick(2)
        Else
            dAdjustableNozzLength = dNozzleLength + dActualVesselDiameter / 2 - dThirdSide
        End If
    End If
    objNozzle.Length = dAdjustableNozzLength
' Direction of the Nozzle
   oDir.Set Sin(parNozzleOrientation1), Cos(parNozzleOrientation1), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parRadialPosition1 + sptOffset(2) - depth(2)) * Sin(parNozzleOrientation1) + (parOffset1 * Cos(parNozzleOrientation1)), _
                    (parRadialPosition1 + sptOffset(2) - depth(2)) * Cos(parNozzleOrientation1) + (parOffset1 * Sin(parNozzleOrientation1)), _
                    -parNozzlePosition1
     
     oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3 (B1)
    Set objNozzle = CreateRetrieveDynamicNozzlePHForEquipment(parPortIndex3, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNpdUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = objNozzle
    Set oDistribPort = objNozzle
    
    dMeetPoint = Sqr(((dActualVesselDiameter / 2) ^ 2) - ((pipeDiam(3) / 2) ^ 2))
    dNozzleLength = parRadialPosition2 - dMeetPoint

    If (Abs(parOffset2) + pipeDiam(3) / 2) > dActualVesselDiameter / 2 Then   'Nozzle is Outside the Vessel
        dAdjustableNozzLength = flangeThick(3)
    Else 'Nozzle is touching the Vessel
        dThirdSide = Sqr(Abs((dActualVesselDiameter / 2) ^ 2) - ((parOffset2) + (pipeDiam(3) / 2)) ^ 2)
        If dNozzleLength <= 0.1 Then
            dAdjustableNozzLength = flangeThick(3)
        Else
            dAdjustableNozzLength = dNozzleLength + dActualVesselDiameter / 2 - dThirdSide
        End If
    End If
    objNozzle.Length = dAdjustableNozzLength
' Direction of the Nozzle
    oDir.Set Sin(parNozzleOrientation2), Cos(parNozzleOrientation2), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parRadialPosition2 + sptOffset(3) - depth(3)) * Sin(parNozzleOrientation2) + (parOffset1 * Cos(parNozzleOrientation1)), _
                    (parRadialPosition2 + sptOffset(3) - depth(3)) * Cos(parNozzleOrientation2) + (parOffset1 * Sin(parNozzleOrientation1)), _
                    -parNozzlePosition2
     
    oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 4 (Nozzle F2-Mounted on Top Flange)
    Set objNozzle = CreateRetrieveDynamicNozzlePHForEquipment(parPortIndex4, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNpdUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = objNozzle
    Set oDistribPort = objNozzle
    
    dMeetPoint = Sqr(((dActualVesselDiameter / 2) ^ 2) - ((pipeDiam(4) / 2) ^ 2))
    dNozzleLength = parRadialPosition3 - dMeetPoint

    If (Abs(parOffset3) + pipeDiam(4) / 2) > dActualVesselDiameter / 2 Then   'Nozzle is Outside the Vessel
        dAdjustableNozzLength = flangeThick(4)
    Else 'Nozzle is touching the Vessel
        dThirdSide = Sqr(Abs((dActualVesselDiameter / 2) ^ 2) - ((parOffset3) + (pipeDiam(4) / 2)) ^ 2)
        If dNozzleLength <= 0.1 Then
            dAdjustableNozzLength = flangeThick(4)
        Else
            dAdjustableNozzLength = dNozzleLength + dActualVesselDiameter / 2 - dThirdSide
        End If
    End If
    objNozzle.Length = dAdjustableNozzLength
' Direction of the Nozzle
    oDir.Set Sin(parNozzleOrientation3), Cos(parNozzleOrientation3), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parRadialPosition3 + sptOffset(4) - depth(4)) * Sin(parNozzleOrientation3) + (parOffset1 * Cos(parNozzleOrientation1)), _
                    (parRadialPosition3 + sptOffset(4) - depth(4)) * Cos(parNozzleOrientation3) + (parOffset1 * Sin(parNozzleOrientation1)), -parNozzlePosition3
     
    oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 5 (Nozzle F3-Mounted on Top Flange)
    Set objNozzle = CreateRetrieveDynamicNozzlePHForEquipment(parPortIndex5, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd5, _
                                            parNpdUnitType5, parEndPreparation5, parScheduleThickness5, _
                                            parEndStandard5, parPressureRating5, parFlowDirection5, _
                                            parId5, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = objNozzle
    Set oDistribPort = objNozzle
    
    dMeetPoint = Sqr(((dActualVesselDiameter / 2) ^ 2) - ((pipeDiam(5) / 2) ^ 2))
    dNozzleLength = parRadialPosition4 - dMeetPoint

    If (Abs(parOffset4) + pipeDiam(5) / 2) > dActualVesselDiameter / 2 Then   'Nozzle is Outside the Vessel
        dAdjustableNozzLength = flangeThick(5)
    Else 'Nozzle is touching the Vessel
        dThirdSide = Sqr(Abs((dActualVesselDiameter / 2) ^ 2) - ((parOffset4) + (pipeDiam(5) / 2)) ^ 2)
        If dNozzleLength <= 0.1 Then
            dAdjustableNozzLength = flangeThick(5)
        Else
        
            dAdjustableNozzLength = dNozzleLength + dActualVesselDiameter / 2 - dThirdSide
        End If
    End If
    objNozzle.Length = dAdjustableNozzLength
' Direction of the Nozzle
    oDir.Set Sin(parNozzleOrientation4), Cos(parNozzleOrientation4), 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parRadialPosition4 + sptOffset(5) - depth(5)) * Sin(parNozzleOrientation4) + (parOffset1 * Cos(parNozzleOrientation1)), _
                     (parRadialPosition4 + sptOffset(5) - depth(5)) * Cos(parNozzleOrientation4) + (parOffset1 * Sin(parNozzleOrientation1)), _
                     -parNozzlePosition4
     
    oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set m_oCodeListMetadata = Nothing
   
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

