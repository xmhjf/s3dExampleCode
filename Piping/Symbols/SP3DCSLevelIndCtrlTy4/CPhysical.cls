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
'   Author:          MS
'   Creation Date:  Tuesday, Jan 21 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This symbol details are available in PDS Piping Component Data Reference Manual. SN=I324DZ
' Page No D-126
'This Symbol only consist of Physical aspect and is made up of Vertical Cylinder, Two Arms of Cylinder
'one Indicator of Box and Three nozzles with length, one along positive Z , second along negative Z and third along
' positive X axis.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
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
    Dim ObjFloatBody As Object
    Dim ObjArm1 As Object
    Dim ObjArm2 As Object
    Dim ObjIndicator As Object
    
    Dim parNozzle3toCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFloatDiameter As Double
    Dim parFloatToptoCenter As Double
    Dim parFloatBottoCenter As Double
    Dim parArmtoCenter As Double
    Dim parIndicatorHeight1 As Double
    Dim parIndicatorHeight2 As Double
    Dim parAngle As Double
    Dim parArm1Length As Double
    Dim parArm2Length As Double
    Dim parIndicatorWidth As Double
    Dim parIndicatorLength As Double
    Dim ArmDiameter As Double
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
    parFace3toCenter = arrayOfInputs(2)         'P2
    parNozzle3toCenter = arrayOfInputs(3)   'P3
    parFace1toCenter = arrayOfInputs(4)         'P4
    parFace2toCenter = arrayOfInputs(5)         'P5
    parFloatDiameter = arrayOfInputs(6)         'P6
    parFloatToptoCenter = arrayOfInputs(7)      'P7
    parFloatBottoCenter = arrayOfInputs(8)      'P8
    parArmtoCenter = arrayOfInputs(9)           'P9
    parIndicatorHeight1 = arrayOfInputs(10)           'P10
    parIndicatorHeight2 = arrayOfInputs(11)         'P11
    parAngle = arrayOfInputs(12)                    'P12
    parArm1Length = arrayOfInputs(13)           'P13
    parArm2Length = arrayOfInputs(14)           'P14
    parIndicatorWidth = arrayOfInputs(15)       'P15
    parIndicatorLength = arrayOfInputs(16)      'P16
    parNPD = arrayOfInputs(17)
    parEndPreparation = arrayOfInputs(18)
    parScheduleThickness = arrayOfInputs(19)
    parEndStandard = arrayOfInputs(20)
    parPressureRating = arrayOfInputs(21)
    parFlowDirection = arrayOfInputs(22)
    parPortIndex1 = arrayOfInputs(23)
    parNpd1 = arrayOfInputs(24)
    parEndPreparation1 = arrayOfInputs(25)
    parScheduleThickness1 = arrayOfInputs(26)
    parEndStandard1 = arrayOfInputs(27)
    parPressureRating1 = arrayOfInputs(28)
    parFlowDirection1 = arrayOfInputs(29)
    parPortIndex2 = arrayOfInputs(30)
    parNpd2 = arrayOfInputs(31)
    parEndPreparation2 = arrayOfInputs(32)
    parScheduleThickness2 = arrayOfInputs(33)
    parEndStandard2 = arrayOfInputs(34)
    parPressureRating2 = arrayOfInputs(35)
    parFlowDirection2 = arrayOfInputs(36)
    parPortIndex3 = arrayOfInputs(37)
    parNpd3 = arrayOfInputs(38)
    parEndPreparation3 = arrayOfInputs(39)
    parScheduleThickness3 = arrayOfInputs(40)
    parEndStandard3 = arrayOfInputs(41)
    parPressureRating3 = arrayOfInputs(42)
    parFlowDirection3 = arrayOfInputs(43)
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parId3 = arrayOfInputs(46)
    parNPDUnitType = arrayOfInputs(47)
    parNPDUnitType1 = arrayOfInputs(48)
    parNPDUnitType2 = arrayOfInputs(49)
    parNPDUnitType3 = arrayOfInputs(50)
    
    iOutput = 0
    ArmDiameter = 0.05
 ' Insert your code for output 1(FloatBody)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, -parFloatBottoCenter, 0
    endPoint.Set 0, parFloatToptoCenter, 0
    Set ObjFloatBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, parFloatDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFloatBody
    Set ObjFloatBody = Nothing
    
 ' Insert your code for output 2(Arm1)
    Dim Arm1Angle As Double
    Arm1Angle = PI + parAngle
    startPoint.Set 0, parArmtoCenter, 0
    endPoint.Set (parArm1Length + ArmDiameter / 2) * Cos(Arm1Angle), _
                            parArmtoCenter, _
                            -(parArm1Length + ArmDiameter / 2) * Sin(Arm1Angle)
                            
    Set ObjArm1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, ArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm1
    Set ObjArm1 = Nothing
    
 ' Insert your code for output 3(Arm2)
    Dim Arm2Angle As Double
    Arm2Angle = 1.5 * PI + parAngle
    startPoint.Set parArm1Length * Cos(Arm1Angle), _
                            parArmtoCenter, _
                            -parArm1Length * Sin(Arm1Angle)
    endPoint.Set parArm1Length * Cos(Arm1Angle) + (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Sin(parAngle), _
                            parArmtoCenter, _
                            -parArm1Length * Sin(Arm1Angle) + (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Cos(parAngle)
    Set ObjArm2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, ArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm2
    Set ObjArm2 = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
 ' Insert your code for output 4(Indicator)
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLineStrPts(0 To 14) As Double
    Dim indiAngle As Double
    Dim projVect As AutoMath.DVector
    Set projVect = New AutoMath.DVector
    Dim geometryfactory As geometryfactory
    Set geometryfactory = New geometryfactory
    
    indiAngle = 2 * PI + parAngle
    projVect.Set Cos(Arm2Angle), 0, -Sin(Arm2Angle)
    Dim endx As Double
    Dim endy As Double
    endx = parArm1Length * Cos(Arm1Angle) + (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Sin(parAngle)
    endy = parArm1Length * Sin(Arm1Angle) - (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Cos(parAngle)

   'Point 1
    dblLineStrPts(0) = endx + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(1) = parArmtoCenter + parIndicatorHeight1
    dblLineStrPts(2) = -endy - Sin(indiAngle) * parIndicatorLength / 2
    'Point 2
    dblLineStrPts(3) = endx + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(4) = parArmtoCenter - parIndicatorHeight2
    dblLineStrPts(5) = -endy - Sin(indiAngle) * parIndicatorLength / 2
    'Point 3
    dblLineStrPts(6) = endx - Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(7) = parArmtoCenter - parIndicatorHeight2
    dblLineStrPts(8) = -endy + Sin(indiAngle) * parIndicatorLength / 2
    'Point 4
    dblLineStrPts(9) = endx - Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(10) = parArmtoCenter + parIndicatorHeight1
    dblLineStrPts(11) = -endy + Sin(indiAngle) * parIndicatorLength / 2
    'Point 5
    dblLineStrPts(12) = endx + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(13) = parArmtoCenter + parIndicatorHeight1
    dblLineStrPts(14) = -endy - Sin(indiAngle) * parIndicatorLength / 2
    
    Set objLineStr = geometryfactory.LineStrings3d.CreateByPoints(Nothing, 5, dblLineStrPts)
    
    Set ObjIndicator = PlaceProjection(m_OutputColl, objLineStr, projVect, parIndicatorWidth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIndicator
    Set ObjIndicator = Nothing
    Set projVect = Nothing
    Dim objTmp As IJDObject
    Set objTmp = objLineStr
    Set objLineStr = Nothing
    objTmp.Remove
    Set geometryfactory = Nothing
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
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
    oNozzle.Length = parFace1toCenter - parFloatBottoCenter

' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0
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
    oNozzle.Length = parFace2toCenter - parFloatToptoCenter

' Direction of the Nozzle
    oDir.Set 0, 1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, parFace2toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0
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
    oNozzle.Length = parFace3toCenter

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace3toCenter + sptOffset(parPortIndex3) - depth(parPortIndex3), parNozzle3toCenter, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
        
End Sub
