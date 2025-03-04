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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Tuesday, Jan 21 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No.: I324CZ, Page No. D-126 of PDS Piping Component Data Reference Guide.
' This Symbol is created with Six Outputs, all of which are physical aspect outputs.
'    ObjFloatBody, ObjArm1, ObjArm2 using 'PlaceCylinder' function, and ObjIndicator using 'CreateProjection'.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols.
'   26.Feb.2006     dkl  TR 111053 Rectiifed the implementation of moving the indicator to place correctly
'                                  for 90 and 270 degrees.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim ObjFloatBody As Object
    Dim ObjArm1 As Object
    Dim ObjArm2 As Object
    Dim ObjIndicator As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
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
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'P2
    parFace2toCenter = arrayOfInputs(3) 'P3
    parFloatDiameter = arrayOfInputs(4) 'P4
    parFloatToptoCenter = arrayOfInputs(5)  'P5
    parFloatBottoCenter = arrayOfInputs(6)     'P6
    parArmtoCenter = arrayOfInputs(7)   'P7
    parIndicatorHeight1 = arrayOfInputs(8)  'P8
    parIndicatorHeight2 = arrayOfInputs(9)  'P9
    parAngle = arrayOfInputs(10)    'P10
    parArm1Length = arrayOfInputs(11)   'P11
    parArm2Length = arrayOfInputs(12)   'P12
    parIndicatorWidth = arrayOfInputs(13)   'P13
    parIndicatorLength = arrayOfInputs(14)  'P14
    parNPD = arrayOfInputs(15)
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
    parId1 = arrayOfInputs(35)
    parId2 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)

    iOutput = 0

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
 ' Insert your code for output 1(FloatBody)
     stPoint.Set 0, -parFloatBottoCenter, 0
    enPoint.Set 0, parFloatToptoCenter, 0
    Set ObjFloatBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFloatDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFloatBody
    Set ObjFloatBody = Nothing

' The following matrix implementation is used to provide functionality to rotate the
' entire indicator arm.
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    Set oAboutVec = New DVector
    oAboutVec.Set 0, 1, 0
    oTransMat.LoadIdentity
    oTransMat.Rotate parAngle, oAboutVec
    
 ' Insert your code for output 2(Arm1)
    Dim CylinderMeetPoint As Double
    Dim Arm1Dia As Double
    Dim Arm2Dia As Double
    Arm1Dia = 0.05
    Arm2Dia = 0.05

    CylinderMeetPoint = Sqr(((parFloatDiameter / 2) * (parFloatDiameter / 2)) - ((Arm1Dia / 2) * (Arm1Dia / 2)))
    
    stPoint.x = parArm1Length + Arm2Dia / 2
    stPoint.y = parArmtoCenter
    stPoint.z = 0
    
    enPoint.x = CylinderMeetPoint
    enPoint.y = stPoint.y
    enPoint.z = 0
    
    Set ObjArm1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, Arm1Dia, True)
    
    ObjArm1.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm1
    Set ObjArm1 = Nothing

' Insert your code for output 3(Arm2)
    stPoint.x = parArm1Length
    stPoint.y = parArmtoCenter
    stPoint.z = 0
    
    enPoint.x = stPoint.x
    enPoint.y = stPoint.y
    enPoint.z = parArm2Length - parIndicatorWidth / 2
    
    Set ObjArm2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, Arm2Dia, True)
    
    ObjArm2.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm2
    Set ObjArm2 = Nothing
        
 ' Insert your code for output 4(Indicator)
    Dim linePts(0 To 14) As Double
    Dim optLine As IngrGeom3D.LineString3d
    Dim projVect As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
       
    'Point 1
    linePts(0) = parArm1Length - parIndicatorLength / 2
    linePts(1) = parArmtoCenter + parIndicatorHeight1
    linePts(2) = parArm2Length - parIndicatorWidth / 2
    'Point 2
    linePts(3) = linePts(0)
    linePts(4) = linePts(1)
    linePts(5) = parArm2Length + parIndicatorWidth / 2
    'Point 3
    linePts(6) = parArm1Length + parIndicatorLength / 2
    linePts(7) = linePts(1)
    linePts(8) = linePts(5)
    'Point 4
    linePts(9) = linePts(6)
    linePts(10) = linePts(1)
    linePts(11) = linePts(2)
    'Point 5
    linePts(12) = linePts(0)
    linePts(13) = linePts(1)
    linePts(14) = linePts(2)
    
    projVect.Set 0, -1, 0
    Set optLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, linePts)
    
    Set ObjIndicator = PlaceProjection(m_OutputColl, optLine, projVect, parIndicatorHeight1 + parIndicatorHeight2, True)
    
    ObjIndicator.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIndicator
    Set ObjIndicator = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set projVect = Nothing
    Set geomFactory = Nothing
    Set oTransMat = Nothing
    Set oAboutVec = Nothing
    
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
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
