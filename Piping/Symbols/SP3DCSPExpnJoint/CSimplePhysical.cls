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
'**********************************************************************************
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author: svsmylav
'   Creation Date:  Friday, Oct 8 2004
'
'   Description:
'     This is PDS on-the-fly S28A  Expansion Joint Symbol.
'
'  Change History:
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const MIN_LENGTH = 0.0001


Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double

    Dim oPart As IJDPart
    Dim oPipeComponent As IJDPipeComponent
    
    Set oPart = arrayOfInputs(1)
'    Set oPipeComponent = oPart
    
'    Dim oNozzleData() As NozzleData
'    NozzleInitialize oPart, oNozzleData
    
'    Dim oNozzleCol As Collection
'    Set oNozzleCol = New Collection
    
    Dim oOutputCol As Collection
    Set oOutputCol = New Collection
    
    Dim oT4x4Current As IJDT4x4
    Set oT4x4Current = New DT4x4
    
    Dim oT4x4Temp As IJDT4x4
    Set oT4x4Temp = New DT4x4
    
    Dim oGeometryFactory As GeometryFactory
    Set oGeometryFactory = New GeometryFactory

    Dim oOriginPos As DPosition
    Set oOriginPos = New DPosition
    oOriginPos.Set 0, 0, 0
    
'**********************************************************************************
'**********************************************************************************
'   BEGIN
'**********************************************************************************
'**********************************************************************************
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parCylinderDiameter As Double
    Dim parNozzleLength As Double
    Dim parNozzleLength1 As Double
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

' Inputs
    parFacetoCenter = arrayOfInputs(2)  'A
    parFace1toCenter = arrayOfInputs(3) 'B
    parCylinderDiameter = arrayOfInputs(4)  'C
    parNozzleLength = arrayOfInputs(5)  'D
    parNozzleLength1 = arrayOfInputs(6)  'E
    parNPD = arrayOfInputs(7)
    parEndPreparation = arrayOfInputs(8)
    parScheduleThickness = arrayOfInputs(9)
    parEndStandard = arrayOfInputs(10)
    parPressureRating = arrayOfInputs(11)
    parFlowDirection = arrayOfInputs(12)
    parPortIndex1 = arrayOfInputs(13)
    parNpd1 = arrayOfInputs(14)
    parEndPreparation1 = arrayOfInputs(15)
    parScheduleThickness1 = arrayOfInputs(16)
    parEndStandard1 = arrayOfInputs(17)
    parPressureRating1 = arrayOfInputs(18)
    parFlowDirection1 = arrayOfInputs(19)
    parPortIndex2 = arrayOfInputs(20)
    parNpd2 = arrayOfInputs(21)
    parEndPreparation2 = arrayOfInputs(22)
    parScheduleThickness2 = arrayOfInputs(23)
    parEndStandard2 = arrayOfInputs(24)
    parPressureRating2 = arrayOfInputs(25)
    parFlowDirection2 = arrayOfInputs(26)
    parId1 = arrayOfInputs(27)
    parId2 = arrayOfInputs(28)
    parNPDUnitType = arrayOfInputs(29)
    parNPDUnitType1 = arrayOfInputs(30)
    parNPDUnitType2 = arrayOfInputs(31)

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

    Dim dNozzleLength As Double

    dNozzleLength = parNozzleLength
    If dNozzleLength < flangeThick(parPortIndex1) Then dNozzleLength = flangeThick(parPortIndex1)
    oNozzle.Length = dNozzleLength

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFacetoCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    oOutputCol.Add oNozzle
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

    dNozzleLength = parNozzleLength1
    If dNozzleLength < flangeThick(parPortIndex2) Then dNozzleLength = flangeThick(parPortIndex2)
    oNozzle.Length = dNozzleLength

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    oOutputCol.Add oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
 ' Insert your code for output 3(oCylinderCapped1)
'    Ensure that cylinder length (sec1) is greater than zero
    Dim sec As Double

    sec = parFacetoCenter + parFace1toCenter - parNozzleLength - parNozzleLength1
    If sec < MIN_LENGTH Then sec = MIN_LENGTH
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parFacetoCenter + parNozzleLength, 0, 0
    enPoint.Set stPoint.x + sec, stPoint.y, stPoint.z
    
    Dim oCylinderCapped1 As Object
    Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parCylinderDiameter, True)
    
    oOutputCol.Add oCylinderCapped1

    Dim lOutputIndex As Long
    'add all outputs
    For lOutputIndex = 1 To oOutputCol.Count
        m_OutputColl.AddOutput arrayOfOutputs(lOutputIndex), oOutputCol.Item(lOutputIndex)
    Next lOutputIndex
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

