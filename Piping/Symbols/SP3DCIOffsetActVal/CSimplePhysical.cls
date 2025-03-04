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
'   Author:          MS
'   Creation Date:  Friday, May 23 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "SimplePhysical:" 'Used for error messages
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjActuatorBody As Object
    Dim parFace1toCenter         As Double  'D2
    Dim parFacetoFace        As Double  'D3
    Dim parActuatorHeight        As Double  'D6
    Dim parActuatorDiameter      As Double  'D7
    Dim parActuatorHeight1       As Double  'D8
    Dim parActuatorLength        As Double  'D9
    Dim parActuatorLength1       As Double  'D10
    Dim parActuatorLength2       As Double  'D11
    Dim parActuatorWidth         As Double  'D13
    Dim parActuatorWidth1        As Double  'D14
    Dim parActuatorHeight2       As Double  'D15

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
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parActuatorHeight = arrayOfInputs(4)
    parActuatorDiameter = arrayOfInputs(5)
    parActuatorHeight1 = arrayOfInputs(6)
    parActuatorLength = arrayOfInputs(7)
    parActuatorLength1 = arrayOfInputs(8)
    parActuatorLength2 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorHeight2 = arrayOfInputs(12)

    parNPD = arrayOfInputs(13)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parPortIndex1 = arrayOfInputs(19)
    parNpd1 = arrayOfInputs(20)
    parEndPreparation1 = arrayOfInputs(21)
    parScheduleThickness1 = arrayOfInputs(22)
    parEndStandard1 = arrayOfInputs(23)
    parPressureRating1 = arrayOfInputs(24)
    parFlowDirection1 = arrayOfInputs(25)
    parPortIndex2 = arrayOfInputs(26)
    parNpd2 = arrayOfInputs(27)
    parEndPreparation2 = arrayOfInputs(28)
    parScheduleThickness2 = arrayOfInputs(29)
    parEndStandard2 = arrayOfInputs(30)
    parPressureRating2 = arrayOfInputs(31)
    parFlowDirection2 = arrayOfInputs(32)
    'Text Inputs
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parNPDUnitType = arrayOfInputs(35)
    parNPDUnitType1 = arrayOfInputs(36)
    parNPDUnitType2 = arrayOfInputs(37)
    
    iOutput = 0
     m_oSymGeomHelper.OutputCollection = m_OutputColl
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
    
' Insert your code for output 3(ConvergingCone)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    ' A value of 0.0000001 is used in Z direction to avoid symbol placement failure(gives assertion errors).
    ' This should be removed  once it is resolved.
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0.0000001
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam(parPortIndex1), 0
     
' Insert your code for output 4(DivergingCone)
    ' A value of 0.0000001 is used in Z direction to avoid symbol placement failure(gives assertion errors).
    ' This should be removed  once it is resolved.
    stPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0.0000001
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam(parPortIndex2), 0
 
' Insert your code for output 5(Actuator Stem1)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight + parActuatorDiameter / 2, 0
   
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parActuatorDiameter
   

' Insert your code for output 6(Actuator Stem2)
    stPoint.Set 0, parActuatorHeight, 0
    enPoint.Set 0, parActuatorHeight, -(parActuatorWidth1 - parActuatorWidth / 2)
   
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parActuatorDiameter
    
' Insert your code for output 7(Actuator Body)
    stPoint.Set -parActuatorLength1 + parActuatorLength2, _
                            parActuatorHeight - (parActuatorHeight1 / 2 - parActuatorHeight2), _
                            -(parActuatorWidth1 - parActuatorWidth / 2)
    enPoint.Set parActuatorLength2, _
                            parActuatorHeight + parActuatorHeight2 + parActuatorHeight1 / 2, _
                            -(parActuatorWidth1 + parActuatorWidth / 2)
    
    Set ObjActuatorBody = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
