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
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly Description="S1A2  Free Vent without Screen Symbol."
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI           As Double
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
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjFreeeVentBody As Object
    
    Dim parFacetoEnd As Double
    Dim parCylinderDiameter As Double

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
    
' Inputs
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parCylinderDiameter = arrayOfInputs(3)
    parNPD = arrayOfInputs(4)
    parEndPreparation = arrayOfInputs(5)
    parScheduleThickness = arrayOfInputs(6)
    parEndStandard = arrayOfInputs(7)
    parPressureRating = arrayOfInputs(8)
    parFlowDirection = arrayOfInputs(9)
    parPortIndex1 = arrayOfInputs(10)
    parNpd1 = arrayOfInputs(11)
    parEndPreparation1 = arrayOfInputs(12)
    parScheduleThickness1 = arrayOfInputs(13)
    parEndStandard1 = arrayOfInputs(14)
    parPressureRating1 = arrayOfInputs(15)
    parFlowDirection1 = arrayOfInputs(16)
    parId1 = arrayOfInputs(17)
    parNPDUnitType = arrayOfInputs(18)
    parNPDUnitType1 = arrayOfInputs(19)
    
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
    oNozzle.Length = 3 * flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    Dim sec As Double
    sec = parFacetoEnd - 3 * flangeThick(parPortIndex1)
    oPlacePoint.Set -(3 * flangeThick(parPortIndex1) + sec / 2) - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
 ' Insert your code for output 1(FreeeVentBody)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -sec / 2, 0, 0
    enPoint.Set sec / 2, 0, 0
    Set ObjFreeeVentBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parCylinderDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFreeeVentBody
    Set ObjFreeeVentBody = Nothing
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
