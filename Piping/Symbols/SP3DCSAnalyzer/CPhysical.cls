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
' Symbol Model No.: I301AZ, Page No. D-125 PDS Piping Component Data Reference Guide.
' Symbol is created with two Outputs which are created as follows:
'   ObjBody - Using 'PlaceCylinder' function and
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   21.Jul.2003      ACM                             TR 44555 Put a check for nozzle length for port,such that it should'nt fail when Nozzle Length is less than FlangeThickness.
'                                                                  Code is modified to place Nozzle from the axis(center).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)     As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object
    
    Dim parAnalyzerThickness As Double
    Dim parAnalyzerDiameter As Double
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
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parAnalyzerThickness = arrayOfInputs(2)
    parAnalyzerDiameter = arrayOfInputs(3)
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

' Insert your code for output 1(Body of Analyzer)
    stPoint.Set 0, -parAnalyzerThickness / 2, 0
    enPoint.Set 0, parAnalyzerThickness / 2, 0
    Set ObjBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parAnalyzerDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
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
    
''Nozzle is placed from center.i.e Nozzle length is parAnalyzerDiameter /2
''This check helps to avoid failure of nozzle, if length is less than the flangeThick(parPortIndex1)

    If parAnalyzerDiameter / 2 > flangeThick(parPortIndex1) Then
        oNozzle.Length = parAnalyzerDiameter / 2
    Else
    oNozzle.Length = flangeThick(parPortIndex1)
    End If

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parAnalyzerDiameter / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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


