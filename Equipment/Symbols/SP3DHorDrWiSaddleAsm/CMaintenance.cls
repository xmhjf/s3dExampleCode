VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Monday, May 26 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol, Davit for Bottom Horizontal Cover, is prepared as per DOW Chemical Company manual.
'   Dimensions are considered for two pipe size viz., 20" and 24" . All assumed values are as per the DOW
'   documents unless otherwise specified.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Maintenance:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Dim PI As Double

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim pipeDiam(1 To 1)        As Double
    Dim flangeThick(1 To 1)     As Double
    Dim sptOffset(1 To 1)       As Double
    Dim flangeDiam(1 To 1)      As Double
    Dim depth(1 To 1)           As Double

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
    Dim parNpdUnitType1 As String
    Dim parNozzlePosition As Double
    Dim parDomeHeight1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozzlePosition = arrayOfInputs(2)
    parDomeHeight1 = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parId1 = arrayOfInputs(18)
    parNPDUnitType = arrayOfInputs(19)
    parNpdUnitType1 = arrayOfInputs(20)

    m_oGeomHelper.OutputCollection = m_OutputColl
    iOutput = 0
 
    'Place a Cylinder of Diameter twice the Flange diameter.
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim envelopeDia As Double
    Dim clearance As Double
    Dim flangeGap      As Double
    Dim davitRotationAngle As Double
    Dim davitSupDepth As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    davitRotationAngle = PI / 2
    davitSupDepth = 0.178
    'Retriving the data
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNpdUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    'As per DOW document for 20" pipe value from Flange edge to Davit pin axis (F) is 2.5" and for 24" pipe
    'with 300 pressure 3"
    Dim pinDist As Double
    If pipeDiam(parPortIndex1) = 600 And parPressureRating1 = 300 Then
        pinDist = 0.076
    Else
        pinDist = 0.064
    End If
    
    clearance = 0.1
    flangeGap = 0.004
    envelopeDia = flangeDiam(parPortIndex1) * 2 + pinDist * 2 + clearance * 2
    startPoint.Set Sin(davitRotationAngle) * (flangeDiam(parPortIndex1) / 2 + pinDist), _
                            Cos(davitRotationAngle) * (flangeDiam(parPortIndex1) / 2 + pinDist), _
                            -flangeGap
    endPoint.Set Sin(davitRotationAngle) * (flangeDiam(parPortIndex1) / 2 + pinDist), _
                            Cos(davitRotationAngle) * (flangeDiam(parPortIndex1) / 2 + pinDist), _
                            -flangeGap - (flangeThick(parPortIndex1)) / 3 - davitSupDepth
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), startPoint, endPoint, envelopeDia
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


