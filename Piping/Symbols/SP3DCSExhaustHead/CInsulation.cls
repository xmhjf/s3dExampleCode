VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This Symbol details are available in PDS Piping Component Data Reference Guide Symbol Number SN=S7AZ, Page No D-101
'This Symbol consist of Two Nozzles, a exhaust ( truncated cone), and a down Pipe (Cylinder) as a part of
'Physical Aspect. Insulation for Truncated Cone  Pipe and both the Nozzles
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   11.Apr.2005     svsmylav        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                           Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Sub Class_Initialize()
'''
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
    Dim ObjExConeIns As Object
    Dim ObjNoz1Ins  As Object
    Dim ObjPipeIns As Object
    Dim ObjNoz2Ins  As Object
    
    Dim parExhaustHeight As Double
    Dim parExhaustDiameter As Double
    Dim parPipeOffset As Double
    Dim parPipeHeight As Double
    Dim parInsulationThickness  As Double
    Dim parInsulationDiameter As Double

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
    parExhaustHeight = arrayOfInputs(2) 'P2
    parExhaustDiameter = arrayOfInputs(3)   'p3
    parPipeOffset = arrayOfInputs(4)            'p4
    parPipeHeight = arrayOfInputs(5)        'p5
    parInsulationThickness = arrayOfInputs(6)
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
    
    iOutput = 0

' Insert your code for output 1(Exhaust Cone Insulation)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    Dim baseCenter As AutoMath.DPosition
    Dim topCenter As AutoMath.DPosition
    Set baseCenter = New AutoMath.DPosition
    Set topCenter = New AutoMath.DPosition
    baseCenter.Set 0, flangeThick(parPortIndex1), 0
    topCenter.Set 0, parExhaustHeight + parInsulationThickness, 0
    Set ObjExConeIns = PlaceCone(m_OutputColl, baseCenter, topCenter, pipeDiam(parPortIndex1) / 2 + parInsulationThickness, _
                                                        parExhaustDiameter / 2 + parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExConeIns
    Set ObjExConeIns = Nothing
    Set baseCenter = Nothing
    Set topCenter = Nothing
    
' Insert your code for output 2(Nozzle 1 Insulation)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, 0
    endPoint.Set 0, flangeThick(parPortIndex1) + parInsulationThickness, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        endPoint.Set 0, flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    If pipeDiam(parPortIndex1) > flangeDiam(parPortIndex1) Then parInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
    Set ObjNoz1Ins = Nothing
    
 ' Insert your code for output 2(Down Pipe Insulation)
    startPoint.Set parPipeOffset, parExhaustHeight + parInsulationThickness, 0
    endPoint.Set parPipeOffset, (parExhaustHeight - parPipeHeight) + flangeThick(parPortIndex2), 0
    parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjPipeIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeIns
    Set ObjPipeIns = Nothing
    
' Place Nozzle 2 Insulation
    startPoint.Set parPipeOffset, (parExhaustHeight - parPipeHeight), 0
    endPoint.Set parPipeOffset, (parExhaustHeight - parPipeHeight) + flangeThick(parPortIndex2) + parInsulationThickness, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        endPoint.Set parPipeOffset, (parExhaustHeight - parPipeHeight) + flangeThick(parPortIndex2) + parInsulationThickness + 0.0001, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    If pipeDiam(parPortIndex2) > flangeDiam(parPortIndex2) Then parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
    Set ObjNoz2Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
