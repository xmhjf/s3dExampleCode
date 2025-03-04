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
'   Author:          SSP
'   Creation Date:  Friday, Jan 17 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: S49AC, Page No. D-105 of PDS Piping Component Data Reference Guide.
'    This Symbol is created with Thirteen Outputs, Five of which are physical aspect outputs.
'    Branch Body, Strainer Body using 'PlaceCylinder' function,
'    Three ObjNozzle object by using 'CreateNozzle' function.
'    And all the Insulation aspect ouputs using 'PlaceCylinder' function.
'    The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   14.Sep.2004     MS     TR 63296 Corrected the arrayofInputs index number for PressureRating of Port1
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
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Double
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object
    Dim ObjPipeinNoz1Ins As Object
    Dim ObjPipeinNoz2Ins As Object
    Dim ObjPipeinNoz3Ins As Object
    Dim ObjNoz1Ins As Object
    Dim ObjNoz2Ins As Object
    Dim ObjNoz3Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parStrToptoCenter As Double
    Dim parStrDiameter As Double
    Dim parBranchDiameter As Double
    Dim parStrWidth As Double
    Dim parFace3toStrTop As Double
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
    parFace1toCenter = arrayOfInputs(2)     'p2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parStrToptoCenter = arrayOfInputs(4)    'P4
    parStrDiameter = arrayOfInputs(5)             'P5
    parBranchDiameter = arrayOfInputs(6)             'P6
    parStrWidth = arrayOfInputs(7)                'P7
    parFace3toStrTop = arrayOfInputs(8)        'P8
    parInsulationThickness = arrayOfInputs(9)
    parNPD = arrayOfInputs(10)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(11)
    parScheduleThickness = arrayOfInputs(12)
    parEndStandard = arrayOfInputs(13)
    parPressureRating = arrayOfInputs(14)
    parFlowDirection = arrayOfInputs(15)
    parPortIndex1 = arrayOfInputs(16)
    parNpd1 = arrayOfInputs(17)
    parEndPreparation1 = arrayOfInputs(18)
    parScheduleThickness1 = arrayOfInputs(19)
    parEndStandard1 = arrayOfInputs(20)
    parPressureRating1 = arrayOfInputs(21)
    parFlowDirection1 = arrayOfInputs(22)
    parPortIndex2 = arrayOfInputs(23)
    parNpd2 = arrayOfInputs(24)
    parEndPreparation2 = arrayOfInputs(25)
    parScheduleThickness2 = arrayOfInputs(26)
    parEndStandard2 = arrayOfInputs(27)
    parPressureRating2 = arrayOfInputs(28)
    parFlowDirection2 = arrayOfInputs(29)
    parPortIndex3 = arrayOfInputs(30)
    parNpd3 = arrayOfInputs(31)
    parEndPreparation3 = arrayOfInputs(32)
    parScheduleThickness3 = arrayOfInputs(33)
    parEndStandard3 = arrayOfInputs(34)
    parPressureRating3 = arrayOfInputs(35)
    parFlowDirection3 = arrayOfInputs(36)
    'Text Inputs
    parId1 = arrayOfInputs(37)
    parId2 = arrayOfInputs(38)
    parId3 = arrayOfInputs(39)
    parNPDUnitType = arrayOfInputs(40)
    parNPDUnitType1 = arrayOfInputs(41)
    parNPDUnitType2 = arrayOfInputs(42)
    parNPDUnitType3 = arrayOfInputs(43)
    
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                            
' Insert your code for output 3(Pipe portion in Noz1, Insulation)
    stPoint.Set 0, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    parInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjPipeinNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz1Ins
    Set ObjPipeinNoz1Ins = Nothing
    
 ' Insert your code for output 4(Noz1 Insulation)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        enPoint.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    If pipeDiam(parPortIndex1) > flangeDiam(parPortIndex1) Then parInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
    Set ObjNoz1Ins = Nothing
 
 ' Insert your code for output 5(Pipe portion in Noz2, Insulation)
    stPoint.Set parFace2toCenter - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjPipeinNoz2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz2Ins
    Set ObjPipeinNoz2Ins = Nothing
    
 ' Insert your code for output 6(Noz2 Insulation)
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        enPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness - 0.0001, 0, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    If pipeDiam(parPortIndex2) > flangeDiam(parPortIndex2) Then parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
    Set ObjNoz2Ins = Nothing
        
' Insert your code for output 7(Pipe portion in Noz3, Insulation)
    stPoint.Set 0, parStrToptoCenter + parFace3toStrTop - flangeThick(parPortIndex3), 0
    enPoint.Set 0, parStrToptoCenter, 0
    
    parInsulationDiameter = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    Set ObjPipeinNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz3Ins
    Set ObjPipeinNoz3Ins = Nothing
    
 ' Insert your code for output 8(Noz3 Insulation)
    stPoint.Set 0, parStrToptoCenter + parFace3toStrTop - flangeThick(parPortIndex3) - parInsulationThickness, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex3) = 0 And parInsulationThickness = 0 Then
        stPoint.Set 0, parStrToptoCenter + parFace3toStrTop - flangeThick(parPortIndex3) - parInsulationThickness - 0.0001, 0
    End If
    enPoint.Set 0, parStrToptoCenter + parFace3toStrTop, 0
 
    parInsulationDiameter = flangeDiam(parPortIndex3) + parInsulationThickness * 2
    If pipeDiam(parPortIndex3) > flangeDiam(parPortIndex3) Then parInsulationDiameter = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    Set ObjNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz3Ins
    Set ObjNoz3Ins = Nothing
    
' Insert your code for output 9(Insulation for Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parStrToptoCenter - parStrWidth, 0
    parInsulationDiameter = parBranchDiameter + 2 * parInsulationThickness
    
    Set ObjInsBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing
    
 ' Insert your code for output 10(Insulation for Strainer)
     stPoint.Set 0, parStrToptoCenter - parStrWidth - parInsulationThickness, 0
     enPoint.Set 0, parStrToptoCenter + parInsulationThickness, 0
    parInsulationDiameter = parStrDiameter + 2 * parInsulationThickness
    Set ObjInsStrainer = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainer
    Set ObjInsStrainer = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
