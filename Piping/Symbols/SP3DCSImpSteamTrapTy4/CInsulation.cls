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
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol consist of Physical and Insulation Aspects. Physical aspect is
'   made up of Two Nozzles with Length, Horizontal Cylinder and Vertical Cylinder
'   Insulation aspect consist of Insulation for both Nozzles, Insulation for the pipe portion in
'   both Nozzles, Insultion for Hori and Verital Cylinder.
'   For more details about this Symbol please refer PDS Piping Component reference Guide
'   Symbol Number SN=S32BC or S32BZ
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   11.Apr.2005     svsmylav        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                           Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim ObjHoriCylinIns As Object
    Dim ObjVertCylinIns As Object
    Dim ObjPipeinNoz1Ins As Object
    Dim ObjPipeinNoz2Ins As Object
    Dim ObjNoz1Ins As Object
    Dim ObjNoz2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylinderDiameter As Double
    Dim parVertCylinHeightFromCen As Double
    Dim parHoriCylinLengthFromCen As Double
    Dim parInsulationThickness  As Double
    
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
    parFace1toCenter = arrayOfInputs(2)     'P2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parCylinderDiameter = arrayOfInputs(4)     'P4
    parVertCylinHeightFromCen = arrayOfInputs(5)      'P5
    parHoriCylinLengthFromCen = arrayOfInputs(6)      'P6
    parInsulationThickness = arrayOfInputs(7)
    parNPD = arrayOfInputs(8)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(9)
    parScheduleThickness = arrayOfInputs(10)
    parEndStandard = arrayOfInputs(11)
    parPressureRating = arrayOfInputs(12)
    parFlowDirection = arrayOfInputs(13)
    parPortIndex1 = arrayOfInputs(14)
    parNpd1 = arrayOfInputs(15)
    parEndPreparation1 = arrayOfInputs(16)
    parScheduleThickness1 = arrayOfInputs(17)
    parEndStandard1 = arrayOfInputs(18)
    parPressureRating1 = arrayOfInputs(19)
    parFlowDirection1 = arrayOfInputs(20)
    parPortIndex2 = arrayOfInputs(21)
    parNpd2 = arrayOfInputs(22)
    parEndPreparation2 = arrayOfInputs(23)
    parScheduleThickness2 = arrayOfInputs(24)
    parEndStandard2 = arrayOfInputs(25)
    parPressureRating2 = arrayOfInputs(26)
    parFlowDirection2 = arrayOfInputs(27)
    'Text Inputs
    parId1 = arrayOfInputs(28)
    parId2 = arrayOfInputs(29)
    parNPDUnitType = arrayOfInputs(30)
    parNPDUnitType1 = arrayOfInputs(31)
    parNPDUnitType2 = arrayOfInputs(32)
    
    iOutput = 0

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
                                            
 ' Insert your code for output 5(Horizontal Cylin Insulation)
    Dim parInsulationDiameter As Double
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, 0
    endPoint.Set parHoriCylinLengthFromCen + parInsulationThickness, 0, 0
    parInsulationDiameter = parCylinderDiameter + parInsulationThickness * 2
    Set ObjHoriCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinIns
    Set ObjHoriCylinIns = Nothing
    
 ' Insert your code for output 6(Vertical Cylin Insulation)
    startPoint.Set 0, -parVertCylinHeightFromCen - parInsulationThickness, 0
    endPoint.Set 0, parCylinderDiameter / 2 + parInsulationThickness, 0
    parInsulationDiameter = parCylinderDiameter + parInsulationThickness * 2
    Set ObjVertCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylinIns
    Set ObjVertCylinIns = Nothing
    
 ' Insert your code for output 7(Pipe portion in Noz1, Insulation)
    startPoint.Set 0, -parVertCylinHeightFromCen, 0
    endPoint.Set 0, -parFace1toCenter + flangeThick(parPortIndex1), 0
    parInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjPipeinNoz1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz1Ins
    Set ObjPipeinNoz1Ins = Nothing
    
 ' Insert your code for output 8(Noz1 Insulation)
    startPoint.Set 0, -parFace1toCenter, 0
    endPoint.Set 0, -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        endPoint.Set 0, -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    If pipeDiam(parPortIndex1) > flangeDiam(parPortIndex1) Then parInsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
    Set ObjNoz1Ins = Nothing
 
 ' Insert your code for output 9(Pipe portion in Noz2, Insulation)
    startPoint.Set parFace2toCenter - flangeThick(parPortIndex2), 0, 0
    endPoint.Set parHoriCylinLengthFromCen, 0, 0
    parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjPipeinNoz2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz2Ins
    Set ObjPipeinNoz2Ins = Nothing
    
 ' Insert your code for output 10(Noz2 Insulation)
    startPoint.Set parFace2toCenter, 0, 0
    endPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        endPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness + 0.0001, 0, 0
    End If
    parInsulationDiameter = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    If pipeDiam(parPortIndex2) > flangeDiam(parPortIndex2) Then parInsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
    Set ObjNoz2Ins = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
