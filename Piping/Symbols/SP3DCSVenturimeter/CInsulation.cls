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
'   Creation Date:  Tuesday, Jan 21 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This Symbol details are taken from PDS Piping Component Data Manual. Model no SN=I213AZ
'Page no D-123
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
    Dim ObjDivergConeIns As Object
    Dim ObjConvergConeIns As Object
    Dim ObjVenturiIns As Object
    Dim ObjNozzle1Ins As Object
    Dim ObjNozzle2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parVenturiLength As Double
    Dim parVenturiDiameter As Double
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
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parVenturiLength = arrayOfInputs(4)
    parVenturiDiameter = arrayOfInputs(5)
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

    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                                            
 ' Insert your code for output 6(Converging Cone)
    Dim basecenter As AutoMath.DPosition
    Dim topcenter As AutoMath.DPosition
    Set basecenter = New AutoMath.DPosition
    Set topcenter = New AutoMath.DPosition
    
    basecenter.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    topcenter.Set -parVenturiLength / 2, 0, 0
    Set ObjConvergConeIns = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam(parPortIndex1) / 2 + parInsulationThickness, _
                                                                parVenturiDiameter / 2 + parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergConeIns
    Set ObjConvergConeIns = Nothing
    
 ' Insert your code for output 7(Venturi)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint  As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parVenturiLength / 2, 0, 0
    endPoint.Set parVenturiLength / 2, 0, 0
    Set ObjVenturiIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parVenturiDiameter + parInsulationThickness * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVenturiIns
    Set ObjVenturiIns = Nothing
    
 ' Insert your code for output 8(Nozzle1 Insulation)
    Dim InsulationDia As Double
    
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex1) = 0 And parInsulationThickness = 0 Then
        endPoint.Set -parFace1toCenter + flangeThick(parPortIndex1) + parInsulationThickness + 0.0001, 0, 0
    End If
    
    InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then
        InsulationDia = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
    Set ObjNozzle1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1Ins
    Set ObjNozzle1Ins = Nothing
    
 ' Insert your code for output 9(Diverging Cone)
    basecenter.Set parFace2toCenter - flangeThick(parPortIndex2), 0, 0
    topcenter.Set parVenturiLength / 2, 0, 0
    Set ObjDivergConeIns = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam(parPortIndex2) / 2 + parInsulationThickness, _
                                                                parVenturiDiameter / 2 + parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergConeIns
    Set ObjDivergConeIns = Nothing
    Set basecenter = Nothing
    Set topcenter = Nothing
    
 ' Insert your code for output 10(Nozzle2 Insulation)
    startPoint.Set parFace2toCenter, 0, 0
    endPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness, 0, 0
'   If flange thickness and insulation thickness are zero, provide a small value to ensure non-zero cylinder length value.
    If flangeThick(parPortIndex2) = 0 And parInsulationThickness = 0 Then
        endPoint.Set parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness - 0.0001, 0, 0
    End If

    InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then
        InsulationDia = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
    Set ObjNozzle2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2Ins
    Set ObjNozzle2Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
