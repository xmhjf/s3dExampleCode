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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Friday, Jan 17 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   19.Nov.2004     MS                      TR 67097  Provided checks to take care of cases when flange diameter retrieved from port is zero.
'   08.SEP.2006     KKC                     DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG                     TR-124959 Insulation for the ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double

    Dim iOutput     As Integer
    Dim ObjNozzle1Insul As Object
    Dim ObjNoz1ExtnInsul As Object
    Dim ObjJunction1Insul As Object
    Dim ObjSkewPipeInsul As Object
    Dim ObjJunction2Insul As Object
    Dim ObjNoz2ExtnInsul As Object
    Dim ObjNozzle2Insul As Object
    Dim parInsulationThickness As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
  
 ' Dynamic Port Input parameters variable declaration
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
    parOffset = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
    parId1 = arrayOfInputs(26)
    parId2 = arrayOfInputs(27)
    parNPDUnitType = arrayOfInputs(28)
    parNPDUnitType1 = arrayOfInputs(29)
    parNPDUnitType2 = arrayOfInputs(30)
    
    iOutput = 0
 ' Insert your code for output 6(Insulation for Nozzle1)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
 
 
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia As Double
    
    If CmpDblGreaterthan(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1)) Then
       InsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    Else
        InsulationDia = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    End If
        
    If CmpDblEqual(flangeThick(parPortIndex1), 0) Then
         flangeThick(parPortIndex1) = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1) - parInsulationThickness), 0, 0
    
    Set ObjNozzle1Insul = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1Insul
    Set ObjNozzle1Insul = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
 ' Insert your code for output 7(Insulation Nozzle1 Extension)
    Dim NozInsStPoint   As New AutoMath.DPosition
    Dim NozInsEnPoint   As New AutoMath.DPosition
    Dim NozInsulationDia As Double

    NozInsStPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    NozInsEnPoint.Set -(parFace1toCenter / 3 + parFace1toCenter / 3), 0, 0
    NozInsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    
    Set ObjNoz1ExtnInsul = PlaceCylinder(m_OutputColl, NozInsStPoint, NozInsEnPoint, NozInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1ExtnInsul
    Set ObjNoz1ExtnInsul = Nothing
    Set NozInsStPoint = Nothing
    Set NozInsEnPoint = Nothing
    
 ' Insert your code for output 8(Junction 1 Insulation)
    Dim JunInsStPoint   As New AutoMath.DPosition
    Dim JunInsEnPoint   As New AutoMath.DPosition
    Dim JunInsulationDia As Double
    JunInsStPoint.Set -(parFace1toCenter / 3 + parFace1toCenter / 3 + parInsulationThickness), 0, 0
    JunInsEnPoint.Set -(parFace1toCenter / 3 - parInsulationThickness), 0, 0
    JunInsulationDia = pipeDiam(parPortIndex1) * 2 + parInsulationThickness * 2
    
    Set ObjJunction1Insul = PlaceCylinder(m_OutputColl, JunInsStPoint, JunInsEnPoint, JunInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction1Insul
    Set ObjJunction1Insul = Nothing
    Set JunInsStPoint = Nothing
    Set JunInsEnPoint = Nothing
    
 ' Insert your code for output 9(Insulation for Skew Pipe)
    Dim Angle As Double
    Dim Hypotenous  As Double
    Dim SkewInsStPoint   As New AutoMath.DPosition
    Dim SkewInsEnPoint   As New AutoMath.DPosition
    Dim SkewInsulationDia As Double
    Angle = Atn(parOffset / (parFace1toCenter / 3))
    Hypotenous = 0.5 * pipeDiam(parPortIndex1) * Tan(Angle)
    SkewInsStPoint.x = -(Hypotenous * Cos(Angle) + parFace1toCenter / 3)
    SkewInsStPoint.y = -Hypotenous * Sin(Angle)
    SkewInsStPoint.z = 0
    SkewInsEnPoint.x = (Hypotenous * Cos(Angle) + parFace2toCenter / 3)
    SkewInsEnPoint.y = (parOffset + (parFace2toCenter / parFace1toCenter) * parOffset) + _
                        Hypotenous * Sin(Angle)
    SkewInsEnPoint.z = 0
    SkewInsulationDia = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    
    Set ObjSkewPipeInsul = PlaceCylinder(m_OutputColl, SkewInsStPoint, SkewInsEnPoint, SkewInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSkewPipeInsul
    Set ObjSkewPipeInsul = Nothing
    Set SkewInsStPoint = Nothing
    Set SkewInsEnPoint = Nothing
    
 ' Insert your code for output 10(Junction 2 Insulation)
    Dim JunAInsStPoint   As New AutoMath.DPosition
    Dim JunAInsEnPoint   As New AutoMath.DPosition
    Dim JunAInsulationDia As Double
    JunAInsStPoint.Set (parFace2toCenter / 3 - parInsulationThickness), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    JunAInsEnPoint.Set (parFace2toCenter / 3 + parFace2toCenter / 3 + parInsulationThickness), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    JunAInsulationDia = pipeDiam(parPortIndex2) * 2 + parInsulationThickness * 2
    
    Set ObjJunction2Insul = PlaceCylinder(m_OutputColl, JunAInsStPoint, JunAInsEnPoint, JunAInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction2Insul
    Set ObjJunction2Insul = Nothing
    Set JunAInsStPoint = Nothing
    Set JunAInsEnPoint = Nothing
    
 ' Insert your code for output 11(Insulation for Nozzle 2 Extn)
    Dim NozAInsStPoint   As New AutoMath.DPosition
    Dim NozAInsEnPoint   As New AutoMath.DPosition
    Dim NozAInsulationDia As Double
    NozAInsStPoint.Set (parFace2toCenter / 3 + parFace2toCenter / 3), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    NozAInsEnPoint.Set (parFace2toCenter - flangeThick(parPortIndex2)), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    NozAInsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    
    Set ObjNoz2ExtnInsul = PlaceCylinder(m_OutputColl, NozAInsStPoint, NozAInsEnPoint, NozAInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2ExtnInsul
    Set ObjNoz2ExtnInsul = Nothing
    Set NozAInsStPoint = Nothing
    Set NozAInsEnPoint = Nothing
     
 ' Insert your code for output 12(Insulation for Nozzle 2)
    Dim Noz2InsStPoint   As New AutoMath.DPosition
    Dim Noz2InsEnPoint   As New AutoMath.DPosition
    Dim Noz2InsulationDia As Double
    If CmpDblGreaterthan(pipeDiam(parPortIndex2), flangeDiam(parPortIndex2)) Then
       InsulationDia = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    Else
       InsulationDia = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    End If
        
    If CmpDblEqual(flangeThick(parPortIndex2), 0) Then
         flangeThick(parPortIndex2) = NEGLIGIBLE_THICKNESS
    End If
    
    Noz2InsStPoint.Set parFace2toCenter, parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    Noz2InsEnPoint.Set (parFace2toCenter - flangeThick(parPortIndex2) - parInsulationThickness), _
                        parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0

    Set ObjNozzle2Insul = PlaceCylinder(m_OutputColl, Noz2InsStPoint, Noz2InsEnPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2Insul
    Set ObjNozzle2Insul = Nothing
    Set Noz2InsStPoint = Nothing
    Set Noz2InsEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
