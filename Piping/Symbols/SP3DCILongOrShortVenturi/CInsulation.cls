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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Dec 30 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ILSV, PDMS Category Name=IFEE)
'   The symbol consists of eleven outputs, eight physical and three insulation outputs.
'   The 4 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    Dim iOutput     As Double
    Dim ObjVenturiBodyIns As Object
    Dim ObjNozz3BodyIns As Object
    Dim ObjNozz4BodyIns As Object
    
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentDiameter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzleOffset2 As Double
    Dim parNozzleOffset3 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double
    
'   Dynamic Port Input parameters variable declaration
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
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)  'D2
    parInstrumentLength1 = arrayOfInputs(3) 'D3
    parInstrumentLength2 = arrayOfInputs(4) 'D4
    parInstrumentLength3 = arrayOfInputs(5) 'D5
    parInstrumentDiameter = arrayOfInputs(6)       'D6
    parNozzleOffset = arrayOfInputs(7)     'D7
    parNozzleOffset1 = arrayOfInputs(8)     'D8
    parNozzleOffset2 = arrayOfInputs(9)        'D9
    parNozzleOffset3 = arrayOfInputs(10)  'D10
    parPortRotation1 = arrayOfInputs(11)                'D11
    parPortRotation2 = arrayOfInputs(12)         'D12
    parInsulationThickness = arrayOfInputs(13)
'   Dynamic Port Input parameters variable declaration
    parNPD = arrayOfInputs(14)
    parEndPreparation = arrayOfInputs(15)
    parScheduleThickness = arrayOfInputs(16)
    parEndStandard = arrayOfInputs(17)
    parPressureRating = arrayOfInputs(18)
    parFlowDirection = arrayOfInputs(19)
    parPortIndex1 = arrayOfInputs(20)
    parNpd1 = arrayOfInputs(21)
    parEndPreparation1 = arrayOfInputs(22)
    parScheduleThickness1 = arrayOfInputs(23)
    parEndStandard1 = arrayOfInputs(24)
    parPressureRating1 = arrayOfInputs(25)
    parFlowDirection1 = arrayOfInputs(26)
    parPortIndex2 = arrayOfInputs(27)
    parNpd2 = arrayOfInputs(28)
    parEndPreparation2 = arrayOfInputs(29)
    parScheduleThickness2 = arrayOfInputs(30)
    parEndStandard2 = arrayOfInputs(31)
    parPressureRating2 = arrayOfInputs(32)
    parFlowDirection2 = arrayOfInputs(33)
    parPortIndex3 = arrayOfInputs(34)
    parNpd3 = arrayOfInputs(35)
    parEndPreparation3 = arrayOfInputs(36)
    parScheduleThickness3 = arrayOfInputs(37)
    parEndStandard3 = arrayOfInputs(38)
    parPressureRating3 = arrayOfInputs(39)
    parFlowDirection3 = arrayOfInputs(40)
    parPortIndex4 = arrayOfInputs(41)
    parNpd4 = arrayOfInputs(42)
    parEndPreparation4 = arrayOfInputs(43)
    parScheduleThickness4 = arrayOfInputs(44)
    parEndStandard4 = arrayOfInputs(45)
    parPressureRating4 = arrayOfInputs(46)
    parFlowDirection4 = arrayOfInputs(47)
    parId1 = arrayOfInputs(48)
    parId2 = arrayOfInputs(49)
    parId3 = arrayOfInputs(50)
    parId4 = arrayOfInputs(51)
    parNPDUnitType = arrayOfInputs(52)
    parNPDUnitType1 = arrayOfInputs(53)
    parNPDUnitType2 = arrayOfInputs(54)
    parNPDUnitType3 = arrayOfInputs(55)
    parNPDUnitType4 = arrayOfInputs(56)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' The orgin is considered to be along the flow axis and mid of parInstrumentLength2
' Insert your code for output 9(Venturi Body Insulation)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    Dim parInsulationDiameter As Double
    'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam(parPortIndex1), pipeDiam(parPortIndex2), _
                        flangeDiam(parPortIndex1), flangeDiam(parPortIndex2))
    parInsulationDiameter = maxDia + 2 * parInsulationThickness
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength, 0, 0
    enPoint.Set parInstrumentLength2 / 2 + parInstrumentLength3, 0, 0
    Set ObjVenturiBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVenturiBodyIns
    Set ObjVenturiBodyIns = Nothing
    
' Insert your code for output 10(Nozzle3 Body Insulation)
    stPoint.Set -(parNozzleOffset1 - parNozzleOffset), _
                        (parNozzleOffset2) * Sin(parPortRotation1), _
                        (parNozzleOffset2) * Cos(parPortRotation1)
    enPoint.Set -(parNozzleOffset1 - parNozzleOffset), 0, 0
    parInsulationDiameter = pipeDiam(parPortIndex3) + parInsulationThickness * 2
    If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then
        parInsulationDiameter = flangeDiam(parPortIndex3) + parInsulationThickness * 2
    End If
    Set ObjNozz3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz3BodyIns
    Set ObjNozz3BodyIns = Nothing
    
' Insert your code for output 11(Nozzle4 Body Insulation)
    stPoint.Set 0, _
                        (parNozzleOffset3) * Sin(parPortRotation2), _
                        (parNozzleOffset3) * Cos(parPortRotation2)
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam(parPortIndex4) + parInsulationThickness * 2
    If flangeDiam(parPortIndex4) > pipeDiam(parPortIndex4) Then
        parInsulationDiameter = flangeDiam(parPortIndex4) + parInsulationThickness * 2
    End If
   Set ObjNozz4BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz4BodyIns
    Set ObjNozz4BodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax4(A As Double, B As Double, C As Double, D As Double) As Double
    Dim MaxValue As Double
    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

