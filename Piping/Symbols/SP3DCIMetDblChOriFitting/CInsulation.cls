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
'   Creation Date:  Friday, Jan 10 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMDC2,
'   PDMS Category Name=IFRC). The symbol consists of 12  outputs, nine physical,
'   two insulation outputs and one Operation output.
'   The 4 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

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
    Dim ObjInsBody As Object
    Dim ObjInsPipeLine As Object
    
    Dim parInstrumentLength      As Double
    Dim parInstrumentLength1     As Double
    Dim parInstrumentHeight      As Double
    Dim parInstrumentHeight1     As Double
    Dim parInstrumentLength2     As Double
    Dim parInstrumentLength3     As Double
    Dim parInstrumentWidth       As Double
    Dim parInstrumentWidth1      As Double
    Dim parArmLength             As Double
    Dim parNozzleOffset          As Double
    Dim parNozzleOffset1         As Double
    Dim parArmHeight             As Double
    Dim parArmHeight1            As Double
    Dim parInstrumentLength4     As Double
    Dim parInstrumentLength5     As Double
    Dim parInsulationThickness   As Double
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
    Dim parInsulationDiameter As Double
    
    Dim parOperationHeight As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)      'D2
    parInstrumentLength1 = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength2 = arrayOfInputs(6)     'D6
    parInstrumentLength3 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parArmLength = arrayOfInputs(10)            'D10
    parNozzleOffset = arrayOfInputs(11)         'D11
    parNozzleOffset1 = arrayOfInputs(12)        'D12
    parArmHeight = arrayOfInputs(13)            'D14
    parArmHeight1 = arrayOfInputs(14)           'D15
    parInstrumentLength4 = arrayOfInputs(15)    'D16
    parInstrumentLength5 = arrayOfInputs(16)    'D17
    parInsulationThickness = arrayOfInputs(17)
    parNPD = arrayOfInputs(18)              'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(19)
    parScheduleThickness = arrayOfInputs(20)
    parEndStandard = arrayOfInputs(21)
    parPressureRating = arrayOfInputs(22)
    parFlowDirection = arrayOfInputs(23)
    parPortIndex1 = arrayOfInputs(24)
    parNpd1 = arrayOfInputs(25)
    parEndPreparation1 = arrayOfInputs(26)
    parScheduleThickness1 = arrayOfInputs(27)
    parEndStandard1 = arrayOfInputs(28)
    parPressureRating1 = arrayOfInputs(29)
    parFlowDirection1 = arrayOfInputs(30)
    parPortIndex2 = arrayOfInputs(31)
    parNpd2 = arrayOfInputs(32)
    parEndPreparation2 = arrayOfInputs(33)
    parScheduleThickness2 = arrayOfInputs(34)
    parEndStandard2 = arrayOfInputs(35)
    parPressureRating2 = arrayOfInputs(36)
    parFlowDirection2 = arrayOfInputs(37)
    parPortIndex3 = arrayOfInputs(38)
    parNpd3 = arrayOfInputs(39)
    parEndPreparation3 = arrayOfInputs(40)
    parScheduleThickness3 = arrayOfInputs(41)
    parEndStandard3 = arrayOfInputs(42)
    parPressureRating3 = arrayOfInputs(43)
    parFlowDirection3 = arrayOfInputs(44)
    parPortIndex4 = arrayOfInputs(45)
    parNpd4 = arrayOfInputs(46)
    parEndPreparation4 = arrayOfInputs(47)
    parScheduleThickness4 = arrayOfInputs(48)
    parEndStandard4 = arrayOfInputs(49)
    parPressureRating4 = arrayOfInputs(50)
    parFlowDirection4 = arrayOfInputs(51)
    parOperationHeight = arrayOfInputs(52)     'D13
'   Comment to make test with text inputs
    parId1 = arrayOfInputs(53)
    parId2 = arrayOfInputs(54)
    parId3 = arrayOfInputs(55)
    parId4 = arrayOfInputs(56)
    parNPDUnitType = arrayOfInputs(57)
    parNPDUnitType1 = arrayOfInputs(58)
    parNPDUnitType2 = arrayOfInputs(59)
    parNPDUnitType3 = arrayOfInputs(60)
    parNPDUnitType4 = arrayOfInputs(61)
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

' Insert your code for output 10(InsBody)
    stPoint.Set -parInstrumentLength2 - parInsulationThickness, _
                -parInstrumentHeight - parInsulationThickness, _
                parInstrumentWidth + parInsulationThickness

    enPoint.Set parInstrumentLength3 + parInsulationThickness, _
                parInstrumentHeight1 + parInsulationThickness, _
                -parInstrumentWidth1 - parInsulationThickness
    Set ObjInsBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    Set ObjInsBody = Nothing
    
' Insert your code for output 11(InsPipeLine)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    stPoint.Set -parInstrumentLength - parInstrumentLength4, 0, 0
    enPoint.Set parInstrumentLength1 + parInstrumentLength5, 0, 0
'   Assumption: Instrument Left Hand Side/Right Hand Side flanges (At distances D2 and D3 respectively)
'   diameter is assumed to be 1.2 times corresponding pipe diameter.
    Dim LHSFlangeDiameter As Double
    Dim RHSFlangeDiameter As Double
    
    LHSFlangeDiameter = 1.2 * pipeDiam(parPortIndex1)
    RHSFlangeDiameter = 1.2 * pipeDiam(parPortIndex2)
    parInsulationDiameter = ReturnMax4(flangeDiam(parPortIndex1), LHSFlangeDiameter, _
                            RHSFlangeDiameter, flangeDiam(parPortIndex2)) + _
                            2 * parInsulationThickness
    Set ObjInsPipeLine = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPipeLine
    Set ObjInsPipeLine = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
