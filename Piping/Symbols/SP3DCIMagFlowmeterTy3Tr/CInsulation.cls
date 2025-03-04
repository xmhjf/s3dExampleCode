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
'   Creation Date:  Tuesday, Jan 7 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMFM3,
'   PDMS Category Name=IFMY). The symbol consists of seven outputs, five physical and
'   two insulation outputs. Symbol consists of a base (cylinder) and a Transmitter(box).
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
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

Private Sub Class_Initialize()
'''
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function

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
    Dim ObjInsFlowmtrBase As Object
    Dim ObjInsTransmitter As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentOffset As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentOffset1 As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
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

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentOffset = arrayOfInputs(5)  'D5
    parInstrumentLength = arrayOfInputs(6)  'D6
    parInstrumentOffset1 = arrayOfInputs(7) 'D7
    parInstrumentHeight = arrayOfInputs(8)  'D8
    parInstrumentWidth = arrayOfInputs(9)   'D9
    parInstrumentWidth1 = arrayOfInputs(10) 'D10
    parInsulationThickness = arrayOfInputs(11)
    parNPD = arrayOfInputs(12)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(13)
    parScheduleThickness = arrayOfInputs(14)
    parEndStandard = arrayOfInputs(15)
    parPressureRating = arrayOfInputs(16)
    parFlowDirection = arrayOfInputs(17)
    parPortIndex1 = arrayOfInputs(18)
    parNpd1 = arrayOfInputs(19)
    parEndPreparation1 = arrayOfInputs(20)
    parScheduleThickness1 = arrayOfInputs(21)
    parEndStandard1 = arrayOfInputs(22)
    parPressureRating1 = arrayOfInputs(23)
    parFlowDirection1 = arrayOfInputs(24)
    parPortIndex2 = arrayOfInputs(25)
    parNpd2 = arrayOfInputs(26)
    parEndPreparation2 = arrayOfInputs(27)
    parScheduleThickness2 = arrayOfInputs(28)
    parEndStandard2 = arrayOfInputs(29)
    parPressureRating2 = arrayOfInputs(30)
    parFlowDirection2 = arrayOfInputs(31)
    parId1 = arrayOfInputs(32)
    parId2 = arrayOfInputs(33)
    parNPDUnitType = arrayOfInputs(34)
    parNPDUnitType1 = arrayOfInputs(35)
    parNPDUnitType2 = arrayOfInputs(36)
    
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' Insert your code for output 6(Insulation for FlowmtrBase)
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
    
    stPoint.Set -parFace1toCenter - parInsulationThickness, 0, 0
    enPoint.Set parFace2toCenter + parInsulationThickness, 0, 0
'   Use maximum value among pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2 and parFlowDiameter
'   for computing insulation diameter
    parInsulationDiameter = ReturnMax5(pipeDiam(parPortIndex1), pipeDiam(parPortIndex2), _
                                        flangeDiam(parPortIndex1), flangeDiam(parPortIndex2), _
                                        parFlowDiameter) + 2 * parInsulationThickness
    Set ObjInsFlowmtrBase = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFlowmtrBase
    Set ObjInsFlowmtrBase = Nothing
    
' Insert your code for output 7(Insulation  for Transmitter)
'   Note: Insulation box is considered to start from the Flow Meter Base periphery to the
'   Box outer face. So separate insulation is not provided for the connector.
    stPoint.Set parInstrumentOffset - parInstrumentLength / 2 - parInsulationThickness, _
                parFlowDiameter / 2, _
                -parInstrumentWidth - parInsulationThickness
    enPoint.Set parInstrumentOffset + parInstrumentLength / 2 + parInsulationThickness, _
                parInstrumentOffset1 + parInstrumentHeight + parInsulationThickness, _
                parInstrumentWidth1 + parInsulationThickness
    Set ObjInsTransmitter = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTransmitter
    Set ObjInsTransmitter = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
