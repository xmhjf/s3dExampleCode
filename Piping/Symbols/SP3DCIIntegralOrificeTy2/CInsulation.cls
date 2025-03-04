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
'   Author:          SSP
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIOL),
'   Symbol consists of 12 outputs in which 1 insulation, 11 physical aspect outputs.
'   Insulation aspect consist of a single cylinder.
'   The nozzles are fully Parametric.
'
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
    Dim ObjBodyInsCyl1 As Object
    Dim ObjBodyInsCyl2 As Object
    Dim ObjBodyInsBox As Object
    
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentWidth2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentHeight4 As Double
    Dim parInstrumentHeight5 As Double
    Dim parInstrumentWidth3 As Double
    Dim parInstrumentHeight6 As Double
    Dim parInstrumentHeight7 As Double
    Dim parInstrumentHeight8 As Double
    Dim parInstrumentWidth4 As Double
    Dim parInstrumentOffset As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentLength4 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim Diameter As Double
    
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
    parInstrumentWidth = arrayOfInputs(2) ''D2
    parInstrumentWidth1 = arrayOfInputs(3) ''D3
    parInstrumentHeight = arrayOfInputs(4) ''D4
    parInstrumentHeight1 = arrayOfInputs(5) ''D5
    parInstrumentHeight2 = arrayOfInputs(6) ''D6
    parInstrumentWidth2 = arrayOfInputs(7) ''D7
    parInstrumentHeight3 = arrayOfInputs(8) ''D8
    parInstrumentHeight4 = arrayOfInputs(9) ''D9
    parInstrumentHeight5 = arrayOfInputs(10) ''D10
    parInstrumentWidth3 = arrayOfInputs(11) ''D11
    parInstrumentHeight6 = arrayOfInputs(12) ''D12
    parInstrumentHeight7 = arrayOfInputs(13) ''D13
    parInstrumentHeight8 = arrayOfInputs(14) ''D14
    parInstrumentWidth4 = arrayOfInputs(15) ''D15
    parInstrumentOffset = arrayOfInputs(16) ''D16
    parInstrumentDiameter = arrayOfInputs(17) ''D17
    parInstrumentLength = arrayOfInputs(18) ''D18
    parInstrumentLength1 = arrayOfInputs(19) ''D19
    parInstrumentLength2 = arrayOfInputs(20) ''D20
    parInstrumentLength3 = arrayOfInputs(21) 'D21
    parInstrumentLength4 = arrayOfInputs(22) ''D22
    parInsulationThickness = arrayOfInputs(23)
    parNPD = arrayOfInputs(24)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(25)
    parScheduleThickness = arrayOfInputs(26)
    parEndStandard = arrayOfInputs(27)
    parPressureRating = arrayOfInputs(28)
    parFlowDirection = arrayOfInputs(29)
    parPortIndex1 = arrayOfInputs(30)
    parNpd1 = arrayOfInputs(31)
    parEndPreparation1 = arrayOfInputs(32)
    parScheduleThickness1 = arrayOfInputs(33)
    parEndStandard1 = arrayOfInputs(34)
    parPressureRating1 = arrayOfInputs(35)
    parFlowDirection1 = arrayOfInputs(36)
    parPortIndex2 = arrayOfInputs(37)
    parNpd2 = arrayOfInputs(38)
    parEndPreparation2 = arrayOfInputs(39)
    parScheduleThickness2 = arrayOfInputs(40)
    parEndStandard2 = arrayOfInputs(41)
    parPressureRating2 = arrayOfInputs(42)
    parFlowDirection2 = arrayOfInputs(43)
    'Text Inputs
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0
' Insert your code for output 1(Insulation for CP1)
   Dim InsulationDia As Double
   
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
                                            
    InsulationDia = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    If pipeDiam(parPortIndex1) < flangeDiam(parPortIndex1) Then InsulationDia = flangeDiam(parPortIndex1) + 2 * parInsulationThickness

    stPoint.Set -(parInstrumentLength + parInstrumentLength1 / 2), 0, 0
    enPoint.Set -parInstrumentLength1 / 2, 0, 0
    
    Set ObjBodyInsCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsCyl1
    Set ObjBodyInsCyl1 = Nothing
    
' Insert your code for output 2(Insulation for CP2)
    
    InsulationDia = pipeDiam(parPortIndex2) + 2 * parInsulationThickness
    If pipeDiam(parPortIndex2) < flangeDiam(parPortIndex2) Then InsulationDia = flangeDiam(parPortIndex2) + 2 * parInsulationThickness
    
    stPoint.Set (parInstrumentLength2 + parInstrumentLength1 / 2), 0, 0
    enPoint.Set parInstrumentLength1 / 2, 0, 0
    
    Set ObjBodyInsCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsCyl2
    Set ObjBodyInsCyl2 = Nothing

' Insert your code for output 3(Insulation for Box between CP1 & CP2)
    
   
    stPoint.Set -parInstrumentLength1 / 2 - parInsulationThickness, _
                -parInstrumentHeight1 - parInsulationThickness, _
                parInstrumentWidth1 + parInsulationThickness
                                        
    enPoint.Set parInstrumentLength1 / 2 + parInsulationThickness, _
                parInstrumentHeight + parInsulationThickness, _
                -parInstrumentWidth - parInsulationThickness
    
    Set ObjBodyInsBox = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsBox
    Set ObjBodyInsBox = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

