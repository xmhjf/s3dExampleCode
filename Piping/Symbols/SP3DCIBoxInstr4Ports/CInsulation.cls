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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          VM
'   Creation Date:  Oct 1, 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on PDS 4BOX4 model code
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   07.Oct.2005     kkk         CR-60561: Vivek provided the symbol and we did Minor Changes.
'
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
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    Dim iOutput     As Double
    Dim ObjInstrumentBody As Object
    Dim ObjInstrumentBodyIns As Object
        
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parAngle As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
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
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)   'A
    parInstrumentWidth = arrayOfInputs(3)   'B
    parInstrumentHeight = arrayOfInputs(4)   'C
    parInstrumentLength1 = arrayOfInputs(5)
    parInstrumentLength2 = arrayOfInputs(6)
    parAngle = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parNPD = arrayOfInputs(9)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(10)
    parScheduleThickness = arrayOfInputs(11)
    parEndStandard = arrayOfInputs(12)
    parPressureRating = arrayOfInputs(13)
    parFlowDirection = arrayOfInputs(14)
    parPortIndex1 = arrayOfInputs(15)
    parNpd1 = arrayOfInputs(16)
    parEndPreparation1 = arrayOfInputs(17)
    parScheduleThickness1 = arrayOfInputs(18)
    parEndStandard1 = arrayOfInputs(19)
    parPressureRating1 = arrayOfInputs(20)
    parFlowDirection1 = arrayOfInputs(21)
    parPortIndex2 = arrayOfInputs(22)
    parNpd2 = arrayOfInputs(23)
    parEndPreparation2 = arrayOfInputs(24)
    parScheduleThickness2 = arrayOfInputs(25)
    parEndStandard2 = arrayOfInputs(26)
    parPressureRating2 = arrayOfInputs(27)
    parFlowDirection2 = arrayOfInputs(28)
    parPortIndex3 = arrayOfInputs(29)
    parNpd3 = arrayOfInputs(30)
    parEndPreparation3 = arrayOfInputs(31)
    parScheduleThickness3 = arrayOfInputs(32)
    parEndStandard3 = arrayOfInputs(33)
    parPressureRating3 = arrayOfInputs(34)
    parFlowDirection3 = arrayOfInputs(35)
    parPortIndex4 = arrayOfInputs(36)
    parNpd4 = arrayOfInputs(37)
    parEndPreparation4 = arrayOfInputs(38)
    parScheduleThickness4 = arrayOfInputs(39)
    parEndStandard4 = arrayOfInputs(40)
    parPressureRating4 = arrayOfInputs(41)
    parFlowDirection4 = arrayOfInputs(42)
    
    'Text Inputs
    parId1 = arrayOfInputs(43)
    parId2 = arrayOfInputs(44)
    parId3 = arrayOfInputs(45)
    parId4 = arrayOfInputs(46)
    parNPDUnitType = arrayOfInputs(47)
    parNPDUnitType1 = arrayOfInputs(48)
    parNPDUnitType2 = arrayOfInputs(49)
    parNPDUnitType3 = arrayOfInputs(50)
    parNPDUnitType4 = arrayOfInputs(51)
    
    iOutput = 0

 ' Insert your code for output 5(Valve Body Insulation)
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
                                            
    
    stPoint.Set -parInsulationThickness - parInstrumentLength * 0.5, -parInsulationThickness - parInstrumentWidth * 0.5, -parInsulationThickness - parInstrumentHeight * 0.5
    enPoint.Set parInsulationThickness + parInstrumentLength * 0.5, parInsulationThickness + parInstrumentWidth * 0.5, parInsulationThickness + parInstrumentHeight * 0.5
    Set ObjInstrumentBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
    
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstrumentBodyIns
    Set ObjInstrumentBodyIns = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

