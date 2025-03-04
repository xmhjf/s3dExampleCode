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
'   Creation Date:  Friday, Jan 3 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PMS Category Name=IKG1)
'   The symbol consists of 8 outputs; six physical, one insulation and one Operation output.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
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
    Dim ObjValveBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parValveHeight As Double
    Dim parValveHeight1 As Double
    Dim parValveWidth As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorDiameter As Double
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
    
    Dim parOperationHeight As Double
    Dim parOperationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)   'D3
    parValveHeight = arrayOfInputs(4)     'D4
    parValveHeight1 = arrayOfInputs(5)    'D5
    parValveWidth = arrayOfInputs(6)      'D6
    parActuatorHeight = arrayOfInputs(7)      'D7
    parActuatorHeight1 = arrayOfInputs(8)     'D8
    parActuatorDiameter = arrayOfInputs(9)    'D9
    parInsulationThickness = arrayOfInputs(10)
    parNPD = arrayOfInputs(11)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(12)
    parScheduleThickness = arrayOfInputs(13)
    parEndStandard = arrayOfInputs(14)
    parPressureRating = arrayOfInputs(15)
    parFlowDirection = arrayOfInputs(16)
    parPortIndex1 = arrayOfInputs(17)
    parNpd1 = arrayOfInputs(18)
    parEndPreparation1 = arrayOfInputs(19)
    parScheduleThickness1 = arrayOfInputs(20)
    parEndStandard1 = arrayOfInputs(21)
    parPressureRating1 = arrayOfInputs(22)
    parFlowDirection1 = arrayOfInputs(23)
    parPortIndex2 = arrayOfInputs(24)
    parNpd2 = arrayOfInputs(25)
    parEndPreparation2 = arrayOfInputs(26)
    parScheduleThickness2 = arrayOfInputs(27)
    parEndStandard2 = arrayOfInputs(28)
    parPressureRating2 = arrayOfInputs(29)
    parFlowDirection2 = arrayOfInputs(30)
    parOperationHeight = arrayOfInputs(31)     'D10
    parOperationDiameter = arrayOfInputs(32) 'D11
    'Text Inputs
    parId1 = arrayOfInputs(33)
    parId2 = arrayOfInputs(34)
    parNPDUnitType = arrayOfInputs(35)
    parNPDUnitType1 = arrayOfInputs(36)
    parNPDUnitType2 = arrayOfInputs(37)
    
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
                                            
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam(parPortIndex1), pipeDiam(parPortIndex2), flangeDiam(parPortIndex1), flangeDiam(parPortIndex2))
    
    stPoint.Set -parFace1toCenter - parInsulationThickness, _
                        -maxDia / 2 - parInsulationThickness, _
                        parValveWidth / 2 + parInsulationThickness
    enPoint.Set parFace2toCenter + parInsulationThickness, _
                        parValveHeight1 + parInsulationThickness, _
                        -parValveWidth / 2 - parInsulationThickness
    Set ObjValveBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
