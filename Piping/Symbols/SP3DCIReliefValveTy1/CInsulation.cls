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
'   Author:          MS
'   Creation Date:  Tuesday 7 Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT1, PDMS Category Name=IVPG)
'   The symbol consists of eight outputs,  five physical and three insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
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
    Dim sptoffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim ObjValveBodyIns As Object
    Dim ObjNozz1BodyIns As Object
    Dim ObjNozz2BodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
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
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentHeight2 = arrayOfInputs(10)    'D10
    parInstrumentHeight3 = arrayOfInputs(11)    'D11
    parInstrumentDiameter = arrayOfInputs(12)   'D12
    parInstrumentDiameter1 = arrayOfInputs(13)  'D13
    parInsulationThickness = arrayOfInputs(14)
    parNPD = arrayOfInputs(15)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(16)
    parScheduleThickness = arrayOfInputs(17)
    parEndStandard = arrayOfInputs(18)
    parPressureRating = arrayOfInputs(19)
    parFlowDirection = arrayOfInputs(20)
    parPortIndex1 = arrayOfInputs(21)
    parNpd1 = arrayOfInputs(22)
    parEndPreparation1 = arrayOfInputs(23)
    parScheduleThickness1 = arrayOfInputs(24)
    parEndStandard1 = arrayOfInputs(25)
    parPressureRating1 = arrayOfInputs(26)
    parFlowDirection1 = arrayOfInputs(27)
    parPortIndex2 = arrayOfInputs(28)
    parNpd2 = arrayOfInputs(29)
    parEndPreparation2 = arrayOfInputs(30)
    parScheduleThickness2 = arrayOfInputs(31)
    parEndStandard2 = arrayOfInputs(32)
    parPressureRating2 = arrayOfInputs(33)
    parFlowDirection2 = arrayOfInputs(34)
'Comment to make test with text inputs
    parId1 = arrayOfInputs(35)
    parId2 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)
        
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim Diameter As Double
'The origin is considered to be at intersection of nozzle axis
 ' Insert your code for output 4(Valve Body Insulation)
    stPoint.Set -parInstrumentLength1 - parInsulationThickness, _
                        -parInstrumentHeight - parInsulationThickness, _
                        parInstrumentWidth + parInsulationThickness

    enPoint.Set parInstrumentLength + parInsulationThickness, _
                        parInstrumentHeight1 + parInstrumentHeight3 + parInstrumentHeight2 + parInsulationThickness, _
                        -parInstrumentWidth1 - parInsulationThickness

    Set ObjValveBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
 ' Insert your code for output 5(Nozzle1 Body Insulation)
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
                                            flangeDiam, sptoffset, depth
    Dim InsulationDiameter As Double
    InsulationDiameter = pipeDiam(1) + parInsulationThickness * 2
    If flangeDiam(1) > pipeDiam(1) Then
        InsulationDiameter = flangeDiam(1) + parInsulationThickness * 2
    End If
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parInstrumentLength1, 0, 0
    Set ObjNozz1BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz1BodyIns
    Set ObjNozz1BodyIns = Nothing
    
 ' Insert your code for output 6(Nozzle2 Body Insulation)
    InsulationDiameter = pipeDiam(2) + parInsulationThickness * 2
    If flangeDiam(2) > pipeDiam(2) Then InsulationDiameter = flangeDiam(2) + parInsulationThickness * 2
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, -parInstrumentHeight, 0
    Set ObjNozz2BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz2BodyIns
    Set ObjNozz2BodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

