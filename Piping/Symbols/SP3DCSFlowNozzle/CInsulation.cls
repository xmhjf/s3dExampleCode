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
'   Creation Date:  Wed, Jan 29 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details are available in PDS Piping Data Reference Guide at page no D-106 and SN=1219AZ
'
'   Consist of Physical and Insulation Aspects. Physical aspect is made up of Two Nozzles, One FlowNozzleCylBody of Cylinder,
'   and Filter body of truncated cone. Insulation aspect consist of a Cylinder
'   for FlowNozzleBody Insulation.Option Explicit
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
    Dim ObjFlowNozzleBodyIns As Object
    
    Dim parFacetoFace As Double
    Dim parFlowNozzleDiameter1 As Double
    Dim parFlowNozzleLength As Double
    Dim parFlowNozzleEndDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFlowNozzleDiameter2 As Double
    
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
    parFacetoFace = arrayOfInputs(2)                         'P2
    parFlowNozzleDiameter1 = arrayOfInputs(3)              'P3
    parFlowNozzleLength = arrayOfInputs(4)                  'P4
    parFlowNozzleEndDiameter = arrayOfInputs(5)           'P5
    parFlowNozzleDiameter2 = arrayOfInputs(6)              'P6
    parNPD = arrayOfInputs(7)                       'Input parameters for Dynamic Nozzles
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
    parInsulationThickness = arrayOfInputs(27)
    'Text Inputs
    parId1 = arrayOfInputs(28)
    parId2 = arrayOfInputs(29)
    parNPDUnitType = arrayOfInputs(30)
    parNPDUnitType1 = arrayOfInputs(31)
    parNPDUnitType2 = arrayOfInputs(32)
    
    iOutput = 0

 ' Insert your code for output 4(FlowNozzleBodyIns)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    Dim InsulationDiameter As Double
    
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
                                            
    Dim maxDiameter As Double
    maxDiameter = ReturnMax3(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1), parFlowNozzleDiameter1)

    InsulationDiameter = maxDiameter + 2 * parInsulationThickness
    
    startPt.Set -parFacetoFace / 2, 0, 0
    endPt.Set parFacetoFace / 2, 0, 0
    Set ObjFlowNozzleBodyIns = PlaceCylinder(m_OutputColl, startPt, endPt, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowNozzleBodyIns
    Set ObjFlowNozzleBodyIns = Nothing
    
    Set m_oCodeListMetadata = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax3(A As Double, B As Double, C As Double) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    ReturnMax3 = MaxValue
End Function
