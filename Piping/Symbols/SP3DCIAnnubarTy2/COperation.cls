VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          BG
'   Creation Date:  Friday, September 19 - 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IA2,
'   PDMS Category Name=IFAH). The symbol consists of five oupputs: Three physical,
'   and two insulation aspect outputs.
'
'   Note: This symbol expects the user to place an Olet component on the pipe and this (Annubar) should be placed
'   at the connect point2  of the Olet.The parameters values are to be taken from connect point2 of the Olet.'
'   The port for the symbol is fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   05.Sep.2003     BG   TR-CP 49538  Created new class file COperation.cls to add operation aspect to the Symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Operation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim parInstrumentLength As Double
    Dim parFacetoEnd        As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1    As Double
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
    
    Dim parOperationHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)  'D2
    parFacetoEnd = arrayOfInputs(3)         'D3
    parInstrumentHeight = arrayOfInputs(4)  'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
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
    parOperationHeight = arrayOfInputs(20)
    parId1 = arrayOfInputs(21)
    parNPDUnitType = arrayOfInputs(22)
    parNPDUnitType1 = arrayOfInputs(23)
        
    iOutput = 0
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output (Operation aspect)
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

    Dim ObjOpEnvelope As Object
    stPoint.Set parFacetoEnd, -parInstrumentHeight1, -pipeDiam(1) / 2 - parInsulationThickness
    enPoint.Set parFacetoEnd + parOperationHeight, parInstrumentHeight, pipeDiam(1) / 2 + parInsulationThickness
    Set ObjOpEnvelope = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOpEnvelope
    Set ObjOpEnvelope = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

