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
'   Creation Date:  Monday, Sep 16 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM1)
'The symbol consists of seven outputs, five physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjFlowMeterBody As Object
    Dim ObjVerCyl As Object
    Dim ObjFlowMeterTopBody As Object
    Dim ObjFlowMeterBodyIns As Object
    Dim ObjFlowMeterTopBodyIns As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)   'D2
    parFlowDiameter = arrayOfInputs(3)          'D3
    parInstrumentHeight = arrayOfInputs(4)       'D4
    parInstrumentDiameter = arrayOfInputs(5)     'D5
    parInstrumentWidth = arrayOfInputs(6)       'D6
    parInstrumentWidth1 = arrayOfInputs(7)      'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'Origin is considered to be mid of parFacetoFace and along the flow axis
 ' Insert your code for output 4(Insulation of Flowmeter)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also parFlowDiameter
    Dim maxDia As Double
    maxDia = ReturnMax5(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2, parFlowDiameter)
    InsulationDia = maxDia + 2 * parInsulationThickness

    stPoint.Set -parFacetoFace / 2 - parInsulationThickness, 0, 0
    enPoint.Set parFacetoFace / 2 + parInsulationThickness, 0, 0

    Set ObjFlowMeterBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterBodyIns
    Set ObjFlowMeterBodyIns = Nothing
    
 ' Insert your code for output 5(Insulation for Top Body)
    stPoint.Set -parInstrumentDiameter / 2 - parInsulationThickness, _
                        0, _
                        parInstrumentWidth + parInsulationThickness
                        
    enPoint.Set parInstrumentDiameter / 2 + parInsulationThickness, _
                        parInstrumentHeight + parInsulationThickness, _
                        -parInstrumentWidth1 - parInsulationThickness
                        
    Set ObjFlowMeterTopBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterTopBodyIns
    Set ObjFlowMeterTopBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
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
