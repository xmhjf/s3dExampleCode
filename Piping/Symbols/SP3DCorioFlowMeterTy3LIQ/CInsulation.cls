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
'   Creation Date:  Tuesday, Oct 8 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM4)
'   The symbol consists of nine outputs, six  physical and two insulation outputs.
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
    Dim ObjFlowmeterBody As Object
    Dim ObjLargeBoxBody As Object
    Dim ObjSmallBoxBody As Object
    Dim ObjConnectingBody As Object
    Dim ObjFlowmeterBodyIns As Object
    Dim ObjBoxBodyIns As Object
    Dim ObjConnectingBodyIns As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentLength1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)        'D2
    parFlowDiameter = arrayOfInputs(3)              'D3
    parInstrumentHeight = arrayOfInputs(4)        'D4
    parInstrumentLength = arrayOfInputs(5)       'D5
    parInstrumentHeight1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)        'D7
    parInstrumentHeight2 = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)                    'D9
    parInstrumentHeight3 = arrayOfInputs(10)      'D10
    parInstrumentLength1 = arrayOfInputs(11)       'D11
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0
 'Origin is considered to be mid of parFacetoFace and along the flow axis
 ' Insert your code for output 5(Flowmeter Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also parFlowDiameter
    Dim maxDia As Double
    maxDia = ReturnMax5(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2, parFlowDiameter)
    parInsulationDiameter = maxDia + 2 * parInsulationThickness
    stPoint.Set -parFacetoFace / 2 - parInsulationThickness, 0, 0
    enPoint.Set parFacetoFace / 2 + parInsulationThickness, 0, 0
    Set ObjFlowmeterBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowmeterBodyIns
    Set ObjFlowmeterBodyIns = Nothing
    
 ' Insert your code for output 6(Box Body Insulation)
    stPoint.Set -parInstrumentLength / 2 - parInsulationThickness, _
                        -parInstrumentHeight2 - parInsulationThickness, _
                        parInstrumentWidth1 + parInsulationThickness
    enPoint.Set parInstrumentLength / 2 + parInsulationThickness, _
                        -parInstrumentHeight + parInstrumentHeight1 + parInsulationThickness, _
                        -parInstrumentWidth / 2 - parInsulationThickness
    Set ObjBoxBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoxBodyIns
    Set ObjBoxBodyIns = Nothing
    
 ' Insert your code for output 7(Connecting Body Insulation)
    Dim meetHeight As Double
    Dim verCylDia As Double
'Diameter of Vertical cylinder is assumed to be 0.2 times the parFlowDiameter
    verCylDia = 0.2 * parFlowDiameter
    meetHeight = Sqr((parFlowDiameter / 2 * parFlowDiameter / 2) - (verCylDia / 2 * verCylDia / 2))
    stPoint.Set 0, -parInstrumentHeight + parInstrumentHeight1, 0
    enPoint.Set 0, -meetHeight, 0
    Set ObjConnectingBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, verCylDia + 2 * parInsulationThickness, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnectingBodyIns
    Set ObjConnectingBodyIns = Nothing
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
