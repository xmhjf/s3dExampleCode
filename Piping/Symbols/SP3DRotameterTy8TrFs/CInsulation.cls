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
'   Creation Date:  Friday, Sep 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR8,
'   PDMS Category Name=IFNN). Physical aspect consist of two nozzles,one vertical cylinder,
'   one horizontal cylinder on right side, one box on left side of the vertical cylinder
'   respectively. Two nozzles are placed one facing negative Y and the other facing
'   positive Y direction.
'   Insulation aspect consist of two cylinders one along X-axis and the other along Z-axis and
'   a box on left side of the vertical cylinder.
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput     As Double
    Dim ObjInsVerBody As Object
    Dim ObjInsSwitch As Object
    Dim ObjInsTransmitterBox As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parFace1toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double
    Dim parInstrumentHeight2 As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)           'D2
    parFlowDiameter = arrayOfInputs(3)         'D3
    parFace1toCenter = arrayOfInputs(4)        'D4
    parInstrumentHeight = arrayOfInputs(5)     'D5
    parInstrumentHeight1 = arrayOfInputs(6)    'D6
    parInstrumentWidth = arrayOfInputs(7)      'D7
    parInstrumentWidth1 = arrayOfInputs(8)     'D8
    parInstrumentLength = arrayOfInputs(9)     'D9
    parInstrumentLength1 = arrayOfInputs(10)   'D10
    parInstrumentDiameter = arrayOfInputs(11)  'D11
    parInstrumentHeight2 = arrayOfInputs(12)   'D12
    parInsulationThickness = arrayOfInputs(13)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Vertical Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

    stPoint.Set 0, -parFace1toCenter - parInsulationThickness, 0
    enPoint.Set 0, parFacetoFace - parFace1toCenter + parInsulationThickness, 0
'   Use maximum value among pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2 and parFlowDiameter
'   for computing insulation diameter
    parInsulationDiameter = ReturnMax5(pipeDiam1, flangeDiam1, _
                                       pipeDiam2, flangeDiam2, _
                                       parFlowDiameter) + 2 * parInsulationThickness
    Set ObjInsVerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVerBody
    Set ObjInsVerBody = Nothing
    
' Insert your code for output 2(Insulation for Switch)
    stPoint.Set parFlowDiameter / 2 - parInsulationThickness, _
                        parInstrumentHeight2 - parFace1toCenter, 0
    enPoint.Set parInstrumentLength1 + parInsulationThickness, _
                        parInstrumentHeight2 - parFace1toCenter, 0
    parInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    Set ObjInsSwitch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsSwitch
    Set ObjInsSwitch = Nothing
    
' Insert your code for output 3(Insulation for Transmitter Box)
    stPoint.Set -parInstrumentLength - parInsulationThickness, _
                -parInstrumentHeight - parInsulationThickness, _
                 parInstrumentWidth1 + parInsulationThickness
    enPoint.Set -parFlowDiameter / 2 + parInsulationThickness, _
                parInstrumentHeight1 - parInstrumentHeight + parInsulationThickness, _
                -parInstrumentWidth - parInsulationThickness
    Set ObjInsTransmitterBox = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTransmitterBox
    Set ObjInsTransmitterBox = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
