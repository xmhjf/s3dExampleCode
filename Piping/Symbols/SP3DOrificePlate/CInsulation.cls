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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IOP, PDMS Category Name=IFOE)
'   The symbol consists of six outputs, four physical and two insulation outputs.
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
    Dim ObjOrificePlateBody As Object
    Dim ObjPaddleBody As Object
    Dim ObjOrificePlateIns As Object
    Dim ObjPaddleIns As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)           'D2
    parInstrumentDiameter = arrayOfInputs(3)   'D3
    parInstrumentHeight = arrayOfInputs(4)     'D4
    parInstrumentWidth = arrayOfInputs(5)      'D5
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
'The origin is considered to be mid of parFacetoFace and on the flow axis
 ' Insert your code for output 3(Orifice Plate Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    InsulationDia = ReturnMax5(parInstrumentDiameter, pipeDiam1, flangeDiam1, _
                                        pipeDiam2, flangeDiam2) + _
                                        2 * parInsulationThickness
    Set ObjOrificePlateIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOrificePlateIns
    Set ObjOrificePlateIns = Nothing
    
 ' Insert your code for output 4(Paddle Insulation)                      
    stPoint.Set -parFacetoFace / 2 - parInsulationThickness, _
                        0, _
                        -parInstrumentWidth / 2 - parInsulationThickness

    enPoint.Set parFacetoFace / 2 + parInsulationThickness, _
                        parInstrumentDiameter / 2 + parInstrumentHeight + parInsulationThickness, _
                        parInstrumentWidth / 2 + parInsulationThickness
    Set ObjPaddleIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPaddleIns
    Set ObjPaddleIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
