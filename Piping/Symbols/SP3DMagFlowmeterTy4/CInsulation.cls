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
'   Creation Date:  Wednesday, Sep 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMFM4,
'   PDMS Category Name=IFMV). The symbol consists of five outputs, four physical and
'   one insulation outputs.  Symbol insulation consists of large box that surrounds all
'   physical aspect outputs.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsFlowmeter As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)         'D2
    parInstrumentLength = arrayOfInputs(3)    'D3
    parInstrumentWidth = arrayOfInputs(4)     'D4
    parInstrumentHeight = arrayOfInputs(5)    'D5
    parInstrumentHeight1 = arrayOfInputs(6)   'D6
    parInstrumentLength1 = arrayOfInputs(7)   'D7
    parInstrumentHeight2 = arrayOfInputs(8)   'D8
    parInstrumentWidth1 = arrayOfInputs(9)    'D9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Flowmeter)
    stPoint.Set -parFacetoFace / 2, _
            -parInstrumentHeight - parInsulationThickness, _
            -parInstrumentWidth / 2 - parInsulationThickness
    enPoint.Set parFacetoFace / 2, _
            parInstrumentHeight2 + parInsulationThickness, _
            parInstrumentWidth / 2 + parInsulationThickness
    Set ObjInsFlowmeter = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFlowmeter
    Set ObjInsFlowmeter = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
