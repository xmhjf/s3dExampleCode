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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIV,
'   PDMS Category Name=IEEF). Physical aspect consist of four nozzles and one revolved linestring.
'   Insulation aspect consist of single entity which is a cylinder.
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
    Dim ObjInsBody As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentDiameter2 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)            'D2 is equal to D3 + D4 + D5, so NOT used.
    parInstrumentLength = arrayOfInputs(3)      'D3
    parInstrumentLength1 = arrayOfInputs(4)     'D4
    parInstrumentLength2 = arrayOfInputs(5)     'D5
    parInstrumentDiameter = arrayOfInputs(6)    'D6
    parInstrumentDiameter1 = arrayOfInputs(7)   'D7
    parInstrumentDiameter2 = arrayOfInputs(8)   'D8
    parPortRotation1 = arrayOfInputs(9)                 'D9
    parPortRotation2 = arrayOfInputs(10)               'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Body)
    stPoint.Set -parInstrumentLength1 / 2 - parInstrumentLength - parInsulationThickness, 0, 0
    enPoint.Set parInstrumentLength1 / 2 + parInstrumentLength2 + parInsulationThickness, 0, 0
    parInsulationDiameter = parInstrumentDiameter2 + 2 * parInsulationThickness
    If parInstrumentDiameter > parInstrumentDiameter2 Then
        parInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    End If
    Set ObjInsBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    Set ObjInsBody = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
