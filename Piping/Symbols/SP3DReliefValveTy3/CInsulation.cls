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
'   Author:          BG
'   Creation Date:  Tuesday, Oct 1 2002
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT3,
'   PDMS Category Name=IVPP). Physical aspect consist of Relief valve body of Box shape, 1 Cylinder
'   Nozzle-1 facing (-)ive side of X-Axis and Nozzle-2 facing (-)ive side of Y-axis.
'   Insulation aspect consist of a single box place covering the whole valve.
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
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)                     'D2
    parFace2toCenter = arrayOfInputs(3)                     'D3
    parInstrumentHeight = arrayOfInputs(4)               'D4
    parInstrumentHeight1 = arrayOfInputs(5)             'D5
    parInstrumentLength = arrayOfInputs(6)              'D6
    parInstrumentLength1 = arrayOfInputs(7)            'D7
    parInstrumentWidth = arrayOfInputs(8)               'D8
    parInstrumentWidth1 = arrayOfInputs(9)             'D9
    parInstrumentHeight2 = arrayOfInputs(10)          'D10
    parInstrumentDiameter = arrayOfInputs(11)      'D11
    parInsulationThickness = arrayOfInputs(12)

    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
'Insert your code for output 1(Insulation Body)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    stPoint.Set parInstrumentLength + parInsulationThickness, _
                            parInstrumentHeight1, -(parInstrumentWidth + parInsulationThickness)
    
    enPoint.Set -parFace1toCenter, -parFace2toCenter, parInstrumentWidth1 + parInsulationThickness
    
    Set ObjInsBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ObjInsBody = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
