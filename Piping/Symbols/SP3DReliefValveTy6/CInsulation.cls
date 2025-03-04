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
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT6, PDMS Category Name=IVPS)
'   The symbol consists of nine outputs, eight physical and one insulation outputs.
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
    Dim ObjValveBodyIns As Object
    
    Dim parFace1toCenter   As Double
    Dim parFace3toCenter   As Double
    Dim parInstrumentHeight    As Double
    Dim parInstrumentHeight1   As Double
    Dim parInstrumentLength    As Double
    Dim parInstrumentLength1   As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1    As Double
    Dim parInstrumentRadius    As Double
    Dim parInstrumentHeight2   As Double
    Dim parInstrumentHeight3   As Double
    Dim parInstrumentRadius1   As Double
    Dim parInstrumentHeight4   As Double
    Dim parInstrumentHeight5   As Double
    Dim parInstrumentRadius2   As Double
    Dim parInstrumentHeight6   As Double
    Dim parInstrumentHeight7   As Double
    Dim parInstrumentRadius3   As Double
    Dim parInstrumentHeight8   As Double
    Dim parInstrumentHeight9   As Double
    Dim parCylOffset   As Double
    Dim parCylOffset1  As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace3toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentRadius = arrayOfInputs(10)     'D10
    parInstrumentHeight2 = arrayOfInputs(11)    'D11
    parInstrumentHeight3 = arrayOfInputs(12)    'D12
    parInstrumentRadius1 = arrayOfInputs(13)    'D13
    parInstrumentHeight4 = arrayOfInputs(14)    'D14
    parInstrumentHeight5 = arrayOfInputs(15)    'D15
    parInstrumentRadius2 = arrayOfInputs(16)    'D16
    parInstrumentHeight6 = arrayOfInputs(17)    'D17
    parInstrumentHeight7 = arrayOfInputs(18)    'D18
    parInstrumentRadius3 = arrayOfInputs(19)    'D19
    parInstrumentHeight8 = arrayOfInputs(20)    'D20
    parInstrumentHeight9 = arrayOfInputs(21)    'D21
    parCylOffset = arrayOfInputs(22)    'D22
    parCylOffset1 = arrayOfInputs(23)   'D23
    parInsulationThickness = arrayOfInputs(24)

    iOutput = 0


'The orgin is considered to be at intersection of the nozzle axis
' Insert your code for output 6(ValveBodyIns)
    Dim maxRad As Double
    Dim maxHeight As Double
    maxRad = ReturnMax3(parInstrumentRadius, parInstrumentRadius1, parInstrumentRadius2)
    maxHeight = ReturnMax4(parInstrumentHeight3, parInstrumentHeight5, _
                                parInstrumentHeight7, parInstrumentHeight9)
    stPoint.Set -parFace1toCenter, _
                        -parFace3toCenter, _
                        -(parInstrumentWidth + 2 * maxRad + parInsulationThickness)
    enPoint.Set parFace1toCenter, _
                        parInstrumentHeight1 + maxHeight + parInsulationThickness, _
                        parInstrumentWidth1 + 2 * parInstrumentRadius3 + parInsulationThickness
    Set ObjValveBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    ReturnMax3 = MaxValue
End Function
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
