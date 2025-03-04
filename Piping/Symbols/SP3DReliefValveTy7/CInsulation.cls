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
'   Creation Date:  Monday, Nov 11 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT7, PDMS Category Name=IVPT)
'   The symbol consists of four outputs, three physical and one insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjBox As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double           'D2
    Dim parFace2toCenter As Double           'D3
    Dim parInstrumentHeight As Double        'D4
    Dim parInstrumentLength As Double        'D6
    Dim parInstrumentLength1 As Double       'D7
    Dim parInstrumentWidth As Double         'D8
    Dim parInstrumentWidth1 As Double        'D9
    Dim parInstrumentHeight1 As Double       'D25 'As per Eden ,its used to place a box.
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInstrumentHeight = arrayOfInputs(4)
    parInstrumentLength = arrayOfInputs(5)
    parInstrumentLength1 = arrayOfInputs(6)
    parInstrumentWidth = arrayOfInputs(7)
    parInstrumentWidth1 = arrayOfInputs(8)
    parInstrumentHeight1 = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

 ' Insert your code for output 2(Insulation Body)
    stPoint.Set -parFace1toCenter, -parFace2toCenter, _
                parInstrumentWidth1 + parInsulationThickness
    enPoint.Set parInstrumentLength + parInsulationThickness, _
                parInstrumentHeight1 + parInsulationThickness, _
                -parInstrumentWidth - parInsulationThickness
    
    Set ObjInsulationBody = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody
    Set ObjInsulationBody = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
