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
'   Creation Date:  Monday, Nov 11 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR1,
'   PDMS Category Name=IFNG)
'   The symbol consists of seven outputs, five physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

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
    Dim ObjInsulationBody1 As Object
    Dim ObjInsulationBody2 As Object
    
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight1 As Double
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentDiameter2 As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)
    parInstrumentDiameter = arrayOfInputs(3)
    parInstrumentHeight1 = arrayOfInputs(4)
    parFacetoFace = arrayOfInputs(5)
    parInstrumentLength = arrayOfInputs(6)
    parInstrumentDiameter1 = arrayOfInputs(7)
    parInstrumentDiameter2 = arrayOfInputs(8)
    parInstrumentLength1 = arrayOfInputs(9)
    parInstrumentHeight2 = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0
 ' Insert your code for output 4(Insulation Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    stPoint.Set -(parInstrumentDiameter / 2 + parInstrumentLength1), 0, 0
    enPoint.Set parInstrumentLength, 0, 0
    
    InsulationDia = ReturnMax4(parInstrumentDiameter1, parInstrumentDiameter2, _
                pipeDiam, flangeDiam) + 2 * parInsulationThickness
                
    Set ObjInsulationBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody1
    Set ObjInsulationBody1 = Nothing
    
 ' Insert your code for output 5(Insulation Body2)
    stPoint.Set 0, _
                -(parInstrumentHeight1 - (parInstrumentHeight2 - parInstrumentHeight) + parInsulationThickness), _
                0
    enPoint.Set 0, _
                parInstrumentHeight2 - parInstrumentHeight1 + parInsulationThickness, _
                0
    
    InsulationDia = parInstrumentDiameter + 2 * parInsulationThickness
    
    Set ObjInsulationBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody2
    Set ObjInsulationBody2 = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
