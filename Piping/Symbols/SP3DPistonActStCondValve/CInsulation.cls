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
'   Author:          svsmylav svsmylav
'   Creation Date:  Wednesday, Oct 30 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ISCV3).
'   The symbol consists of  Ten outputs, Seven physical outputs and
'   Three Insulation Outputs.
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
    Dim pipeDiam3        As Double
    Dim flangeDiam3      As Double
    
    Dim iOutput     As Double
    Dim ObjValveBodyIns As Object
    Dim ObjVertBodyIns  As Object
    Dim ObjNozzle3Ins   As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parNozzleOffset As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorDiameter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFacetoFace = arrayOfInputs(3)    'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parActuatorHeight = arrayOfInputs(5)    'D5
    parNozzleOffset = arrayOfInputs(6)  'D6
    parActuatorHeight1 = arrayOfInputs(7)   'D7
    parActuatorDiameter = arrayOfInputs(8)  'D8
    parFace3toCenter = arrayOfInputs(9) 'D9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Valve Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double

    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    CylDia = parFace1toCenter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'   maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam1 / 2, pipeDiam2 / 2, flangeDiam1 / 2, flangeDiam2 / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjValveBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
' Insert your code for output 2(Insulation for Vertical Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight, 0
    InsulationDia = CylDia + 2 * parInsulationThickness
    Set ObjVertBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertBodyIns
    Set ObjVertBodyIns = Nothing
    
' Insert your code for output 3(Insulation for Nozzle3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset, depth

    stPoint.Set 0, parNozzleOffset, 0
    enPoint.Set parFace3toCenter, parNozzleOffset, 0
    InsulationDia = pipeDiam3 + 2 * parInsulationThickness
    If pipeDiam3 < flangeDiam3 Then InsulationDia = flangeDiam3 + 2 * parInsulationThickness
    Set ObjNozzle3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3Ins
    Set ObjNozzle3Ins = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

