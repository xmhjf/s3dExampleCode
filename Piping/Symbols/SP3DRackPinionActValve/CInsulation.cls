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
'   Author:          SSP
'   Creation Date:  Tuesday, Oct 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

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
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick      As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2       As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjValCylinder As Object
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjStem As Object
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double              'D2
    Dim parFacetoFace As Double                 'D3
    Dim parCylHeight As Double               'D4
    Dim parActuatorHeight As Double             'D5
    Dim parActuatorHeight1 As Double            'D6
    Dim parActuatorHeight2 As Double            'D7
    Dim parActuatorLength As Double             'D8
    Dim parActuatorLength1 As Double            'D9
    Dim parActuatorWidth As Double              'D10
    Dim parActuatorWidth1 As Double             'D11
    Dim parActuatorLength2 As Double            'D12
    Dim parActuatorLength3 As Double            'D13
    Dim parActuatorWidth2 As Double             'D14
    Dim parActuatorWidth3 As Double             'D15
    Dim parActuatorWidth4 As Double             'D16
    Dim parActuatorLength4 As Double            'D17
    Dim parActuatorLength5 As Double            'D18
    Dim parActuatorHeight3 As Double            'D19
    Dim parActuatorHeight4 As Double            'D20
    Dim parActuatorHeight5 As Double            'D21
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorHeight1 = arrayOfInputs(6)
    parActuatorHeight2 = arrayOfInputs(7)
    parActuatorLength = arrayOfInputs(8)
    parActuatorLength1 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorLength2 = arrayOfInputs(12)
    parActuatorLength3 = arrayOfInputs(13)
    parActuatorWidth2 = arrayOfInputs(14)
    parActuatorWidth3 = arrayOfInputs(15)
    parActuatorWidth4 = arrayOfInputs(16)
    parActuatorLength4 = arrayOfInputs(17)
    parActuatorLength5 = arrayOfInputs(18)
    parActuatorHeight3 = arrayOfInputs(19)
    parActuatorHeight4 = arrayOfInputs(20)
    parActuatorHeight5 = arrayOfInputs(21)
    parInsulationThickness = arrayOfInputs(22)
    
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 8(InsulationBody)
    Dim CylEdgeToCen As Double
    Dim CylDia As Double
    CylDia = parFace1toCenter
    CylEdgeToCen = Sqr(parCylHeight ^ 2 + (CylDia / 2) ^ 2)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxRad refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also CylEdgeToCen
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam1 / 2, pipeDiam2 / 2, flangeDiam1 / 2, flangeDiam2 / 2, CylEdgeToCen)
    InsulationDia = 2 * maxRad + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    Set ObjInsulationBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

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
