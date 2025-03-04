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
'   Creation Date:  Monday, Dec 20 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMOP1,
'   PDMS Category Name=IFRB). The symbol consists of 9 outputs: 8 physical,
'   and 1 insulation aspect outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Function ReturnMax4(A As Double, B As Double, C As Double, D As Double) As Double
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput     As Double
    Dim ObjDivergingCone As Object
    Dim ObjConvergingCone As Object
    Dim ObjFlange1 As Object
    Dim ObjFlange2 As Object
    Dim ObjInsulationBody As Object
    
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parOrificeFlangeClearance As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parOrificeFlangeClearance = arrayOfInputs(4)
    parPortRotation1 = arrayOfInputs(5)
    parPortRotation2 = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' The orgin is considered to be along the flow axis and at parFace1toCenter plus parClearanc/2 from nozzle 1
' Insert your code for output 5(InsulationThickness)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim InsulationDia As Double
    Dim flgod1 As Double
    Dim flgod2 As Double
    Dim maxDiameter As Double
'    Assumption
'    flgod1 = pipeDiam1 * 2
'    flgod2 = pipeDiam2 * 2 are taken
    flgod1 = pipeDiam1 * 2
    flgod2 = pipeDiam2 * 2
    
    maxDiameter = ReturnMax4(flgod1, flgod2, flangeDiam1, flangeDiam2)
    InsulationDia = maxDiameter + 2 * parInsulationThickness
    stPoint.Set -parOrificeFlangeClearance / 2 - parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter - parOrificeFlangeClearance / 2, 0, 0
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

