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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG 
'   Creation Date:  Monday, Nov 25 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the catalog available at URL http://www.resistoflex.com, all
'   assumptions are based on the catalog provided at the URL location.
'   Physical aspect consist of one Distant Spacer cylinder and Two nozzles. Insulation aspect consist of
'   a simple cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   14.Mar.2005     svsmylav       TR-49094: Added condition to ensure flange body uses maximum of
'                                          raisedFaceOrSocketDiam, flangeDiam of port1 and port 2 to compute insulation diameter.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub
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
    Dim flangeDiam2      As Double
    Dim raisedFaceOrSocketDiam1 As Double
    Dim raisedFaceOrSocketDiam2 As Double
    
    Dim iOutput     As Double
    Dim ObjDistPieceIns As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 2(Spacer Insulation)
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    Dim insulationDia As Double
    
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                                flangeDiam, sptOffset, depth, raisedFaceOrSocketDiam1
                                                                
    RetrieveParametersForThruBoltedEnds 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                                flangeDiam2, sptOffset, depth, raisedFaceOrSocketDiam2
    insulationDia = ReturnMax4(raisedFaceOrSocketDiam1, raisedFaceOrSocketDiam2, flangeDiam, flangeDiam2) + _
                        2 * parInsulationThickness
    startPoint.Set -parFacetoFace / 2, 0, 0
    endPoint.Set parFacetoFace / 2, 0, 0
    Set ObjDistPieceIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDistPieceIns
    Set ObjDistPieceIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
