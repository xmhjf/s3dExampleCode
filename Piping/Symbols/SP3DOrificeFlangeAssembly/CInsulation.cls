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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is prepared based on Resistoflex Flanged Lined Pipe & Fittings catalog.
'   The symbol consists of Thirteen outputs, Eight physical and  Five  Insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------     -----                  ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC                     DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG                     TR-124959 Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Dim PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Integer
    Dim ObjOriFlAssmBodyIns As Object
    Dim ObjPort3BodyIns As Object
    Dim ObjPort3Ins As Object
    Dim ObjPort4BodyIns As Object
    Dim ObjPort4Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parOrificeFlangeThk As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parOrificeFlangeThk = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 ' Insert your code for output 5(Orifice Flange Assembly Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    InsulationDia = maxDia + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjOriFlAssmBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriFlAssmBodyIns
    Set ObjOriFlAssmBodyIns = Nothing
    
 ' Insert your code for output 6(Port3 Body Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
'The Orifice Flange Diameter is taken to be 0.7 times the flange diameter of Port1.
    Dim OrificeFlangeBodyDia As Double
    OrificeFlangeBodyDia = 0.7 * flangeDiam1
    stPoint.Set parOrificeFlangeThk / 2, _
                        (parFace3toCenter - flangeThick) * Cos(PI / 4), _
                        -(parFace3toCenter - flangeThick) * Sin(PI / 4)
    enPoint.Set parOrificeFlangeThk / 2, _
                            OrificeFlangeBodyDia / 2 * Cos(PI / 4), _
                            -OrificeFlangeBodyDia / 2 * Sin(PI / 4)
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3BodyIns
    Set ObjPort3BodyIns = Nothing
    
 ' Insert your code for output 7(Port3 Insulation)
    stPoint.Set parOrificeFlangeThk / 2, _
                        parFace3toCenter * Cos(PI / 4), _
                        -parFace3toCenter * Sin(PI / 4)
    enPoint.Set parOrificeFlangeThk / 2, _
                        (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 4), _
                        -(parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 4)
    InsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthanOrEqualTo(pipeDiam, flangeDiam) Then InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Ins
    Set ObjPort3Ins = Nothing
    
 ' Insert your code for output 8(Port4 Body Insulation)
'The Face3 to center is assumed to be equal to Face4 to center
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set -parOrificeFlangeThk / 2, _
                        (parFace3toCenter - flangeThick) * Cos(PI / 4), _
                        (parFace3toCenter - flangeThick) * Sin(PI / 4)
    enPoint.Set -parOrificeFlangeThk / 2, _
                        OrificeFlangeBodyDia / 2 * Cos(PI / 4), _
                        OrificeFlangeBodyDia / 2 * Sin(PI / 4)
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort4BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4BodyIns
    Set ObjPort4BodyIns = Nothing
    
 ' Insert your code for output 9(Port4 Insulation)
    stPoint.Set -parOrificeFlangeThk / 2, _
                        parFace3toCenter * Cos(PI / 4), _
                        parFace3toCenter * Sin(PI / 4)
    enPoint.Set -parOrificeFlangeThk / 2, _
                        (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(PI / 4), _
                        (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(PI / 4)
    InsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthanOrEqualTo(pipeDiam, flangeDiam) Then InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort4Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4Ins
    Set ObjPort4Ins = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double
    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
