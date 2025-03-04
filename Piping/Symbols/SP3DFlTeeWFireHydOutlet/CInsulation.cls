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
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
'  This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Flanged T with fire hydrant outlet
'   The symbol consists of seven outputs, four physical and three insulation outputs.
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
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim flangeDiam4      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulationBody1 As Object
    Dim ObjInsulationBody2 As Object
    Dim ObjInsulationBody3 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parFace4toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
 ' Insert your code for output 2(Insulation Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    Dim InsulationDia As Double
    Dim maxDiameter As Double
    maxDiameter = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)

    InsulationDia = maxDiameter + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace1toCenter, 0, 0
    
    Set ObjInsulationBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody1
    Set ObjInsulationBody1 = Nothing
    
 ' Insert your code for output 2(Insulation Body2)
    Dim MeetPoint As Double
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset, depth
    MeetPoint = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam3 / 2) ^ 2)
    
    If pipeDiam3 > flangeDiam3 Then
        InsulationDia = pipeDiam3 + 2 * parInsulationThickness
    Else
        InsulationDia = flangeDiam3 + 2 * parInsulationThickness
    End If

    stPoint.Set 0, 0, -parFace3toCenter
    enPoint.Set 0, 0, -MeetPoint
    
    Set ObjInsulationBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody2
    Set ObjInsulationBody2 = Nothing
    
' Insert your code for output 3(Insulation Body3)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, sptOffset, depth
    MeetPoint = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam4 / 2) ^ 2)

    If pipeDiam4 > flangeDiam4 Then
        InsulationDia = pipeDiam4 + 2 * parInsulationThickness
    Else
        InsulationDia = flangeDiam4 + 2 * parInsulationThickness
    End If
    
    stPoint.Set 0, parFace4toCenter, 0
    enPoint.Set 0, MeetPoint, 0
    
    Set ObjInsulationBody3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody3
    Set ObjInsulationBody3 = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing


    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
