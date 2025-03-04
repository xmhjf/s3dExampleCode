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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Insulation Diameter is greater of Flange and Pipe Dia of both Pipes mating with the Valve
'   Symbol Model No. is: I13AZ Page No. 357 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjBallInsulation As Object

    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

 ' Insert your code for output 6(As a Cylinder)
    Dim InsulationDia As Double
    Dim InsulationDia1 As Double
    Dim InsulationDia2 As Double
    Dim StartInsulationPos       As New AutoMath.DPosition
    Dim EndInsulationPos         As New AutoMath.DPosition
    StartInsulationPos.Set -parFacetoCenter, 0, 0
    EndInsulationPos.Set parFace1toCenter, 0, 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDia1 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia1 = pipeDiam + parInsulationThickness * 2
    End If
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDia2 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDia2 = pipeDiam + parInsulationThickness * 2
    End If
    If InsulationDia1 > InsulationDia2 Then
        InsulationDia = InsulationDia1
    Else
        InsulationDia = InsulationDia2
    End If
    Set ObjBallInsulation = PlaceCylinder(m_OutputColl, StartInsulationPos, EndInsulationPos, _
                                    InsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBallInsulation
    Set ObjBallInsulation = Nothing
    Set StartInsulationPos = Nothing
    Set EndInsulationPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
