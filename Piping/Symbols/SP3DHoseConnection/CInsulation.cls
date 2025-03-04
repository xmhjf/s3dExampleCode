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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          Author:         KKC
'   Creation Date:  Tuesday, Aug 29 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details were taken from PDS Piping Component Data - Reference Guide at SN=F43_1 .
'   Symbol is created with Four Outputs - Two are Insulation aspect
'   outputs. Others are physical aspect outputs and corresponding function are as follows:
'   Nozzle1  - using 'CreateNozzleWithLength' and
'   Nozzle2  - using 'CreateNozzle' function.

'   Change History:
'   dd.mmm.yyyy     who            change description
'   -----------     -----          ------------------
'
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
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness  As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)  'O
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

' Insert your code for output 1(InsulatedBody)
    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim flangeDiam2 As Double
   

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset, depth

    If flangeThick = 0 Then flangeThick = 0.0001
    If flangeThick2 = 0 Then flangeThick2 = 0.0001

    stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    enPoint.Set stPoint.x + parFacetoFace - flangeThick2, 0, 0

    InsulationDia = ReturnMax4(pipeDiam, flangeThick, pipeDiam2, flangeThick2) _
                        + 2 * parInsulationThickness
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing

' Insert your code for output 2(Insulated Port1)
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set stPoint.x + flangeThick + parInsulationThickness, 0, 0
    InsulationDia = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDia = flangeDiam + 2 * parInsulationThickness
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

' Insert your code for output 3(Insulated Flange)
    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set stPoint.x - flangeThick2 - parInsulationThickness, 0, 0
    InsulationDia = pipeDiam2 + parInsulationThickness * 2
    If flangeDiam2 > pipeDiam2 Then InsulationDia = flangeDiam2 + 2 * parInsulationThickness
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
