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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Nov 24 2005
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is prepared based on PDS symbol MC=FEWN, SN=F4.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim CptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim CptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    Dim parFacetoFace   As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 1(Insulated Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, CptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, CptOffset2, depth2

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
    enPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
    Set ObjInsulatedBody = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2 + parInsulationThickness, pipeDiam2 / 2 + parInsulationThickness)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
 ' Insert your code for output 2(Insulation for Port1)
    Dim dInsDiameterPort As Double

    stPoint.Set -parFacetoFace / 2, 0, 0
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = 0.0001 'Check to ensure non-zero length of cylinder
        enPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
    Else
        'Following check will ensure that the insulation cylinder length is not beyond the Expander Flange
        If CmpDblLessThan(parFacetoFace, (flangeThick + parInsulationThickness)) Then
            enPoint.Set parFacetoFace / 2, 0, 0
        Else
            enPoint.Set -(parFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
        End If
    End If

    dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

 ' Insert your code for output 3(Insulation for Port1)
    stPoint.Set parFacetoFace / 2, 0, 0
    If CmpDblEqual(flangeThick2, 0) Then
        flangeThick2 = 0.0001 'Check to ensure non-zero length of cylinder
        enPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
    Else
        'Following check will ensure that the insulation cylinder length is not beyond the Expander Flange
        If CmpDblLessThan(parFacetoFace, (flangeThick2 + parInsulationThickness)) Then
            enPoint.Set -parFacetoFace / 2, 0, 0
        Else
            enPoint.Set parFacetoFace / 2 - flangeThick2 - parInsulationThickness, 0, 0
        End If
    End If
    dInsDiameterPort = flangeDiam2 + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam2, flangeDiam2) Then dInsDiameterPort = pipeDiam2 + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)

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
