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
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MS
'   Creation Date:  September 20, 2004
'   Description:
'   The Details were taken from PDS Piping Component Data - Reference Guide SN=F134

'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
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

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjInsulatedBody1 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedBody2 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort3 As Object


'   Variables defined by svsmylav
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsDiameterPort      As Double
    Dim dInsulationDia  As Double
    Dim dInsulationDia2  As Double

    Dim parFacetoFace As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

' Insert your code for output 1(Insulated Body (Port1 side))
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing

' Insert your code for output 2(Insulated Port1)
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -(parFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
    dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing

' Insert your code for output 3(Insulated Body (Port2 side))
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    enPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
    stPoint.Set 0, 0, 0
    dInsulationDia2 = pipeDiam + parInsulationThickness * 2

    Set ObjInsulatedBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, dInsulationDia / 2, _
                                      dInsulationDia2 / 2, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing

' Insert your code for output 4(Insulated Port2)
    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
    dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing

' Insert your code for output 5(Insulated Branch (Port3 side))
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, parFace3toCenter - flangeThick, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing

' Insert your code for output 6(Insulated Port3)
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, parFace3toCenter - flangeThick - parInsulationThickness, 0
    dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
End Sub
