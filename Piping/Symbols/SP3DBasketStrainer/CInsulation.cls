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
'   Author:         BG
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol details are available in PDS Piping Data Reference Guide at page no D-106 and SN=S53AC
'    and S53AZ
'    Consist of Physical and Insulation Aspects. Physical aspect is made up of Two Nozzles, One Collar of Cylinder,
'    one Collar extension of Box and Filter body of truncated cone. Insulation aspect consist of a Cylinder
'    for Collar Insulation.
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

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjCollarIns As Object

    Dim parCollarThickness As Double
    Dim parBasketLength As Double
    Dim parBasketEndDiameter As Double
    Dim parCollarEndtoCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCollarThickness = arrayOfInputs(2)
    parBasketLength = arrayOfInputs(3)
    parBasketEndDiameter = arrayOfInputs(4)
    parCollarEndtoCenter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0

 ' Insert your code for output 4(CollarIns)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    Dim dInsulationDia As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then
        dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    startPt.Set -parCollarThickness / 2, 0, 0
    endPt.Set parCollarThickness / 2, 0, 0
    Set ObjCollarIns = PlaceCylinder(m_OutputColl, startPt, endPt, dInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarIns
    Set ObjCollarIns = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub
