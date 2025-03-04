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
'   Creation Date:  Wednesday 30, Jun 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol is Victaulic pressfit swaged (Style 505) coupling
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)

    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
    
' Insert your code for output 4 (Coupling Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    Dim dInsulationDiameter As Double
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set -parFacetoFace / 2 - parInsulationThickness, 0, 0
    enPoint.Set parFacetoFace / 2 + parInsulationThickness, 0, 0

'   Assumption: Gap between pipe surface and Coupling-end over the O-ring is taken to be 0.025 times the pipe diameter
    Dim dGap As Double
    dGap = 0.025 * pipeDiam

'   Assumption:- Section of the Coupling over the O-ring is taken to be a semi-circle with radius  equal to
'   0.05 times End-to-End dimension
    Dim dCouplingSecRadius  As Double

    dCouplingSecRadius = 0.05 * parFacetoFace

    If flangeDiam > pipeDiam Then
        dInsulationDiameter = flangeDiam + 2 * parInsulationThickness + 2 * (dGap + dCouplingSecRadius)
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness + 2 * (dGap + dCouplingSecRadius)
    End If

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                         dInsulationDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
