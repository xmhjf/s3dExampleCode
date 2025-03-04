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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Thursday, Oct 11 2007

'   Description: This symbol is Bellows of MDC Vacuum Products
'   Source: http://www.mdcvacuum.com/urd/uniface.urd/ecf0070w.display?1.8.3.3

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11-Oct-07               CR-127648: Provide bellows symbol for vacuum piping
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation:" 'Used for error messages

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
    
    Dim iOutput     As Integer
    
    Dim parLength As Double
    Dim parLength2 As Double
    Dim parDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parLength = arrayOfInputs(2)
    parLength2 = arrayOfInputs(3)
    parDiameter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
    
' Place Insulation for Bellows(Output 1)
    Dim objInsBellow   As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dInsulationDia As Double
    dInsulationDia = parDiameter + (2 * parInsulationThickness)
    
    oStPoint.Set -parLength / 2, 0, 0
    oEnPoint.Set parLength / 2, 0, 0
    
    Set objInsBellow = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBellow
    Set objInsBellow = Nothing
       
' Place Insuation at Port 1(Output 5)
    Dim objInsPort1   As Object
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDia = flangeDiam + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam + (2 * parInsulationThickness)
    End If
    
    oStPoint.Set -parLength / 2, 0, 0
    oEnPoint.Set -parLength / 2 - parLength2, 0, 0
    
    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing

' Place Insuation at Port 2(Output 6)
    Dim objInsPort2   As Object
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDia = flangeDiam + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam + (2 * parInsulationThickness)
    End If
    
    oStPoint.Set parLength / 2, 0, 0
    oEnPoint.Set parLength / 2 + parLength2, 0, 0
    
    Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort2
    Set objInsPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
