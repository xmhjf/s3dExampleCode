VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         dkl
'   Creation Date:  Wednesday, Jul 12 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Valve Operator RacK And Pinion (Source: VAT, Gate valve control system-Series 64,  at www.vatvalve.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentHeight3 As Double
    Dim parWithdrawalDistance As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)  'N
    parInstrumentWidth = arrayOfInputs(3)   'O
    parInstrumentHeight2 = arrayOfInputs(4) 'U
    parInstrumentLength = arrayOfInputs(5)  'P
    parInstrumentLength1 = arrayOfInputs(6) 'V
    parInstrumentHeight3 = arrayOfInputs(7) 'R
    parWithdrawalDistance = arrayOfInputs(8)
    
    iOutput = 0
    
 ' Insert your code for output 5(Maintenance Envelope)
    Dim ObjMaintEnv  As Object
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
        
    stPoint.Set -parInstrumentLength / 2, parInstrumentHeight, -parInstrumentWidth / 2
    enPoint.Set stPoint.x + parInstrumentLength, stPoint.y + parWithdrawalDistance, _
        stPoint.z + parInstrumentWidth
    
    Set ObjMaintEnv = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output 5
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintEnv
    Set ObjMaintEnv = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
