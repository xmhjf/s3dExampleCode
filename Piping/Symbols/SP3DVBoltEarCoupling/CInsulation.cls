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
'   Author:         kkk
'   Creation Date:  Monday, Nov 21, 2005
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    Dim parInsulationThickness As Double
    Dim parClampWidth As Double
    Dim parInsClampOuterDiameter As Double
    Dim parInsClampEarOuterDiameter As Double
    Dim parNumberofBoltEars As Integer

' Inputs

    Set oPartFclt = arrayOfInputs(1)
    parClampWidth = arrayOfInputs(2)
    parInsClampOuterDiameter = arrayOfInputs(3)
    parInsClampEarOuterDiameter = arrayOfInputs(4)
    parNumberofBoltEars = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
     Dim stPoint   As New AutoMath.DPosition
     Dim enPoint   As New AutoMath.DPosition
     
     stPoint.Set -parClampWidth / 2, 0, 0
     enPoint.Set parClampWidth / 2, 0, 0
    
    
     Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                        parInsClampEarOuterDiameter + parInsulationThickness * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
