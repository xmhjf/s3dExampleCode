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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Thursday, September 08 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.09.2004      MS  Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjMaintCyl1 As Object
    Dim ObjMaintCyl2 As Object
    Dim parOperatorHeight1 As Double
    Dim parOperatorHeight As Double
    Dim parLeverLength As Double
    Dim parStemDia As Double
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parLeverLength = arrayOfInputs(4)
    

    iOutput = 0
    parStemDia = parLeverLength * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05

    ' Insert your code for output (Cylinder 1)
    StartPos.Set 0, 0, 0
    EndPos.Set 0, parOperatorHeight1, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, StartPos, EndPos, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
' Insert your code for output (Cylinder 2)
    StartPos.Set 0, parOperatorHeight1, 0
    EndPos.Set 0, parOperatorHeight, 0
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, StartPos, EndPos, parLeverLength + (parOperatorHeight - parOperatorHeight1), True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
