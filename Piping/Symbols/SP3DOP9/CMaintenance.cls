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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
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
'   22.Feb.2007         KKC           TR-112500 Maintenance envelopes for some valve operators do not use capped ends
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
    Dim OperatorHeight As Double
    Dim LeverLength As Double
    Dim LeverLength1 As Double
    Dim parStemDia As Double

  
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    OperatorHeight = arrayOfInputs(2)
    LeverLength = arrayOfInputs(3)
    LeverLength1 = arrayOfInputs(4)
    

    iOutput = 0
    
     parStemDia = LeverLength1 * 0.1
'   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
  
   ' Insert your code for output 1(Cylinder)
    Dim ObjMaintCyl1 As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set 0, OperatorHeight, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parStemDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing

 ' Insert your code for output 2(Cylinder)
    Dim ObjMaintCyl2 As Object
    Dim diam As Double
        If LeverLength > LeverLength1 Then
        diam = LeverLength * 2
    Else
        diam = LeverLength1 * 2
    End If
    
    stPoint.Set 0, OperatorHeight - parStemDia, 0
    enPoint.Set 0, OperatorHeight, 0
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, diam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
