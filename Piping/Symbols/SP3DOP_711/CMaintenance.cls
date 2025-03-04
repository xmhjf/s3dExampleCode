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
'   Creation Date:  Wednesday, September 22 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.22.2004      MS  Added Maintenace Envelope to the operator
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
    Dim ObjMaintBox1 As Object
    Dim ObjMaintCyl2 As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorLength1 As Double
    Dim parOperatorDepth As Double
    Dim parOperatorLength As Double
    Dim parOperatorWidth As Double
    Dim parOperatorCenter As Double
    Dim parLeverLength As Double
    Dim parStemDia As Double
    Dim parHandleDia As Double
    
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorLength1 = arrayOfInputs(3)
    parOperatorDepth = arrayOfInputs(4)
    parOperatorLength = arrayOfInputs(5)
    parOperatorWidth = arrayOfInputs(6)
    parOperatorCenter = arrayOfInputs(7)
    parLeverLength = arrayOfInputs(8)
    
    iOutput = 0
   parStemDia = parOperatorLength * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   
   parHandleDia = parStemDia

 ' Insert your code for output 4(Maintenance Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 5(Maintenance Box 1)
    startPoint.Set -parOperatorLength1, parOperatorHeight - parOperatorDepth, parOperatorWidth / 2
    endPoint.Set parOperatorLength - parOperatorLength1, parOperatorHeight, -parOperatorWidth / 2
    Set ObjMaintBox1 = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintBox1
    Set ObjMaintBox1 = Nothing
    
 ' Insert your code for output 6(Maintenance Cylinder 2)
    startPoint.Set 0, (parOperatorHeight - parOperatorDepth / 2), (parOperatorWidth + parHandleDia) / 2
    endPoint.Set parLeverLength, (parOperatorHeight - parOperatorDepth / 2), (parOperatorWidth + parHandleDia) / 2
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
       
    Set startPoint = Nothing
    Set endPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
