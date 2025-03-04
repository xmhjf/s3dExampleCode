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
    Dim ObjMaintCyl As Object
    Dim ObjMaintCone As Object
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    Dim parOperatorHeight As Double
    Dim parOperatorLength As Double
    Dim parOperatorDiameter As Double
    Dim parOperatorDiameter1 As Double
    Dim parStemDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2) 'y5
    parOperatorLength = arrayOfInputs(3) 'y6
    parOperatorDiameter = arrayOfInputs(4) 'y7
    parOperatorDiameter1 = arrayOfInputs(5) 'y8
    
    iOutput = 0
    parStemDia = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   
' Insert your code for output 1(Stem Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight + parOperatorDiameter / 2, 0
    Set ObjMaintCyl = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl
    Set ObjMaintCyl = Nothing
    
 ' Insert your code for output 2(Cone)
    startPoint.Set parStemDia / 2, parOperatorHeight, 0
    endPoint.Set parOperatorLength, parOperatorHeight, 0
    Set ObjMaintCone = PlaceCone(m_OutputColl, startPoint, endPoint, parOperatorDiameter / 2, parOperatorDiameter1 / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCone
    Set ObjMaintCone = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
