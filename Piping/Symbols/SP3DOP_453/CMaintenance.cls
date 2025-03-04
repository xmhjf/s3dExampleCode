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
    Dim ObjMaintCyl2 As Object
    Dim ObjMaintCone As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDepth As Double
    Dim parOperatorHeight1 As Double
    Dim parOperatorDepth1 As Double
    Dim parOperatorDiameter2 As Double
    Dim parOperatorDiameter As Double
    Dim parOperatorDiameter1 As Double
    Dim parStemDia As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDepth = arrayOfInputs(3)
    parOperatorHeight1 = arrayOfInputs(4)
    parOperatorDepth1 = arrayOfInputs(5)
    parOperatorDiameter2 = arrayOfInputs(6)
    parOperatorDiameter = arrayOfInputs(7)
    parOperatorDiameter1 = arrayOfInputs(8)
    
'    parStemDia = 0.05
    iOutput = 0
     parStemDia = parOperatorDiameter1 * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05

 ' Insert your code for output 1(Vertical Stem)
    startPoint.Set 0, -parOperatorHeight1 + parOperatorDepth1, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 2(Bottom Disc)
    startPoint.Set 0, -parOperatorHeight1 + parOperatorDepth1, 0
    endPoint.Set 0, -parOperatorHeight1, 0
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
 ' Insert your code for output 3(Truncated Cone in Top)
    startPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    endPoint.Set 0, parOperatorHeight, 0
    Set ObjMaintCone = PlaceCone(m_OutputColl, startPoint, endPoint, parOperatorDiameter1 / 2, parOperatorDiameter / 2, True)
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
