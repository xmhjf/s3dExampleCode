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
    Dim ObjMaintCyl As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorHeight1 As Double
    Dim parHandWheelDiameter As Double
    Dim parBasePlateDia As Double

'' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parHandWheelDiameter = arrayOfInputs(4)
    parBasePlateDia = arrayOfInputs(5)
 
    iOutput = 0
    
' Insert your code for output 5(Maintenance Cylinder)
   'Assumed Hand Wheel Cross-section Circuar dia as 10mm(thickness)
   Dim dHandWheelthickness  As Double
   dHandWheelthickness = 0.01
   Dim stPoint As New AutoMath.DPosition
   Dim enPoint As New AutoMath.DPosition
   stPoint.Set 0, parOperatorHeight - parOperatorHeight1, 0
   enPoint.Set 0, parOperatorHeight + dHandWheelthickness / 2, 0
   
   Set ObjMaintCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parHandWheelDiameter + dHandWheelthickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl
    Set ObjMaintCyl = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
