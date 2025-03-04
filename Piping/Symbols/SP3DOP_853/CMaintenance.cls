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
    Dim ObjMaintCyl3 As Object
    Dim ObjMaintCyl4 As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parStemDia As Double
    Dim parShaftDia As Double
    Dim parHandWheelTk As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parOperCenFromTop = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parArmLength = arrayOfInputs(6)
    parActuatorDiameter = arrayOfInputs(7)
    parActuatorOffset = arrayOfInputs(8)
    
    iOutput = 0
    
    parStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If parStemDia < 0.015 Then parStemDia = 0.015
    If parStemDia > 0.05 Then parStemDia = 0.05
    
    parShaftDia = parStemDia
    If parShaftDia >= parActuatorDiameter Then
        parShaftDia = parActuatorDiameter / 2
    End If
    parHandWheelTk = parShaftDia
    
 ' Insert your code for output 5(Maintenance Cylinder 1)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
 ' Insert your code for output 6(Maintenance Cylinder 2)
    startPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    endPoint.Set 0, parOperatorHeight, 0
    Set ObjMaintCyl2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl2
    Set ObjMaintCyl2 = Nothing
    
 ' Insert your code for output 7(Maintenance Cylinder 3)
    startPoint.Set 0, parOperatorHeight - parOperCenFromTop, -parActuatorOffset
    endPoint.Set parArmLength - parHandWheelTk, parOperatorHeight - parOperCenFromTop, -parActuatorOffset
    Set ObjMaintCyl3 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parShaftDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl3
    Set ObjMaintCyl3 = Nothing
    
 ' Insert your code for output 8(Maintenance Cylinder 4)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set parArmLength, parOperatorHeight - parOperCenFromTop, -parActuatorOffset
    Set ObjMaintCyl4 = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl4
    Set ObjMaintCyl4 = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
