VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Dec 20 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'					      that they donot appear for small valve sizes
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjStem  As IngrGeom3D.Projection3d
    Dim ObjOperator As IngrGeom3D.Projection3d
    Dim ObjAxisCylinder As IngrGeom3D.Projection3d
    Dim ObjHandWheelCylinder As IngrGeom3D.Projection3d
        
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parOperatorDepth As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parStemExtension As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    Dim parStemDia As Double
    Dim parAxisDia As Double
    Dim TopCylinderTk As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parActuatorLength = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorOffset = arrayOfInputs(7)
    parStemExtension = arrayOfInputs(8)
    
'    parStemDia = 0.05
'    parAxisDia = 0.05
'    TopCylinderTk = 0.05
    iOutput = 0
    
    parStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If parStemDia < 0.015 Then parStemDia = 0.015
    If parStemDia > 0.05 Then parStemDia = 0.05
    
    parAxisDia = parStemDia
    If parAxisDia >= parActuatorDiameter / 2 Then parAxisDia = parActuatorDiameter / 2
    
    TopCylinderTk = parStemDia

' Insert your code for output 1(Vertical Stem)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parActuatorLength + parStemExtension, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 2(Horizontal Operator)
    startPoint.Set 0, parOperatorHeight - parActuatorLength, 0
    endPoint.Set 0, parOperatorHeight - parActuatorLength - parOperatorDepth, 0
    Set ObjOperator = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperator
    Set ObjOperator = Nothing
    
' Insert your code for output 3(Verical Axis Cylinder)
    startPoint.Set parActuatorOffset, parOperatorHeight - parActuatorLength, 0
    endPoint.Set parActuatorOffset, parOperatorHeight - TopCylinderTk, 0
    Set ObjAxisCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parAxisDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAxisCylinder
    Set ObjAxisCylinder = Nothing
    
' Insert your code for output 4(Top Hand Wheel)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set startPoint.x, parOperatorHeight, startPoint.z
    Set ObjHandWheelCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelCylinder
    Set ObjHandWheelCylinder = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
