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
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 1 2004
'   Description:
'   This Symbol details were taken from PDS Piping Component Data SN=OP_311 Page # D-133.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
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
    Dim ObjMaintStem  As IngrGeom3D.Projection3d
    Dim ObjMaintOperator As IngrGeom3D.Projection3d
    Dim ObjMaintAxisCylinder As IngrGeom3D.Projection3d
    Dim ObjMaintHandWheelCylinder As IngrGeom3D.Projection3d
        
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

' Insert your code for output 5(Maintenance Vertical Stem)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parActuatorLength + parStemExtension, 0
    Set ObjMaintStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
' Insert your code for output 6(Maintenance Horizontal Operator)
    startPoint.Set 0, parOperatorHeight - parActuatorLength, 0
    endPoint.Set 0, parOperatorHeight - parActuatorLength - parOperatorDepth, 0
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
' Insert your code for output 7(Maintenance Verical Axis Cylinder)
    startPoint.Set parActuatorOffset, parOperatorHeight - parActuatorLength, 0
    endPoint.Set parActuatorOffset, parOperatorHeight - TopCylinderTk, 0
    Set ObjMaintAxisCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parAxisDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintAxisCylinder
    Set ObjMaintAxisCylinder = Nothing
    
' Insert your code for output 8(Maintenance Top Hand Wheel)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set startPoint.x, parOperatorHeight, startPoint.z
    Set ObjMaintHandWheelCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintHandWheelCylinder
    Set ObjMaintHandWheelCylinder = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
