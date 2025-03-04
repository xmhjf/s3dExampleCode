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
'   This Symbol details were taken from PDS Piping Component Data SN=OP_292 Page # D-132
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
    Dim ObjMaintVerticalStem As IngrGeom3D.Projection3d
    Dim ObjMaintOperatorDisc As IngrGeom3D.Projection3d
    Dim ObjMaintHoriAxis As IngrGeom3D.Projection3d
    Dim ObjMaintVeriAxis As IngrGeom3D.Projection3d
    Dim ObjHandWheelCylin As IngrGeom3D.Projection3d
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength1 As Double
    Dim parArmLength As Double
    Dim parArmDiameter As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parStemExtension As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    Dim parStemDia As Double
    Dim parVeriticalAxisTk As Double
    Dim parVeriticalAxisDia As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parOperCenFromTop = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parArmLength1 = arrayOfInputs(6)
    parArmLength = arrayOfInputs(7)
    parArmDiameter = arrayOfInputs(8)
    parActuatorLength = arrayOfInputs(9)
    parActuatorDiameter = arrayOfInputs(10)
    parStemExtension = arrayOfInputs(11)
    
'    parStemDia = 0.05
'    parVeriticalAxisDia = 0.05
'    parVeriticalAxisTk = 0.05

    iOutput = 0
          
    parStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If parStemDia < 0.015 Then parStemDia = 0.015
    If parStemDia > 0.05 Then parStemDia = 0.05
    
    parVeriticalAxisDia = parStemDia
    If parVeriticalAxisDia > parArmDiameter / 2 Then parVeriticalAxisDia = parArmDiameter / 2
    
    parVeriticalAxisTk = parStemDia
    

 ' Insert your code for output 6(Maintenance Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parActuatorLength + parStemExtension, 0
    Set ObjMaintVerticalStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintVerticalStem
    Set ObjMaintVerticalStem = Nothing
    
 ' Insert your code for output 7(Maintenance Horizontal Cylinder)
    startPoint.Set 0, parOperatorHeight - parActuatorLength + parOperCenFromTop, 0
    endPoint.Set 0, startPoint.y - parOperatorDepth, 0
    Set ObjMaintOperatorDisc = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperatorDisc
    Set ObjMaintOperatorDisc = Nothing
    
 ' Insert your code for output 8(Maintenance Horizontal Cylinder)
    startPoint.Set 0, parOperatorHeight - parActuatorLength, 0
    endPoint.Set parArmLength, startPoint.y, 0
    Set ObjMaintHoriAxis = PlaceCylinder(m_OutputColl, startPoint, endPoint, parArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintHoriAxis
    Set ObjMaintHoriAxis = Nothing
    
 ' Insert your code for output 9(Maintenance Vertical Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight - parActuatorLength + parArmDiameter / 2, 0
    endPoint.Set parArmLength1, parOperatorHeight - parVeriticalAxisTk, 0
    Set ObjMaintVeriAxis = PlaceCylinder(m_OutputColl, startPoint, endPoint, parVeriticalAxisDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintVeriAxis
    Set ObjMaintVeriAxis = Nothing
    
' Insert your code for output 10 (Maintenance HandWheel Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight, 0
    endPoint.Set parArmLength1, parOperatorHeight - parVeriticalAxisTk, 0
    Set ObjHandWheelCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelCylin
    Set ObjHandWheelCylin = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

