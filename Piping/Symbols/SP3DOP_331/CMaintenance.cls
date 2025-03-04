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
'   Author:         svsmylav
'   Creation Date:  Friday, Oct 1 2004
'   Description:
'   This Symbol details were taken from PDS Piping Component Data SN=OP_331 Page # D-133
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  14.Dec.2007     dkl  TR 128007 Rectified creation of Operator cylinder considering the operator center to top dimension inline with PDS.
'                                  used appropriate cmpdblXXX methods.
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
    Dim ObjMaintStem As Object
    Dim ObjMaintOperator As Object
    Dim ObjMaintEccentricShaft As Object
    Dim ObjMaintHandWheelCylin As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parStemExtension As Double
    Dim parStemDia As Double
    Dim parHandleDia As Double
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
    parStemExtension = arrayOfInputs(9)
    
'    parStemDia = 0.05
'    parHandleDia = 0.05
'    parHandWheelTk = 0.05
    iOutput = 0
    
    parStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(parStemDia, 0.015) Then parStemDia = 0.015
    If CmpDblGreaterthan(parStemDia, 0.05) Then parStemDia = 0.05
    
    parHandleDia = parStemDia
    If CmpDblGreaterthanOrEqualTo(parHandleDia, parActuatorDiameter / 2) Then parHandleDia = parActuatorDiameter / 2
    
    parHandWheelTk = parStemDia
    
' Insert your code for output 5(Maintenance Cylinder for Stem)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight, 0
    Set ObjMaintStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
 ' Insert your code for output 6(Maintenance Operator Cylinder)
    startPoint.Set 0, parOperatorHeight - parStemExtension - (parOperatorDepth - parOperCenFromTop), 0
    endPoint.Set 0, startPoint.y + parOperatorDepth, 0
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
 ' Insert your code for output 7(Maintenance Shaft )
    startPoint.Set 0, parOperatorHeight - parStemExtension, -parActuatorOffset
    endPoint.Set parArmLength - parHandWheelTk, parOperatorHeight - parStemExtension, _
                        -parActuatorOffset
    Set ObjMaintEccentricShaft = PlaceCylinder(m_OutputColl, startPoint, endPoint, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintEccentricShaft
    Set ObjMaintEccentricShaft = Nothing
    
 ' Insert your code for output 8(Maintenance Hand Wheel Cylinder)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set parArmLength, parOperatorHeight - parStemExtension, _
                        -parActuatorOffset
    Set ObjMaintHandWheelCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintHandWheelCylin
    Set ObjMaintHandWheelCylin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
