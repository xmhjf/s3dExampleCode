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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Sunday, Dec 23 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'                                             that they donot appear for small valve sizes
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  14.Dec.2007     dkl  TR 128007 Rectified creation of Operator cylinder considering the operator center to top dimension inline with PDS.
'                                  used appropriate cmpdblXXX methods.
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
    
    Dim iOutput     As Integer
    Dim ObjStem As Object
    Dim ObjOperator As Object
    Dim ObjEccentricShaft As Object
    Dim ObjHandWheelCylin As Object
    
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
    
' Insert your code for output 1(Cylinder for Stem)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Operator Cylinder)
    startPoint.Set 0, parOperatorHeight - parStemExtension - (parOperatorDepth - parOperCenFromTop), 0
    endPoint.Set 0, startPoint.y + parOperatorDepth, 0
    Set ObjOperator = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperator
    Set ObjOperator = Nothing
    
 ' Insert your code for output 3(Shaft )
    startPoint.Set 0, parOperatorHeight - parStemExtension, -parActuatorOffset
    endPoint.Set parArmLength - parHandWheelTk, parOperatorHeight - parStemExtension, _
                        -parActuatorOffset
    Set ObjEccentricShaft = PlaceCylinder(m_OutputColl, startPoint, endPoint, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEccentricShaft
    Set ObjEccentricShaft = Nothing
    
 ' Insert your code for output 4(Hand Wheel Cylinder)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set parArmLength, parOperatorHeight - parStemExtension, _
                        -parActuatorOffset
    Set ObjHandWheelCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelCylin
    Set ObjHandWheelCylin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
