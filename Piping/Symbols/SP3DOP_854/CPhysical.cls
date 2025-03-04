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
'   Creation Date:  Monday, Dec 24 2001
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
    Dim ObjStem As Object
    Dim ObjOperatorCylin As Object
    Dim ObjEccentricShaft As Object
    Dim ObjHandWheel As Object
    
    Dim parStemAngle As Double
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
    parStemAngle = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parOperCenFromTop = arrayOfInputs(5)
    parOperatorDepth = arrayOfInputs(6)
    parArmLength = arrayOfInputs(7)
    parActuatorDiameter = arrayOfInputs(8)
    parActuatorOffset = arrayOfInputs(9)
    
'    parStemDia = 0.05
'    parShaftDia = 0.05
'    parHandWheelTk = 0.05
    
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
    
 ' Insert your code for output 1(Stem)
    startPoint.Set 0, 0, 0
    endPoint.Set -(parOperatorHeight - parOperatorDepth) * Cos(parStemAngle), _
                        (parOperatorHeight - parOperatorDepth) * Sin(parStemAngle), 0
                        
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Operator Cylinder)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set -parOperatorHeight * Cos(parStemAngle), parOperatorHeight * Sin(parStemAngle), 0
    Set ObjOperatorCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperatorCylin
    Set ObjOperatorCylin = Nothing
    
 ' Insert your code for output 3(Cylinder)
    startPoint.Set -(parOperatorHeight - parOperCenFromTop) * Cos(parStemAngle), _
                        (parOperatorHeight - parOperCenFromTop) * Sin(parStemAngle), _
                        -parActuatorOffset
    endPoint.Set -(parOperatorHeight - parOperCenFromTop) * Cos(parStemAngle) + _
                         Sin(parStemAngle) * (parArmLength - parHandWheelTk / 2), _
                        (parOperatorHeight - parOperCenFromTop) * Sin(parStemAngle) + _
                         Cos(parStemAngle) * (parArmLength - parHandWheelTk / 2), _
                         -parActuatorOffset
    Set ObjEccentricShaft = PlaceCylinder(m_OutputColl, startPoint, endPoint, parShaftDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEccentricShaft
    Set ObjEccentricShaft = Nothing
    
 ' Insert your code for output 4(Cylinder for Hand Whee;)
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set -(parOperatorHeight - parOperCenFromTop) * Cos(parStemAngle) + parArmLength * Sin(parStemAngle), _
                        (parOperatorHeight - parOperCenFromTop) * Sin(parStemAngle) + _
                         Cos(parStemAngle) * (parArmLength), startPoint.z
    Set ObjHandWheel = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
