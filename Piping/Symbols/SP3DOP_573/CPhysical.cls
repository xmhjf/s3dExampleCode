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
'   Creation Date:  Sunday, Dec 23 2001
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
    Dim ObjMainCylinder As Object
    Dim ObjEccentricCylin As Object
    Dim ObjRectBox As Object
    Dim ObjLeftCylinder As Object
    Dim ObjRightCylinder As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parActuatorOffset As Double
    Dim parArmLength As Double
    Dim parArmDiameter As Double
    Dim parArmLength1 As Double
    Dim parActuatorBoxLength As Double
    Dim parActuatorBoxLength1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorBoxWidth As Double
    Dim parActuatorBoxDepth1 As Double
    Dim parActuatorBoxDepth As Double
    Dim parActuatorCenter As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorDiameter1 As Double
    Dim parStemDia As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parOperCenFromTop = arrayOfInputs(4)
    parOperatorDepth = arrayOfInputs(5)
    parActuatorOffset = arrayOfInputs(6)
    parArmLength = arrayOfInputs(7)
    parArmDiameter = arrayOfInputs(8)
    parArmLength1 = arrayOfInputs(9)
    parActuatorBoxLength = arrayOfInputs(10)
    parActuatorBoxLength1 = arrayOfInputs(11)
    parActuatorLength = arrayOfInputs(12)
    parActuatorLength1 = arrayOfInputs(13)
    parActuatorBoxWidth = arrayOfInputs(14)
    parActuatorBoxDepth1 = arrayOfInputs(15)
    parActuatorBoxDepth = arrayOfInputs(16)
    parActuatorCenter = arrayOfInputs(17)
    parActuatorDiameter = arrayOfInputs(18)
    parActuatorDiameter1 = arrayOfInputs(19)

'    parStemDia = 0.05
    iOutput = 0
    
    parStemDia = parOperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   

' Insert your code for output 1(Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth / 2, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(Main Cylinder (Horizontal))
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set 0, parOperatorHeight + parOperatorDepth / 2, 0
    Set ObjMainCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCylinder
    Set ObjMainCylinder = Nothing
    
 ' Insert your code for output 3(Eccentric Cylinder)
    startPoint.Set 0, parOperatorHeight, -parActuatorOffset
    endPoint.Set parArmLength, parOperatorHeight, -parActuatorOffset
    Set ObjEccentricCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEccentricCylin
    Set ObjEccentricCylin = Nothing
    
 ' Insert your code for output 4(Box)
    startPoint.Set parArmLength, parOperatorHeight - (parActuatorBoxDepth - parActuatorBoxDepth1), _
                        parActuatorBoxLength - parActuatorOffset
    endPoint.Set startPoint.x + parActuatorBoxWidth, startPoint.y + parActuatorBoxDepth, _
                        -parActuatorBoxLength1 - parActuatorOffset
    Set ObjRectBox = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectBox
    Set ObjRectBox = Nothing
    
' Insert your code for output 5(Left Side Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight + parActuatorCenter, -parActuatorOffset - parActuatorBoxLength1
    endPoint.Set startPoint.x, startPoint.y, -parActuatorOffset - parActuatorLength1
    Set ObjLeftCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCylinder
    Set ObjLeftCylinder = Nothing
    
' Insert your code for output 5(Right Side Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight + parActuatorCenter, parActuatorBoxLength - parActuatorOffset
    endPoint.Set startPoint.x, startPoint.y, parActuatorLength - parActuatorOffset
    Set ObjRightCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCylinder
    Set ObjRightCylinder = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
