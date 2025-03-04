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
'   This Symbol details were taken from PDS Piping Component Data SN=OP_573 Page # D-143
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
    Dim ObjMaintStem As Object
    Dim ObjMaintMainCylinder As Object
    Dim ObjMaintEccentricCylin As Object
    Dim ObjMaintRectBox As Object
    Dim ObjMaintLeftCylinder As Object
    Dim ObjMaintRightCylinder As Object
    
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
   

' Insert your code for output 7(Maintenance Cylinder)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth / 2, 0
    Set ObjMaintStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
 ' Insert your code for output 8(Maintenance Main Cylinder (Horizontal))
    startPoint.Set endPoint.x, endPoint.y, endPoint.z
    endPoint.Set 0, parOperatorHeight + parOperatorDepth / 2, 0
    Set ObjMaintMainCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintMainCylinder
    Set ObjMaintMainCylinder = Nothing
    
 ' Insert your code for output 9(Maintenance Eccentric Cylinder)
    startPoint.Set 0, parOperatorHeight, -parActuatorOffset
    endPoint.Set parArmLength, parOperatorHeight, -parActuatorOffset
    Set ObjMaintEccentricCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintEccentricCylin
    Set ObjMaintEccentricCylin = Nothing
    
 ' Insert your code for output 10(Maintenance Box)
    startPoint.Set parArmLength, parOperatorHeight - (parActuatorBoxDepth - parActuatorBoxDepth1), _
                        parActuatorBoxLength - parActuatorOffset
    endPoint.Set startPoint.x + parActuatorBoxWidth, startPoint.y + parActuatorBoxDepth, _
                        -parActuatorBoxLength1 - parActuatorOffset
    Set ObjMaintRectBox = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRectBox
    Set ObjMaintRectBox = Nothing
    
' Insert your code for output 11(Maintenance Left Side Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight + parActuatorCenter, -parActuatorOffset - parActuatorBoxLength1
    endPoint.Set startPoint.x, startPoint.y, -parActuatorOffset - parActuatorLength1
    Set ObjMaintLeftCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintLeftCylinder
    Set ObjMaintLeftCylinder = Nothing
    
' Insert your code for output 12(Maintenance Right Side Cylinder)
    startPoint.Set parArmLength1, parOperatorHeight + parActuatorCenter, parActuatorBoxLength - parActuatorOffset
    endPoint.Set startPoint.x, startPoint.y, parActuatorLength - parActuatorOffset
    Set ObjMaintRightCylinder = PlaceCylinder(m_OutputColl, startPoint, endPoint, parActuatorDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintRightCylinder
    Set ObjMaintRightCylinder = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
