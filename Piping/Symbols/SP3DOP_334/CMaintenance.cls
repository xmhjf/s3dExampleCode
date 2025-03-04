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
'   This Symbol details were taken from PDS Piping Component Data SN=OP_334
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages
Const MIN_LENGTH = 0.0001

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    
    Dim ObjMaintStem As Object
    Dim ObjMaintFixedActShaft As Object
    Dim ObjMaintOPCyl As Object
    Dim ObjMaintStemExtn As Object
    Dim ObjMaintVarLenActShaft As Object
    Dim ObjMaintActuatorCylinder As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parArmLength1 As Double
    Dim parStemExtension As Double
    
    Dim dStemDia As Double
    Dim dHandleDia As Double
    Dim dHandWheelTk As Double
    
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2) ''Y4
    parOperatorDiameter = arrayOfInputs(3) ''Y5
    parOperCenFromTop = arrayOfInputs(4) ''Y6
    parOperatorDepth = arrayOfInputs(5) ''Y7
    parArmLength = arrayOfInputs(6) ''Y8
    parActuatorDiameter = arrayOfInputs(7) ''Y9
    parActuatorOffset = arrayOfInputs(8) ''Y10
    parArmLength1 = arrayOfInputs(9) ''Y11
    parStemExtension = arrayOfInputs(10) ''Y12
    
    iOutput = 0

    dStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If dStemDia < 0.015 Then dStemDia = 0.015
    If dStemDia > 0.05 Then dStemDia = 0.05
    
    dHandleDia = dStemDia
    If dHandleDia >= parActuatorDiameter / 2 Then dHandleDia = parActuatorDiameter / 2
    
    dHandWheelTk = dStemDia
    
' Insert your code for output 7(Maintenance Stem )
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintStem = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
 ' Insert your code for output 8(Maintenance Operator Cylinder)
    oStartPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    oEndPoint.Set 0, parOperatorHeight, 0
    Set ObjMaintOPCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOPCyl
    Set ObjMaintOPCyl = Nothing
    
 ' Insert your code for output 9(Maintenance Stem Extension above Operator Cylinder)
 'This check is to ensure that PlaceCylider wont fail, if StemExtension is ZERO.
    If parStemExtension <= 0 Then parStemExtension = MIN_LENGTH
    oStartPoint.Set 0, parOperatorHeight, 0
    oEndPoint.Set 0, oStartPoint.y + parStemExtension, 0
    Set ObjMaintStemExtn = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStemExtn
    Set ObjMaintStemExtn = Nothing
    
 ' Insert your code for output 10(Maintenance Fixed portion of Atuator Shaft)
    oStartPoint.Set 0, parOperatorHeight - parOperatorDepth / 2, -parActuatorOffset
    oEndPoint.Set parArmLength, oStartPoint.y, oStartPoint.z
                        
   Set ObjMaintFixedActShaft = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintFixedActShaft
    Set ObjMaintFixedActShaft = Nothing
    
 ' Insert your code for output 11(Maintenance Variable portion of Atuator Shaft)
 'This check is to ensure that PlaceCylider wont fail, if parArmLength1 is given as ZERO.
    If parArmLength1 <= 0 Then parArmLength1 = MIN_LENGTH
    
    oStartPoint.Set parArmLength, parOperatorHeight - parOperatorDepth / 2, -parActuatorOffset
    oEndPoint.Set parArmLength + parArmLength1, oStartPoint.y, oStartPoint.z
                        
   Set ObjMaintVarLenActShaft = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintVarLenActShaft
    Set ObjMaintVarLenActShaft = Nothing
    
 ' Insert your code for output 12(Maintenance Actuator Cylinder)
    oStartPoint.Set parArmLength + parArmLength1, parOperatorHeight - parOperatorDepth / 2, -parActuatorOffset
    oEndPoint.Set oStartPoint.x + dHandWheelTk, oStartPoint.y, oStartPoint.z
                        
    Set ObjMaintActuatorCylinder = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintActuatorCylinder
    Set ObjMaintActuatorCylinder = Nothing
      
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
