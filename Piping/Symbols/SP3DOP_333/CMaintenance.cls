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
'   This Symbol details were taken from PDS Piping Component Data SN=OP_333.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Oct.2004     svsmylav       CR-57807 Added Maintenace Envelope to the operator
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Const MIN_LENGTH = 0.0001

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim iOutput As Double
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOperCenFromTop As Double
    Dim parOperatorDepth As Double
    Dim parArmLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parOperatorHeight1 As Double
    Dim parStemExtension As Double
    
    Dim dStemDia As Double
    Dim dExtnRodDia As Double
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
    parOperatorHeight1 = arrayOfInputs(9) ''Y11
    parStemExtension = arrayOfInputs(10) ''Y12
    
    iOutput = 0
    
    dStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If dStemDia < 0.015 Then dStemDia = 0.015
    If dStemDia > 0.05 Then dStemDia = 0.05
    ' Assumed that both Extension Rod and Stem have the same Diameter
    dExtnRodDia = dStemDia
    
    dHandleDia = dStemDia
    If dHandleDia >= parActuatorDiameter / 2 Then dHandleDia = parActuatorDiameter / 2
    
    dHandWheelTk = dStemDia
    
' Insert your code for output 7(Maintenance Stem)
 'This check is to ensure that parOperatorHeight1 is not ZERO.
    If parOperatorHeight1 <= 0 Then parOperatorHeight1 = MIN_LENGTH

    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, parOperatorHeight + parOperatorHeight1, 0
    oStartPoint.z = oStartPoint.z + 0
    Dim ObjMaintStem As Object
    Set ObjMaintStem = PlaceCone(m_OutputColl, oEndPoint, oStartPoint, parOperatorDiameter, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
' Insert your code for output 8(Maintenance Operator)
    oStartPoint.Set 0, parOperatorHeight + parOperatorHeight1, 0
    oEndPoint.Set 0, parOperatorHeight + parOperatorHeight1 + _
                        parOperatorDepth + parStemExtension, 0
    Dim ObjMaintOperator As Object
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
 ' Insert your code for output 9(Maintenance Actuator)
    oStartPoint.Set 0, parOperatorHeight + parOperatorHeight1 + parOperatorDepth / 2, -parActuatorOffset
    oEndPoint.Set parArmLength, oStartPoint.y, oStartPoint.z
    Dim ObjMaintActuator As Object
   Set ObjMaintActuator = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintActuator
    Set ObjMaintActuator = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
