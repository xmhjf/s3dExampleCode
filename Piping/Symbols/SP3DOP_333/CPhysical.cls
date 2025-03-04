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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Moday, Jan 12 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details were taken from PDS Piping Component Data SN=OP_333
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    
    
    Dim ObjStem As Object
    Dim ObjExtnRod As Object
    Dim ObjOPCyl As Object
    Dim ObjStemExtn As Object
    Dim ObjEccentricShaft As Object
    Dim ObjActuatorCylinder As Object
    Dim ObjActuatorShaft As Object
    
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
    
' Insert your code for output 1(Stem in valve portion)
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, parOperatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 2(Extension Rod)
 'This check is to ensure that PlaceCylider wont fail, if parOperatorHeight1 is ZERO.
    If parOperatorHeight1 <= 0 Then parOperatorHeight1 = MIN_LENGTH
    
    oStartPoint.Set 0, parOperatorHeight, 0
    oEndPoint.Set 0, parOperatorHeight + parOperatorHeight1, 0
    Set ObjExtnRod = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dExtnRodDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExtnRod
    Set ObjExtnRod = Nothing
    
 ' Insert your code for output 3(Operator Cylinder)
    oStartPoint.Set 0, parOperatorHeight + parOperatorHeight1, 0
    oEndPoint.Set 0, oStartPoint.y + parOperatorDepth, 0
    Set ObjOPCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOPCyl
    Set ObjOPCyl = Nothing
    
 ' Insert your code for output 4(Stem Extension above Operator Cylinder)
 'This check is to ensure that PlaceCylider wont fail, if StemExtension is ZERO.
    If parStemExtension <= 0 Then parStemExtension = MIN_LENGTH
    oStartPoint.Set 0, parOperatorHeight + parOperatorHeight1 + parOperatorDepth, 0
    oEndPoint.Set 0, oStartPoint.y + parStemExtension, 0
    Set ObjStemExtn = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dExtnRodDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemExtn
    Set ObjStemExtn = Nothing
    
 ' Insert your code for output 5(Actuator Shaft)
    oStartPoint.Set 0, parOperatorHeight + parOperatorHeight1 + parOperatorDepth / 2, -parActuatorOffset
    oEndPoint.Set parArmLength - dHandWheelTk, oStartPoint.y, oStartPoint.z
                        
   Set ObjActuatorShaft = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorShaft
    Set ObjActuatorShaft = Nothing
    
 ' Insert your code for output 5(Actuator Cylinder)
    oStartPoint.Set oEndPoint.x, oEndPoint.y, oEndPoint.z
    oEndPoint.Set parArmLength, oStartPoint.y, oStartPoint.z
                        
    Set ObjActuatorCylinder = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCylinder
    Set ObjActuatorCylinder = Nothing
      
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
