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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         KKC
'   Creation Date:  Tuesday, April 10 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details were taken from Technegate knife gate valve bevel gear operator
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim iOutput As Double
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parArmLength As Double
    
    Dim dStemDia As Double
    Dim dHandleDia As Double
    Dim dHandWheelTk As Double
    
    Dim oStartPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    
    Set oStartPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parArmLength = arrayOfInputs(4)
    
    If CmpDblEqual(parArmLength, 0) Then
        parArmLength = 0.35 ' 14 in
    End If
    iOutput = 0
    
    dStemDia = parOperatorDiameter * 0.1
    '   Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(dStemDia, 0.015) Then dStemDia = 0.015
    If CmpDblGreaterthan(dStemDia, 0.05) Then dStemDia = 0.05
    
    dHandleDia = dStemDia
    dHandWheelTk = dStemDia
    
' Insert your code for output 7(Maintenance Stem)
    oStartPoint.Set 0, 0, 0
    oEndPoint.Set 0, parOperatorHeight * 0.68, 0
    Dim ObjMaintStem As Object
    
    Set ObjMaintStem = PlaceCone(m_OutputColl, oEndPoint, oStartPoint, parArmLength * 0.4, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintStem
    Set ObjMaintStem = Nothing
    
' Insert your code for output 8(Maintenance Operator)
    oStartPoint.Set 0, parOperatorHeight * 0.68, 0
    oEndPoint.Set 0, parOperatorHeight, 0
    Dim ObjMaintOperator As Object
    
    Set ObjMaintOperator = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parArmLength * 0.4, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintOperator
    Set ObjMaintOperator = Nothing
    
 ' Insert your code for output 9(Maintenance Arm)
    oStartPoint.Set 0, parOperatorHeight * 0.7, 0
    oEndPoint.Set parArmLength, oStartPoint.y, oStartPoint.z
    Dim ObjMaintArm As Object
    
    Set ObjMaintArm = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parOperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintArm
    Set ObjMaintArm = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
