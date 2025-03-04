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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CMaintainence.cls
'   Author:         RRK/KKC
'   Creation Date:  Monday, Apr 09 2007
'   Description: This symbol is graphical implementation of the Technegate hand wheel operator
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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

    Dim parOperatorHeight As Double
    Dim parOperatorHeight1 As Double
    Dim parArmLength As Double
    Dim parArmLength1 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorRadius As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parArmLength = arrayOfInputs(4)
    parArmLength1 = arrayOfInputs(5)
    parActuatorOffset = arrayOfInputs(6)
    parActuatorRadius = arrayOfInputs(7)

    iOutput = 0
    
 ' Insert your code for output 1(Cylinder 1)
    Dim ObjMaintCyl1 As Object
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set -parArmLength * 1.1, parOperatorHeight1, -parActuatorOffset
    oEnPoint.Set parArmLength1 * 1.1, parOperatorHeight1, -parActuatorOffset
    
    Set ObjMaintCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 3 * parActuatorRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl1
    Set ObjMaintCyl1 = Nothing
    
' Insert your code for output 2(Cone)
    Dim ObjMaintCone As Object
'Assumption
    Dim dActuatorDia As Double
    dActuatorDia = 0.5 * parArmLength1
    Dim dStemHeight As Double
    dStemHeight = 0.8 * parOperatorHeight1
    
    oStPoint.Set 0, parOperatorHeight * 1.1, 0
    oEnPoint.Set 0, 0, 0
    
    Set ObjMaintCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dActuatorDia / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCone
    Set ObjMaintCone = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
