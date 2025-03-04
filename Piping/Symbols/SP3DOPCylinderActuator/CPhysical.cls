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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK/KKC
'   Creation Date:  Monday, Apr 09 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
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

    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0

 ' Insert your code for output 1(Spindle)
    Dim objStem As IngrGeom3D.Projection3d
    Dim dStemHeight As Double
    Dim dStemDia As Double
    
'Assumptions
    dStemDia = 0.05 * parArmLength1
    dStemHeight = 0.8 * parOperatorHeight1
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, dStemHeight, 0
    
    Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
    Set objStem = Nothing
    
 ' Insert your code for output 2(Actuator)
    Dim objActuator As IngrGeom3D.Projection3d
    Dim dActuatorDia As Double

'Assumption
    dActuatorDia = 0.5 * parArmLength1
    
    oStPoint.Set 0, dStemHeight, 0
    oEnPoint.Set 0, parOperatorHeight, 0
    
    Set objActuator = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dActuatorDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objActuator
    Set objActuator = Nothing
    
 ' Insert your code for output 3(Cylinder 1)
    Dim objCylinder1 As IngrGeom3D.Projection3d
    Dim dCylinder1Dia As Double
    Dim dCylinder1EndPos As Double

'Assumptions
    dCylinder1Dia = parActuatorRadius
    dCylinder1EndPos = 0.35 * parArmLength1
    
    oStPoint.Set -parArmLength, parOperatorHeight1, -parActuatorOffset
    oEnPoint.Set dCylinder1EndPos, parOperatorHeight1, -parActuatorOffset
    
    Set objCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinder1Dia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder1
    Set objCylinder1 = Nothing
    
 ' Insert your code for output 4(Cylinder 2)
    Dim objCylinder2 As IngrGeom3D.Projection3d
    
    oStPoint.Set dCylinder1EndPos, parOperatorHeight1, -parActuatorOffset
    oEnPoint.Set parArmLength1, parOperatorHeight1, -parActuatorOffset
    
    Set objCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * parActuatorRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder2
    Set objCylinder2 = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
