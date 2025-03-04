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
'   Physical.cls
'   ProgID:         SP3DExternalSpringOP.ExtlSpringOP
'   Author:         KKC
'   Creation Date:  Saturday, 18 May 2007
'   Description:
'   This symbol is taken form the Prince Figure Operator 813(http://www.tycoflowcontrol-na.com/ld/PRCMC-0033-US.pdf)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
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
    Dim ObjLever As Object
    Dim ObjConnectingLever As Object
    Dim ObjConnectingLever1 As Object
    Dim ObjCounterWeight As Object
    
    Dim parOperatorLength As Double
    Dim parOperatorLength1 As Double
    Dim parOperatorWidth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorLength = arrayOfInputs(2)
    parOperatorLength1 = arrayOfInputs(3)
    parOperatorWidth = arrayOfInputs(4)
    
    iOutput = 0
' Insert your code for output 1(Operator Lever Connection 1)
    Dim StPoint As AutoMath.DPosition
    Dim EnPoint As AutoMath.DPosition

    Dim dOperatorDiameter As Double
    dOperatorDiameter = 0.15 * parOperatorWidth

    Set StPoint = New AutoMath.DPosition
    Set EnPoint = New AutoMath.DPosition
    
    StPoint.Set 0, 0, 0
    EnPoint.Set 0, parOperatorLength, 0
    Set ObjConnectingLever = PlaceCylinder(m_OutputColl, StPoint, EnPoint, dOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnectingLever
    Set ObjConnectingLever = Nothing

 ' Insert your code for output 2(Operator Lever Connection 2)

    StPoint.Set 0, parOperatorLength - dOperatorDiameter * 0.6, 0
    EnPoint.Set 0, parOperatorLength - dOperatorDiameter * 0.6, parOperatorWidth - dOperatorDiameter / 2

    Set ObjConnectingLever1 = PlaceCylinder(m_OutputColl, StPoint, EnPoint, 1.2 * dOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnectingLever1
    Set ObjConnectingLever1 = Nothing

 ' Insert your code for output 3(Operator Lever)

    StPoint.Set 0, 1.05 * parOperatorLength, parOperatorWidth - dOperatorDiameter / 2
    EnPoint.Set 0, 1.05 * parOperatorLength - 0.8 * parOperatorLength1, parOperatorWidth - dOperatorDiameter / 2

    Set ObjLever = PlaceCylinder(m_OutputColl, StPoint, EnPoint, 1.3 * dOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLever
    Set ObjLever = Nothing
      
'  Insert your code for output 4(Counter Weight)
    StPoint.Set 0, 1.05 * parOperatorLength - 0.8 * parOperatorLength1, parOperatorWidth - dOperatorDiameter / 2
    EnPoint.Set 0, 1.05 * parOperatorLength - parOperatorLength1, parOperatorWidth - dOperatorDiameter / 2

    Set ObjCounterWeight = PlaceCylinder(m_OutputColl, StPoint, EnPoint, 1.8 * dOperatorDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCounterWeight
    Set ObjCounterWeight = Nothing
    
    Set StPoint = Nothing
    Set EnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
