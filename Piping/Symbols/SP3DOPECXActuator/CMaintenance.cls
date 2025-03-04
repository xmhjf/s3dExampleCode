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

'   ProgID          :  SP3DOPECXActuator.OPECXActuator
'   File            :  CMaintenance.cls
'   Author          :  PK
'   Creation Date   :  Monday, 1 Oct 2007
'   Description     :  ECX Type Operator
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   01-OCT-2007      PK        CR-127652:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim parOperatorHeight    As Double
    Dim parOperatorDiameter  As Double
    Dim parWidth             As Double
    Dim parPositionerOffset  As Double
    Dim dActOperatorHeight   As Double
    Dim iOutput              As Integer
    Dim oStPoint             As AutoMath.DPosition
    Dim oEnPoint             As AutoMath.DPosition
    Dim oCentPos             As AutoMath.DPosition
  
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parWidth = arrayOfInputs(4)
    parPositionerOffset = arrayOfInputs(5)
        
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
       
    iOutput = 0
    oCentPos.Set 0, parPositionerOffset, 0
    dActOperatorHeight = parOperatorHeight - parPositionerOffset
    
    'Insert your code for Maintenance Cylinder
    Dim oMaintCylinder  As Object
    Dim dBiggerValue    As Double
    
    If CmpDblGreaterthanOrEqualTo(parWidth, parOperatorDiameter) Then
        dBiggerValue = parWidth
    Else
        dBiggerValue = parOperatorDiameter
    End If
    
    oStPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + 1.2 * dActOperatorHeight, oCentPos.z

    Set oMaintCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * dBiggerValue, True)

    'Set Output(Maintenance Cylinder
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintCylinder
    Set oMaintCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
