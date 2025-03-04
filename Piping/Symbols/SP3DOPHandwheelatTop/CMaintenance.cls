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
'   File            :  CMaintenance.cls
'   Author          :  PK
'   Creation Date   :  Friday, August 17 2007
'   Description     :  Operator With Handwheel at Top
'   Reference       :  Hydrant Valve Angle Type.doc
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   24.Aug.2007     PK      CR 125185, Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim parOperatorHeight       As Double
    Dim parOperatorDiameter     As Double
    Dim parValveFacetoCenter    As Double
    Dim iOutput                 As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parValveFacetoCenter = arrayOfInputs(4)
    
    iOutput = 0
   
    'Insert your code for output 1 for maintenance(Cone)
    Dim ObjMaintCone        As Object
    Dim oStPoint            As New Automath.DPosition
    Dim oEnPoint            As New Automath.DPosition
    Dim StemMinimumDiameter As Double
    
    StemMinimumDiameter = parOperatorDiameter * 0.1
    
    'Check that the stem diameter is within 15mm and 50mm
    If CmpDblLessThan(StemMinimumDiameter, 0.015) Then StemMinimumDiameter = 0.015
    If CmpDblGreaterthan(StemMinimumDiameter, 0.05) Then StemMinimumDiameter = 0.05
   
    oStPoint.Set 0, parOperatorHeight - parValveFacetoCenter, 0
    oEnPoint.Set 0, 0, 0
'   Assumption: used 55% of the operator diameter as cone radius to enclose operator
    Set ObjMaintCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, 0.55 * parOperatorDiameter, 0.5 * StemMinimumDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCone
    
    Set ObjMaintCone = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
