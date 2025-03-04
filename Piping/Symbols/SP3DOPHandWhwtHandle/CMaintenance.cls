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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          KKC
'   Creation Date:  Tuesday, Jun 10 2008
'   Description:
'    The symbol is prepared based on 120 deg angle valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-144044  Provide a 120 deg angle valve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double

    Dim parOperatorHeight As Double
    Dim dStemMinimumDiameter As Double
    Dim parOperatorDiameter As Double
    Dim dStemExtension As Double
    Dim parLeverLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parLeverLength = arrayOfInputs(4)

'   Initialize
    iOutput = 0
    dStemExtension = parOperatorHeight * 0.2
' Insert your code for output 1(Cone)
    Dim ObjMaintCone As Object
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    oStPoint.Set 0, parOperatorHeight + dStemExtension, 0
    oEnPoint.Set 0, 0, 0
    Set ObjMaintCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parLeverLength, 0, True)

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
