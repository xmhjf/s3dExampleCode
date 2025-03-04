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
'   Creation Date   :  Friday, Sept 10 2007
'   Description     :  Wrench type operator to be used with 3 way diverter combination valve
'                      of Tyco Flow Control
'   Reference       :  http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   24.Aug.2007     PK          CR-126718:Created the symbol.
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
    
    Dim parOperatorHeight   As Double
    Dim parLeverLength  As Double

    
    Dim iOutput    As Integer
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim oCentPos As AutoMath.DPosition
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parLeverLength = arrayOfInputs(3)
    
    Set oCentPos = New DPosition
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    iOutput = 0
    oCentPos.Set 0, 0, 0
    
    Dim dFlangeCenterHeight As Double
    dFlangeCenterHeight = 0.25 * parOperatorHeight
    
    'Insert your code for output 1 for maintenance(Cylinder)
    Dim oMaintCylinder  As Object
    oStPoint.Set oCentPos.x, oCentPos.y + dFlangeCenterHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + 1.5 * parOperatorHeight, oCentPos.z
       
    Set oMaintCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * parLeverLength, True)
            
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintCylinder
       
    Set oMaintCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCentPos = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
