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

'   ProgID          :  SP3DOPSpurGear.OPSpurGear
'   File            :  CMaintenance.cls
'   Author          :  PK
'   Creation Date   :  Friday, Sept 10 2007
'   Description     :  Spur Gear type operator to be used with 3 way diverter combination valve
'                      of Tyco Flow Control
'   Reference       :  http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   10.Sep.2007     PK          CR-126718:Created the symbol.
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
    Dim parArmLength        As Double
    Dim parArmLength1       As Double
    Dim parPinionShaftDia   As Double
        
    Dim iOutput    As Integer
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim oCentPos As AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dFlangeCenterHeight As Double
    dFlangeCenterHeight = 0.25 * parOperatorHeight
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parArmLength = arrayOfInputs(3)
    parArmLength1 = arrayOfInputs(4)
    parPinionShaftDia = arrayOfInputs(5)

    iOutput = 0
    oCentPos.Set 0, 0, 0
            
    'Insert your code for output 1 for maintenance(Cylinder)
    Dim oMaintCylinder  As Object
    oStPoint.Set oCentPos.x, oCentPos.y + dFlangeCenterHeight, oCentPos.z
    oEnPoint.Set oCentPos.x, oCentPos.y + dFlangeCenterHeight + 1.5 * parOperatorHeight, oCentPos.z
       
    Set oMaintCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2.1 * parArmLength, True)
            
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oMaintCylinder
       
    Set oMaintCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub

