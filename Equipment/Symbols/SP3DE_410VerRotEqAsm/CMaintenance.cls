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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:55 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         RUK
'   Creation Date:  Monday, Feb 19 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Cylinder at the Top of equipment
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

    Dim oPartFclt           As PartFacelets.IJDPart
    Dim iOutput             As Double
    Dim ObjMaintEquipment   As Object

    Dim parBottomBodyHeight1 As Double
    Dim parBottomBodyHeight2 As Double
    Dim parBottomBodyDiameter As Double
    Dim parShaftLength As Double
    Dim parShaftDiameter As Double
    Dim parTaperBodyHeight As Double
    Dim parTopBodyHeight As Double
    Dim parTopBodyDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomBodyHeight1 = arrayOfInputs(2)     'P1
    parBottomBodyHeight2 = arrayOfInputs(3)     'P2
    parBottomBodyDiameter = arrayOfInputs(4)    'P3
    parShaftLength = arrayOfInputs(5)           'P4
    parShaftDiameter = arrayOfInputs(6)         'P5
    parTaperBodyHeight = arrayOfInputs(7)       'P6
    parTopBodyHeight = arrayOfInputs(8)         'P7
    parTopBodyDiameter = arrayOfInputs(9)       'P8
    parBundlePullingLength = arrayOfInputs(10)  'P9
    parInsulationThickness = arrayOfInputs(11)

    iOutput = 0

    'Origin is at DP/PP1
    'Input validation
    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or _
            CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parShaftLength, LINEAR_TOLERANCE) Then GoTo ErrorLabel

    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    
    'Insert your code for output 2(MaintEquipment)
    If CmpDblGreaterthan(parBundlePullingLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        Dim stPoint   As New AutoMath.DPosition
        Dim enPoint   As New AutoMath.DPosition

        stPoint.Set 0, 0, parBundlePullingLength + parTopBodyHeight + _
                            parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1
        Set ObjMaintEquipment = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTopBodyDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintEquipment
        Set ObjMaintEquipment = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    End If

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
''
End Sub
