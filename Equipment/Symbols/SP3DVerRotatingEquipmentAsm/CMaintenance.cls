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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:43 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using one maintenance aspect output which use 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   06.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  26.Sep.2006     RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
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
    Dim ObjMaintEquipment As Object

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

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomBodyHeight1 = arrayOfInputs(2)
    parBottomBodyHeight2 = arrayOfInputs(3)
    parBottomBodyDiameter = arrayOfInputs(4)
    parShaftLength = arrayOfInputs(5)
    parShaftDiameter = arrayOfInputs(6)
    parTaperBodyHeight = arrayOfInputs(7)
    parTopBodyHeight = arrayOfInputs(8)
    parTopBodyDiameter = arrayOfInputs(9)
    parBundlePullingLength = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0

    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel

' Insert your code for output 5(MaintEquipment)
     If CmpDblGreaterthan(parBundlePullingLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        Dim stPoint   As New AutoMath.DPosition
        Dim enPoint   As New AutoMath.DPosition
        
        stPoint.Set 0, 0, parBundlePullingLength + parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1
        Set ObjMaintEquipment = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTopBodyDiameter, True)
        
'       Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintEquipment
        Set ObjMaintEquipment = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
