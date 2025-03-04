VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          JF
'   Creation Date:  Thursday, Mar 21 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjSupportSkirt As Object
    Dim ObjTank As Object
    Dim ObjTopDome As Object
    Dim ObjBottomDome As Object
    Dim ObjOperationEnv As Object
    Dim ObjMaintenanceEnv As Object
    Dim ObjInsulationEnv As Object
    
    Dim parSkirtHeight As Double
    Dim parSkirtDiameter As Double
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parDomeHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSkirtHeight = arrayOfInputs(2)
    parSkirtDiameter = arrayOfInputs(3)
    parVesselDiameter = arrayOfInputs(4)
    parVesselHeight = arrayOfInputs(5)
    parDomeHeight = arrayOfInputs(6)
  
    
    iOutput = 0

 ' Insert your code for output 5()

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parOperatedVesselDiameter As Double
    parOperatedVesselDiameter = (parVesselDiameter * 1.125)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, (parSkirtHeight + parVesselHeight + parDomeHeight) * 1.5
    Set ObjOperationEnv = PlaceCylinder(m_OutputColl, stPoint, enPoint, parOperatedVesselDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperationEnv
    Set ObjOperationEnv = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
