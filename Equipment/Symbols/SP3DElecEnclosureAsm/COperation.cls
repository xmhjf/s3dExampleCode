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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          NN
'   Creation Date:  Tuesday, Mar 12 2002
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
    
    Dim iOutput     As Double
    Dim ObjEnclosureOp As Object
    
    Dim parBasicAsmHeight As Double
    Dim parBasicAsmWidth As Double
    Dim parBasicAsmLength As Double
    
    Dim CenterPos       As IJDPosition
    Dim pEnclosureOpU As IJDPosition
    Dim pEnclosureOpL As IJDPosition
    Set pEnclosureOpU = New DPosition
    Set pEnclosureOpL = New DPosition
    Set CenterPos = New DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasicAsmHeight = arrayOfInputs(2)
    parBasicAsmWidth = arrayOfInputs(3)
    parBasicAsmLength = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 7(Enclosure Operation)
    CenterPos.Set 0, 0, 0
    pEnclosureOpL.Set CenterPos.x + parBasicAsmLength, CenterPos.y, CenterPos.z
    pEnclosureOpU.Set CenterPos.x + parBasicAsmLength + parBasicAsmHeight, CenterPos.y + parBasicAsmWidth, CenterPos.z + parBasicAsmHeight
    Set ObjEnclosureOp = PlaceBox(m_OutputColl, pEnclosureOpL, pEnclosureOpU)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEnclosureOp
    Set ObjEnclosureOp = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
