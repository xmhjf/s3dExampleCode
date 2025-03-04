VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:11 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Feb 27 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
        
    'Output variable declarations
    Dim iOutput         As Double
    Dim ObjStand1       As Object
    Dim ObjStand2       As Object
    Dim ObjIntercom     As Object
    
    'User-defined variable declarations
    Dim strOrigin       As String
    Dim dblXLength      As Double
    Dim dblYLength      As Double
    Dim dblZLength      As Double
    Dim oriOrientation  As Orientation

    Set oriOrientation = New Orientation
    Dim parSpeakerHeight As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSpeakerHeight = arrayOfInputs(2)
    
    iOutput = 0

' Insert your code for output 1(Stand 1)
    strOrigin = "E 0 N 0.093 U 0.7465"
    dblXLength = 0.0381
    dblYLength = 0.0381
    dblZLength = 1.493
    
    Set ObjStand1 = CreateBox(m_outputColl, strOrigin, Nothing, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjStand1
    Set ObjStand1 = Nothing
    
' Insert your code for output 2(Stand 2)
    strOrigin = "E 0 S 0.093 U 0.7465"
    dblXLength = 0.0381
    dblYLength = 0.0381
    dblZLength = 1.493
   
    Set ObjStand2 = CreateBox(m_outputColl, strOrigin, Nothing, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjStand2
    Set ObjStand2 = Nothing

' Insert your code for output 3(Intercom box)

    strOrigin = "E 0.15255 N 0 U 1.255"
    dblXLength = 0.267
    dblYLength = 0.276
    dblZLength = parSpeakerHeight
    Set ObjIntercom = CreateBox(m_outputColl, strOrigin, Nothing, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjIntercom
    Set ObjIntercom = Nothing
    Set oriOrientation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
