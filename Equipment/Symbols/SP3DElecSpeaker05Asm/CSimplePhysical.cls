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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:07:41 AM"
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

    ' Output variables
    Dim iOutput         As Double
    Dim ObjRailMount    As Object
    Dim ObjCase         As Object
    Dim ObjSpeaker1     As Object
    Dim ObjSpeaker2     As Object
    
    'User-defined variables
    Dim strOrigin       As String
    Dim dblDiameter     As Double
    Dim dblLength       As Double
    Dim strDirection    As String
    Dim dblXBottom      As Double
    Dim dblYBottom      As Double
    Dim dblXTop         As Double
    Dim dblYTop         As Double
    Dim dblXOffset      As Double
    Dim dblYOffset      As Double
    Dim dblHeight       As Double
    Dim oriOrientation  As Orientation
    Dim parSpeakerHeight As Double
    
'*** Inputs ***'
    Set oPartFclt = arrayOfInputs(1)
    parSpeakerHeight = arrayOfInputs(2)
    iOutput = 0

'*** Outputs ***'


' Insert your code for output 1(Hand rail mount)
    strOrigin = "E 0 N 0 U 0"
    dblDiameter = 0.035
    dblLength = 0.15
    strDirection = "E 90 U 90"
    Set ObjRailMount = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjRailMount
    Set ObjRailMount = Nothing

' Insert your code for output 2(Case Cylinder)
    strOrigin = "W 0 N 0 U 0"
    dblDiameter = 0.127
    dblLength = 0.14
    strDirection = "E 90 U 0"
    Set ObjCase = CreateCylinder(m_outputColl, strOrigin, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCase
    Set ObjCase = Nothing

' Insert your code for output 3(Speaker 1)
    strOrigin = "W .227 N 0 U 0"
    dblXBottom = 0.127
    dblYBottom = 0.101
    dblXTop = parSpeakerHeight
    dblYTop = 0.178
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.314
    Set oriOrientation = New Orientation
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = -90
    oriOrientation.RotationAboutX = 90
    oriOrientation.ApplyRotations
    Set ObjSpeaker1 = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, _
        dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSpeaker1
    Set ObjSpeaker1 = Nothing

' Insert your code for output 4(Speaker 2)
    strOrigin = "E .227 N 0 U 0"
    dblXBottom = 0.127
    dblYBottom = 0.101
    dblXTop = parSpeakerHeight
    dblYTop = 0.178
    dblXOffset = 0
    dblYOffset = 0
    dblHeight = 0.314
    oriOrientation.ResetDefaultAxis
    oriOrientation.RotationAboutY = 90
    oriOrientation.RotationAboutX = 90
    oriOrientation.ApplyRotations
    Set ObjSpeaker2 = CreatePyramid(m_outputColl, strOrigin, oriOrientation, dblXBottom, dblYBottom, dblXTop, dblYTop, dblXOffset, _
        dblYOffset, dblHeight)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSpeaker2
    Set ObjSpeaker2 = Nothing
    Set oriOrientation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
