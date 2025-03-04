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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:07:29 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Feb 21 2003
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim objLightPole As Object
    Dim objLightCase As Object
    
    Dim objBox      As Object
    Dim objSnout    As Object
    
    Dim strPosition As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    Dim dblXLength As Double
    Dim dblYLength As Double
    Dim dblZLength As Double
    Dim dblZOffset As Double
    Dim dblTopDiameter As Double
    Dim dblBotDiameter As Double
    
    Dim myOri As Orientation
    Set myOri = New Orientation
    
    Dim parFixtureHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureHeight = arrayOfInputs(2)
    
    iOutput = 0
    
 ' Insert your code for output 1(Light Pole)
    strPosition = "W 0.1 N 0 U " & (parFixtureHeight + 0.2742) / 2
    strDirection = "N 0 U 90"
    dblDiameter = 0.0381
    dblLength = parFixtureHeight + 0.2742
    Set objLightPole = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightPole
    Set objLightPole = Nothing
    
 ' Insert your code for output 1(Light Support)
    strPosition = "E 0.032 N 0 U " & parFixtureHeight + 0.3521
    dblXLength = 0.0254
    dblYLength = 0.0762
    dblZLength = 0.3084
    myOri.ResetDefaultAxis
    myOri.RotationAboutY = 60
    myOri.ApplyRotations
    Set objBox = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox
    Set objBox = Nothing
    
 ' Insert your code for output 4(Light Case)
    strPosition = "E 0.15 N 0 U " & parFixtureHeight + 0.305
    strDirection = "W 270 U 60"
    dblDiameter = 0.34
    dblLength = 0.181
    
    Set objLightCase = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightCase
    Set objLightCase = Nothing

 ' Insert your code for output 5(Support Bend 1)
    strPosition = "E 0.251 N 0 U " & parFixtureHeight + 0.135
    dblZOffset = 0
    dblTopDiameter = 0.17
    dblBotDiameter = 0.255
    dblLength = 0.21
    
    myOri.ResetDefaultAxis
    myOri.RotationAboutY = 60
    myOri.ApplyRotations

    Set objSnout = CreateSnout(m_outputColl, strPosition, myOri, dblZOffset, dblTopDiameter, dblBotDiameter, dblLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objSnout
    Set objSnout = Nothing
    Set myOri = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
