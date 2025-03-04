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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:07:59 AM"
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
    Dim objCylinder As Object
    Dim objLightMount      As Object
    Dim objLightSupport    As Object
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0
 ' Insert your code for output 1(Light Mount)
    dblXLength = 0.102
    dblYLength = 0.115
    dblZLength = 0.124
    strPosition = "E .051 N 0 U 0"
    myOri.ResetDefaultAxis
    Set objLightMount = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightMount
    Set objLightMount = Nothing
    
 ' Insert your code for output 2(Light Support)
    dblXLength = 0.228
    dblYLength = 0.0254
    dblZLength = 0.0508
    strPosition = "E 0.216 N 0 U 0"
    myOri.ResetDefaultAxis
    myOri.RotationAboutX = 90
    myOri.ApplyRotations
    Set objLightSupport = CreateBox(m_outputColl, strPosition, myOri, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightSupport
    Set objLightSupport = Nothing

 ' Insert your code for output 3(Light Case)
    strPosition = "E 0.2798 N 0 D 0.0861"
    strDirection = "N 0 U 90"
    dblDiameter = 0.3048
    dblLength = 0.1468

    Set objCylinder = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing

 ' Insert your code for output 4(light)
    strPosition = "E 0.2798 N 0 D 0.2645"
    dblZOffset = 0
    dblTopDiameter = 0.1905
    dblBotDiameter = 0.2032
    dblLength = 0.21

    myOri.ResetDefaultAxis
    myOri.RotationAboutY = 90
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
