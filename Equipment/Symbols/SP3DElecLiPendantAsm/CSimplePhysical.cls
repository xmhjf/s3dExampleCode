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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:07:49 AM"
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
    Dim objLightMount As Object
    Dim objLightSupport As Object
    Dim objSnout    As Object
    Dim ObjConduitPort As Object
    
    Dim strPosition As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    Dim dblZOffset As Double
    Dim dblTopDiameter As Double
    Dim dblBotDiameter As Double
    
    Dim myOri As Orientation
    Set myOri = New Orientation
    
    Dim parFixtureDrop As Double
    Dim parFixtureDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFixtureDrop = arrayOfInputs(2)
    parFixtureDiameter = arrayOfInputs(3)

    iOutput = 0
 
 ' Insert your code for output 1(objLight Mount)
    strPosition = "E 0 N 0 U " & (parFixtureDrop / 2)
    strDirection = "N 0 U 90"
    dblDiameter = 0.0254
    dblLength = parFixtureDrop

    Set objLightMount = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightMount
    Set objLightMount = Nothing

' Insert your code for output 2(Light Support)
    strPosition = "E 0 N 0 U " & (parFixtureDiameter * 0.45) / 2 + (parFixtureDiameter * 0.7)
    strDirection = "N 0 U 90"
    dblDiameter = parFixtureDiameter
    dblLength = parFixtureDiameter * 0.45

    Set objLightSupport = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objLightSupport
    Set objLightSupport = Nothing

 ' Insert your code for output 3(light)
    strPosition = "E 0 N 0 U " & (parFixtureDiameter * 0.7) / 2
    dblZOffset = 0
    dblTopDiameter = parFixtureDiameter * 0.75
    dblBotDiameter = parFixtureDiameter * 0.5
    dblLength = parFixtureDiameter * 0.7

    myOri.ResetDefaultAxis
    myOri.RotationAboutY = -90
    myOri.ApplyRotations

    Set objSnout = CreateSnout(m_outputColl, strPosition, myOri, dblZOffset, dblTopDiameter, dblBotDiameter, dblLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objSnout
    Set objSnout = Nothing
    Set myOri = Nothing
    
' Insert your code for output 4(Conduit Port)
    Dim pConduitPort As AutoMath.DPosition
    Set pConduitPort = New DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New DVector
    
    pConduitPort.Set 0, 0, parFixtureDrop
    oDir.Set 0, 0, 1
    Set ObjConduitPort = CreateConduitNozzlePH(pConduitPort, oDir, m_outputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set oDir = Nothing
    Set pConduitPort = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
