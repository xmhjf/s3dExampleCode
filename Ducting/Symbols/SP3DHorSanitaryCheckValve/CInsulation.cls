VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          VRG
'   Creation Date:  Wednesday, Sep 19 2007
'   Description:
'  This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Integer
    Dim parFacetoFace As Double
    Dim parBodyDiameter As Double
    Dim parPortVOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parBodyDiameter = arrayOfInputs(3)
    parPortVOffset = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    iOutput = 0
    
 ' Insert your code for InsulatedBody(Output 4)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oInsulatedBody As Object
       
    oStPoint.Set -parFacetoFace / 2 - parInsulationThickness, -0.1 * parBodyDiameter, 0
    oEnPoint.Set parFacetoFace / 2 + parInsulationThickness, -0.1 * parBodyDiameter, 0
    Set oInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBodyDiameter + 2 * parPortVOffset, True)
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "InsulatedBody", oInsulatedBody
    Set oInsulatedBody = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
