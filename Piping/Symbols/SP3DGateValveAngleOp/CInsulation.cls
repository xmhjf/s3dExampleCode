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
'   Author:         RRK
'   Creation Date:  Wednusday,21 February
'   Description:    Insulated Valve, construct box FacetoFace x (Insulation + Valve Height)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
''  21.Feb.2007      RRK                     TR-113129 Added Insulation aspect
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjInsulatedBox As Object
    Dim parFacetoFace As Double
    Dim parValveHeight   As Double
    Dim parInsulationThickness As Double
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim dValveBottomtoCL As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parValveHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0
'   Distance from Valve Bottom to its flow centerline
    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                    pipeDiam, flangeThick, flangeDiam, _
                    sptOffset, depth
    If CmpDblEqual(flangeDiam, 0) Then
        flangeDiam = pipeDiam * 2
    End If
    dValveBottomtoCL = flangeDiam * 0.6
    
    
 ' Insert your code for output 14(InsulatedBox)
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -parFacetoFace / 2, -(dValveBottomtoCL + parInsulationThickness), _
                                        (dValveBottomtoCL + parInsulationThickness)
    oEnPoint.Set parFacetoFace / 2, (parValveHeight - dValveBottomtoCL + parInsulationThickness), _
                                    -(dValveBottomtoCL + parInsulationThickness)
    
    Set ObjInsulatedBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBox
    Set ObjInsulatedBox = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

