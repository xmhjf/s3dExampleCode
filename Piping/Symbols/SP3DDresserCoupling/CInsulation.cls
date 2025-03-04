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
'   Author:         KKC
'   Creation Date:  Thursday, Aug 23 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy           who                      change description
'   -----------         -----                      ------------------
'   10.SEP.2007           KKC            CR-122746  Create realistic Dresser Coupling Symbol
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object

    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parMajorBodyDiameter As Double
    Dim parNoofBolts As Long
    Dim ParPortToPort As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parMajorBodyDiameter = arrayOfInputs(3)
    parNoofBolts = arrayOfInputs(4)
    ParPortToPort = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
   
 ' Insert your code for output 2(Insulated Body)
    Dim oStPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2, 0, 0
 
    Dim parInsulationDiameter As Double
    parInsulationDiameter = parMajorBodyDiameter + 2 * parInsulationThickness

    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
