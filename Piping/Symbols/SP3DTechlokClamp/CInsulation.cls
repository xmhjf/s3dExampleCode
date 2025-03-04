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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Jul 15 2004
'   Description:
'    Clamp symbol details are provided in Techlok.pdf.
'    The Symbol is created with 19 Outputs - one insualtion Output and other are physical outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjClampBodyIns As Object

    Dim parClampInnerDiameter As Double
    Dim parBoltCenterDistance As Double
    Dim parClampOuterDiameter As Double
    Dim parClampLength As Double
    Dim parClampWidth As Double
    Dim parLugSeparation As Double
    Dim parBoltDiameter As Double
    Dim parBoltLength As Double
    Dim parBoltLugWidth As Double
    Dim parBoltPitch As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampInnerDiameter = arrayOfInputs(2) 'A
    parBoltCenterDistance = arrayOfInputs(3) 'B
    parClampOuterDiameter = arrayOfInputs(4) 'C
    parClampLength = arrayOfInputs(5)   'D
    parClampWidth = arrayOfInputs(6)    'E
    parLugSeparation = arrayOfInputs(7) 'F
    parBoltDiameter = arrayOfInputs(8)  'H
    parBoltLength = arrayOfInputs(9)    'I
    parBoltLugWidth = arrayOfInputs(10) 'J
    parBoltPitch = arrayOfInputs(11)    'K
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0

' Insert your code for output 1(Clamp Body Insulation)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set -parBoltLugWidth / 2 - parInsulationThickness, _
                -parClampLength / 2 - parInsulationThickness, _
                -parClampOuterDiameter / 2 - parInsulationThickness

    enPoint.Set parBoltLugWidth / 2 + parInsulationThickness, _
                parClampLength / 2 + parInsulationThickness, _
                parClampOuterDiameter / 2 + parInsulationThickness

    Set ObjClampBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampBodyIns
    Set ObjClampBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
