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
'   Creation Date:  Wednesday, Aug 25 2004
'   Description:
'    This is Style HP-70 Victaulic Coupling. Symbol details are provided in 06.12.pdf.
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

    Dim parClampOuterDiameter As Double
    Dim parClampLength As Double
    Dim parClampWidth As Double
    Dim parInsulationThickness  As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampOuterDiameter = arrayOfInputs(2)
    parClampLength = arrayOfInputs(3)
    parClampWidth = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    Dim ClampInnerDiameter As Double
    Dim BoltCenterDistance As Double
    Dim LugSeparation As Double
    Dim BoltDiameter As Double
    Dim BoltLength As Double
    Dim BoltLugWidth As Double
    
'   Assumptions: Following dimensions are taken to be proportions of the input parameters
    ClampInnerDiameter = parClampOuterDiameter / 1.2
    LugSeparation = 0.6 * parClampOuterDiameter
    BoltDiameter = 0.1 * parClampOuterDiameter
    BoltLength = LugSeparation + 4 * BoltDiameter
    BoltLugWidth = parClampWidth
    BoltCenterDistance = parClampLength - 3 * BoltDiameter
    
    iOutput = 0

' Insert your code for output 1(Clamp Body Insulation)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set -BoltLugWidth / 2 - parInsulationThickness, _
                -parClampLength / 2 - parInsulationThickness, _
                -parClampOuterDiameter / 2 - parInsulationThickness

    enPoint.Set BoltLugWidth / 2 + parInsulationThickness, _
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
