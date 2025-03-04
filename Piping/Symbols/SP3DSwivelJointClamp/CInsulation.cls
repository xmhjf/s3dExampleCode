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
'    Swivel Joint Clamp symbol details are provided in biopharm.pdf, MC=CSJ, page no. 84 (Model A12MPS of
'    Tri-clamp is represented by this symbol)
'    The Symbol is created with five Outputs - one insualtion Output and other are physical outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjClampBodyIns As Object

    Dim parClampDiameter As Double
    Dim parClampWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parClampDiameter = arrayOfInputs(2)
    parClampWidth = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

' Insert your code for output 1(Clamp Body Insulation)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dYmax As Double
    Dim dDim8 As Double
    Dim dEarLength As Double
    Dim dCylinderDimension As Double 'Used for height and diameter

    dDim8 = parClampWidth * 2

    stPoint.Set -dDim8 * Cos(PI / 3) - parInsulationThickness, _
                -parClampDiameter / 2 - parInsulationThickness, _
                -parClampDiameter / 2 - parInsulationThickness

    dEarLength = parClampDiameter * 0.625
    dCylinderDimension = 0.75 * parClampWidth
    
    dYmax = parClampWidth / 2 + dCylinderDimension * 0.75 + dDim8 * Sin(PI / 3)
    If dYmax < (parClampDiameter / 2) Then dYmax = parClampDiameter / 2

    enPoint.Set dDim8 * Cos(PI / 3) + parInsulationThickness, _
                dYmax + parInsulationThickness, _
                dEarLength + parInsulationThickness

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
