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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Thursday, Nov 23 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   08.SEP.2006     KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG                   TR-124959 Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim ObjInsulatedBody As Object
    
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(2)
    
    iOutput = 0

 ' Insert your code for output 1(Insulated Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim endofInsulation       As New AutoMath.DPosition
    If CmpDblEqual(flangeThick, 0) Then
        flangeThick = NEGLIGIBLE_THICKNESS
    End If
    endofInsulation.Set flangeThick + parInsulationThickness, 0, 0
    Dim parInsulationDiameter As Double
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, CenterPos, endofInsulation, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
