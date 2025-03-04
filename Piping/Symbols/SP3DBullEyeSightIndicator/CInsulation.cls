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
'   Author:          SSP
'   Creation Date:  Friday, Nov 15 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Bull Eye Sight Indicator.Refering www.resistoflex.com
'   The symbol consists of 18 outputs, 16 physical and 2 insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------     -----                  ------------------
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is added.
'   08.SEP.2006     KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG                    TR-124959 Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
 '''
End Sub
Private Function ReturnMax4(A As Double, B As Double, C As Double, D As Double) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
        ReturnMax4 = MaxValue
End Function


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Integer
    Dim ObjInsulationBody1 As Object
    Dim ObjInsulationBody2 As Object
    Dim parFacetoCenter As Double
    Dim parIndicatorHeight As Double
    Dim parInsulationThickness As Double
   
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parIndicatorHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
        
    iOutput = 0
    
 ' Insert your code for output 15(Insulation Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim InsulationDiameter As Double
    Dim maxDiameter As Double
    maxDiameter = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    InsulationDiameter = maxDiameter + 2 * parInsulationThickness
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set parFacetoCenter, 0, 0
    
    Set ObjInsulationBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody1
    Set ObjInsulationBody1 = Nothing
    
 ' Insert your code for output 16(Chamber Body Insulation)
    If CmpDblLessThanOrEqualTo(flangeDiam1, pipeDiam1) Then
        flangeDiam1 = 1.2 * pipeDiam1
    End If
    InsulationDiameter = flangeDiam1 + 2 * parInsulationThickness
    stPoint.Set 0, parIndicatorHeight, 0
    enPoint.Set 0, -parIndicatorHeight, 0
    
    Set ObjInsulationBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody2
    Set ObjInsulationBody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
