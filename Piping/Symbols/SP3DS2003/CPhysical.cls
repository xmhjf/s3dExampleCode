VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Tuesday, Dec 18 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'					      that they donot appear for small valve sizes
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim iOutput     As Double
    Dim OperatorHeight As Double
    Dim OperatorHeight1 As Double
    Dim OperatorDiameter1 As Double
    Dim OperatorDiameter As Double
    Dim parStemDia1 As Double
    Dim parDiscTk As Double
    Dim CenterPos As New AutoMath.DPosition
    Dim objThinStem As IngrGeom3D.Projection3d
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    Dim objThickStem As IngrGeom3D.Projection3d
    Dim objDisc As IngrGeom3D.Projection3d
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    OperatorHeight = arrayOfInputs(2) ' Ht of thinner stem
    OperatorHeight1 = arrayOfInputs(3) ' Ht of Thicker Stem
    OperatorDiameter1 = arrayOfInputs(4) ' Dia of Disc
    OperatorDiameter = arrayOfInputs(5) 'Dia of Thicker Stem
    
    iOutput = 0
    CenterPos.Set 0, 0, 0
'    parStemDia1 = 0.05
'    parDiscTk = 0.05
    
    parStemDia1 = OperatorDiameter * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia1 < 0.015 Then parStemDia1 = 0.015
   If parStemDia1 > 0.05 Then parStemDia1 = 0.05
   If parStemDia1 > OperatorDiameter / 2 Then parStemDia1 = OperatorDiameter / 2
   
   parDiscTk = parStemDia1

' Place Thinner Stem
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z
    EndPos.Set CenterPos.x, CenterPos.y + OperatorHeight, CenterPos.z
    Set objThinStem = PlaceCylinder(m_OutputColl, StartPos, EndPos, parStemDia1, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objThinStem
    Set objThinStem = Nothing
    
' Place Thicker Stem
    StartPos.Set CenterPos.x, CenterPos.y + OperatorHeight, CenterPos.z
    EndPos.Set CenterPos.x, CenterPos.y + OperatorHeight1 + OperatorHeight - parDiscTk, CenterPos.z
    Set objThickStem = PlaceCylinder(m_OutputColl, StartPos, EndPos, OperatorDiameter, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objThickStem
    Set objThickStem = Nothing

' Palce Disc over the Stem
    StartPos.Set CenterPos.x, CenterPos.y + OperatorHeight1 + OperatorHeight - parDiscTk, CenterPos.z
    EndPos.Set CenterPos.x, CenterPos.y + OperatorHeight1 + OperatorHeight, CenterPos.z
    Set objDisc = PlaceCylinder(m_OutputColl, StartPos, EndPos, OperatorDiameter1, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDisc
    Set objDisc = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
