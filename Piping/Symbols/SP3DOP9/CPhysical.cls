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
    Dim LeverLength As Double
    Dim LeverLength1 As Double
    Dim parStemDia As Double
    Dim parHandleDia As Double
    Dim StartPos  As New AutoMath.DPosition
    Dim EndPos    As New AutoMath.DPosition
  
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    OperatorHeight = arrayOfInputs(2)
    LeverLength = arrayOfInputs(3)
    LeverLength1 = arrayOfInputs(4)
    
'    parStemDia = 0.05
'    parHandleDia = 0.05
    iOutput = 0
    
     parStemDia = LeverLength1 * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   
   parHandleDia = parStemDia
    
' Place Vertical Cylinder (Stem)
    StartPos.Set 0, 0, 0
    EndPos.Set 0, OperatorHeight, 0
    Dim objStem As IngrGeom3D.Projection3d
    Set objStem = PlaceCylinder(m_OutputColl, StartPos, EndPos, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
    Set objStem = Nothing

' Place Horizontal Cylinder
    StartPos.Set -LeverLength1, (OperatorHeight - parHandleDia / 2), 0
    EndPos.Set LeverLength, StartPos.y, 0
    Dim objHandle As IngrGeom3D.Projection3d
    Set objHandle = PlaceCylinder(m_OutputColl, StartPos, EndPos, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
    Set objHandle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
