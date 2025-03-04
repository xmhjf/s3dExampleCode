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
'   Creation Date:  Wednesday, Dec 19 2001
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
    Dim ObjVerCylin As IngrGeom3D.Projection3d
    Dim ObjHoriCylin As IngrGeom3D.Projection3d
    Dim ObjIncliCylin As IngrGeom3D.Projection3d
    Dim parOperatorHeight1 As Double
    Dim parOperatorHeight As Double
    Dim parLeverLength As Double
    Dim parStemDia As Double
    Dim parHandleDia As Double
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorHeight1 = arrayOfInputs(3)
    parLeverLength = arrayOfInputs(4)
    
'    parStemDia = 0.05
'    parHandleDia = 0.05
    iOutput = 0
    parStemDia = parLeverLength * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   
   parHandleDia = parStemDia
    
'Place Vertical Cylinder
    StartPos.Set 0, 0, 0
    EndPos.Set 0, parOperatorHeight1, 0
    Set ObjVerCylin = PlaceCylinder(m_OutputColl, StartPos, EndPos, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin
    Set ObjVerCylin = Nothing
    
'Place Horizontal Cylinder
    StartPos.Set 0, parOperatorHeight1 - parHandleDia / 2, 0
    EndPos.Set parLeverLength, StartPos.y, 0
    Set ObjHoriCylin = PlaceCylinder(m_OutputColl, StartPos, EndPos, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylin
    Set ObjHoriCylin = Nothing
    
'Place Inclined Cylinder
    StartPos.Set 0, parOperatorHeight1 - parHandleDia / 2, 0
    EndPos.x = -(parOperatorHeight - parOperatorHeight1) ' Since Tan(45) is Equal to 1
    EndPos.y = StartPos.y + (parOperatorHeight - parOperatorHeight1)
    EndPos.z = 0
    Set ObjIncliCylin = PlaceCylinder(m_OutputColl, StartPos, EndPos, parHandleDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIncliCylin
    Set ObjIncliCylin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
