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
'   Creation Date:  Saturday, Dec 22 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   27.Aug.2003     MS/ACM TR47240:  Operator stem diameters are made proportionate so
'					      that they donot appear for small valve sizes
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    Dim ObjStem As Object
    Dim ObjTopCone As Object
    Dim ObjBottomCone As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDepth As Double
    Dim parOperatorDiameter As Double
    Dim parOperatorDiameter2 As Double
    Dim parOperatorHeight1 As Double
    Dim parOperatorDepth1 As Double
    Dim parOperatorDiameter1 As Double
    Dim parOperatorDiameter3 As Double
    Dim parStemDia As Double
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDepth = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parOperatorDiameter2 = arrayOfInputs(5)
    parOperatorHeight1 = arrayOfInputs(6)
    parOperatorDepth1 = arrayOfInputs(7)
    parOperatorDiameter1 = arrayOfInputs(8)
    parOperatorDiameter3 = arrayOfInputs(9)
    
'    parStemDia = 0.05
    iOutput = 0
    
    parStemDia = parOperatorDiameter2 * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   
' Insert your code for output 1(Cylinder)
    startPoint.Set 0, -parOperatorHeight1 + parOperatorDepth1, 0
    endPoint.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, startPoint, endPoint, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 2(Cone)
    startPoint.Set 0, parOperatorHeight - parOperatorDepth, 0 'Base Center point
    endPoint.Set 0, parOperatorHeight, 0  'Top Center point
    Set ObjTopCone = PlaceCone(m_OutputColl, startPoint, endPoint, parOperatorDiameter2 / 2, parOperatorDiameter / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCone
    Set ObjTopCone = Nothing

' Insert your code for output 2(Cone)
    startPoint.Set 0, -(parOperatorHeight1 - parOperatorDepth1), 0    'Base Center point
    endPoint.Set 0, -parOperatorHeight1, 0 'Top Center point
    Set ObjBottomCone = PlaceCone(m_OutputColl, startPoint, endPoint, parOperatorDiameter3 / 2, parOperatorDiameter1 / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomCone
    Set ObjTopCone = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
