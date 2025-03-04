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
'   Creation Date:  Friday, Dec 21 2001
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
    Dim ObjVeriticalStem As Object
    Dim ObjRectanglePrism As Object
    
    Dim parOperatorHeight As Double
    Dim parOperatorDepth As Double
    Dim parOperatorLength As Double
    Dim parOperatorLength1 As Double
    Dim parOperatorWidth As Double
    Dim parStemDia As Double
    Dim LowerCorner As New AutoMath.DPosition
    Dim UpperCorner As New AutoMath.DPosition
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDepth = arrayOfInputs(3)
    parOperatorLength = arrayOfInputs(4)
    parOperatorLength1 = arrayOfInputs(5)
    parOperatorWidth = arrayOfInputs(6)
    
'    parStemDia = 0.05
    iOutput = 0
    
    parStemDia = (parOperatorLength + parOperatorLength1) * 0.1
   '   Check that the stem diameter is within 15mm and 50mm
   If parStemDia < 0.015 Then parStemDia = 0.015
   If parStemDia > 0.05 Then parStemDia = 0.05
   

' Insert your code for output 1(Vertical Cylinder)
    LowerCorner.Set 0, 0, 0
    UpperCorner.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjVeriticalStem = PlaceCylinder(m_OutputColl, LowerCorner, UpperCorner, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVeriticalStem
    Set ObjVeriticalStem = Nothing
 
 ' Insert your code for output 2(Rectangular Prism)
   LowerCorner.Set -parOperatorLength, parOperatorHeight - parOperatorDepth, parOperatorWidth / 2
   UpperCorner.Set parOperatorLength1, parOperatorHeight, -parOperatorWidth / 2
   Set ObjRectanglePrism = PlaceBox(m_OutputColl, LowerCorner, UpperCorner)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectanglePrism
    Set ObjRectanglePrism = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
