VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          MS
'   Creation Date:  Thursday, September 08 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   08.09.2004      MS  Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjMaintCyl As Object
    Dim ObjMaintBox As Object
    
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
   

' Insert your code for output (Vertical Cylinder)
    LowerCorner.Set 0, 0, 0
    UpperCorner.Set 0, parOperatorHeight - parOperatorDepth, 0
    Set ObjMaintCyl = PlaceCylinder(m_OutputColl, LowerCorner, UpperCorner, parStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintCyl
    Set ObjMaintCyl = Nothing
 
 ' Insert your code for output (Rectangular Prism)
   LowerCorner.Set -parOperatorLength, parOperatorHeight - parOperatorDepth, parOperatorWidth / 2
   UpperCorner.Set parOperatorLength1, parOperatorHeight, -parOperatorWidth / 2
   Set ObjMaintBox = PlaceBox(m_OutputColl, LowerCorner, UpperCorner)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintBox
    Set ObjMaintBox = Nothing
    Set LowerCorner = Nothing
    Set UpperCorner = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
