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
'   Author:          MS
'   Creation Date:  Tuesday, Sept 14 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: V72, Page No. D-30 of PDS Piping Component Data Reference Guide.
'    This Symbol is created with two insualation aspect outputs using 'PlaceCylinder' function.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    
 ' Insert your code for output 3(Insulation Cylinder)
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double
    
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    InsulationDiameter1 = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then InsulationDiameter1 = flangeDiam + 2 * parInsulationThickness
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    InsulationDiameter2 = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then InsulationDiameter2 = flangeDiam + 2 * parInsulationThickness
    
    If InsulationDiameter1 > InsulationDiameter2 Then
        InsulationDiameter = InsulationDiameter1
    Else
        InsulationDiameter = InsulationDiameter2
    End If
    
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
