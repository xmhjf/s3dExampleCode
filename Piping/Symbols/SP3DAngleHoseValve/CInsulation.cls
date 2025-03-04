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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jun 14 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model: V71, Page No. D-31 of PDS Piping Component Data Reference Guide.
' This Symbol is created with two insualation aspect outputs using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who                               change description
'   -----------         -----                                ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'   02.Jan.2006     KKK              CR-91106  Updated the symbol to be based on either face-to-face
'                                                     or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
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
    Dim ObjInsulCylinder As Object
    
    Dim parFacetoCenter As Double
    Dim parHoseConnEndtoCenter As Double
    Dim parInsulationThickness As Double
    Dim parDimension1 As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parDimension1 = arrayOfInputs(2)
    parHoseConnEndtoCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    
    iOutput = 0
    
 ' Insert your code for output 3(Insulation Cylinder)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis = 5) Then 'Applicable for 5 "Face-To-Face dimension basis"
    parFacetoCenter = parDimension1 / 2
    ElseIf (lPartDataBasis = 10) Or (lPartDataBasis <= 1) Then 'Applicable for 1 "Default" & 10 "Face-to-Center dimension basis"
    parFacetoCenter = parDimension1
    End If
    
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    
' Insert your code for output 3(Insulation Cylinder)
    stPoint.Set 0, parHoseConnEndtoCenter, 0
    enPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
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
