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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:     svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is:  I8AZ Page No. 356 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols.
'  29.Mar.2007     dkl  CR 116921 Implemented part data basis to support specification of only Face to Center.
'                                 Used appropriate CmpdblXXX functions for comparing doubles.
'   2.Jul.2007     svsmylav       TR-123363: Added parHandWheelAngle input for completeness.
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
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)  'P2
'    parFace1toCenter = arrayOfInputs(3) 'P3
    parInsulationThickness = arrayOfInputs(4)
'    parHandWheelAngle = arrayOfInputs(5)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
        
' This symbol is based on the following Part data basis values,
' Angle valve, Symmetrical, specified by face-to-center - 200 and
' Angle valve, Asymmetrical, specified by face-to-center 1 and face-to-center 2 - 205 or 1.
    If lPartDataBasis = 205 Or lPartDataBasis <= 1 Then 'Create Asymmetrical valve
        parFacetoCenter = arrayOfInputs(2)  'P2
        parFace1toCenter = arrayOfInputs(3) 'P3
     ElseIf lPartDataBasis = 200 Then 'Create Symmetrical valve
        parFacetoCenter = arrayOfInputs(2)
        parFace1toCenter = parFacetoCenter
     End If
     
 ' Insert your code for output 3(Insulation Cylinder)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    
' Insert your code for output 3(Insulation Cylinder)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parFace1toCenter, 0
    Set ObjInsulCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder
    Set ObjInsulCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
