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
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I253AZ Page No. 367 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
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
    
    Dim iOutput             As Double
    Dim ObjInsFilterBody   As Object
    Dim ObjInsPort1         As Object
    Dim ObjInsBody1         As Object
    Dim ObjInsPort2         As Object
    Dim ObjInsBody2         As Object
    
    Dim parNozz1toCentre As Double
    Dim parNozz2toCentre As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parTrapBodyDiameter As Double
    Dim parTrapHeightBottoCen As Double
    Dim parTrapHeightToptoCen As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozz1toCentre = arrayOfInputs(2)  'P2
    parNozz2toCentre = arrayOfInputs(3)  'P3
    parOffset1 = arrayOfInputs(4)       'P4
    parOffset2 = arrayOfInputs(5)       'P5
    parTrapBodyDiameter = arrayOfInputs(6)      'P6
    parTrapHeightBottoCen = arrayOfInputs(7)     'P7
    parTrapHeightToptoCen = arrayOfInputs(8)     'P8
    parInsulationThickness = arrayOfInputs(9)
    
       iOutput = 0

' Insert your code for output 1(Insulation for Trap  Body)
    stPoint.Set 0, 0, -parTrapHeightBottoCen - parInsulationThickness
    enPoint.Set 0, 0, parTrapHeightToptoCen + parInsulationThickness
    parInsulationDiameter = parTrapBodyDiameter + 2 * parInsulationThickness
    Set ObjInsFilterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFilterBody
    Set ObjInsFilterBody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    stPoint.Set -parNozz1toCentre, 0, -parOffset1
    enPoint.Set -parNozz1toCentre + flangeThick + parInsulationThickness, 0, -parOffset1
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for pipe Port1 Side)
    stPoint.Set -parNozz1toCentre + flangeThick, 0, -parOffset1
    enPoint.Set 0, 0, -parOffset1
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 4(Insulation Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -parNozz2toCentre, 0, parOffset2
    enPoint.Set -parNozz2toCentre + flangeThick + parInsulationThickness, 0, parOffset2
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 5(Insulation for pipe Port2 Side)
    stPoint.Set -parNozz2toCentre + flangeThick, 0, parOffset2
    enPoint.Set 0, 0, parOffset2
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
