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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Jun 6 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No.: S78AC, Page No. 108 of PDS Piping Component Data Reference Guide.
' This Symbol is created with six insualation aspect outputs using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjInsUpTrBody As Object
    Dim ObjInsLowTrBody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parTrapDiameter As Double
    Dim parOffset2 As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset1 = arrayOfInputs(4)
    parTrapDiameter = arrayOfInputs(5)
    parOffset2 = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Upper Trap Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    stPoint.Set -0.5 * parFace1toCenter, 0, 0
    enPoint.Set -0.5 * parFace1toCenter + (parOffset2 + parInsulationThickness) * Cos(PI / 4), 0, _
                                         (parOffset2 + parInsulationThickness) * Sin(PI / 4)
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsUpTrBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsUpTrBody
    Set ObjInsUpTrBody = Nothing
    
' Insert your code for output 2(Insulation for Lower Trap Body)
    stPoint.Set 0, 0, -parOffset1 - parInsulationThickness
    enPoint.Set 0, 0, 0
    parInsulationDiameter = parTrapDiameter + 2 * parInsulationThickness
    Set ObjInsLowTrBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsLowTrBody
    Set ObjInsLowTrBody = Nothing
    
' Insert your code for output 3(Insulation Port1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 4(Insulation for body Port1 Side)
    stPoint.Set -parFace1toCenter + flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 5(Insulation Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 6(Insulation for body Port2 Side)
    stPoint.Set parFace2toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
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
