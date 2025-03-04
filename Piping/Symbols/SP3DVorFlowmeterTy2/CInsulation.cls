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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav 
'   Creation Date:  Monday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM2,
'   PDMS Category Name=IFND). Physical aspect consist of one horizontal cylinder, two vertical cylinders
'   and two nozzles one facing -nagative an the other in positive direction of X-axis.
'   Insulation aspect consist of two cylinders one along X-axis, and teh other along Y-axis.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Private Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim depth           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput     As Double
    Dim ObjInsHorCylinder As Object
    Dim ObjInsVerCylinder As Object
    Dim ObjHorCylinder As Object
    Dim ObjConnector As Object
    Dim ObjVerCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)      'D2
    parFace2toCenter = arrayOfInputs(3)      'D3
    parFlowDiameter = arrayOfInputs(4)       'D4
    parInstrumentHeight = arrayOfInputs(5)   'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentDiameter = arrayOfInputs(7) 'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Horizontal Cylinder)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

    stPoint.Set -parFace1toCenter - parInsulationThickness, 0, 0
    enPoint.Set parFace2toCenter + parInsulationThickness, 0, 0
'   Use maximum value among pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2 and parFlowDiameter
'   for computing insulation diameter
    parInsulationDiameter = ReturnMax5(pipeDiam1, flangeDiam1, _
                                       pipeDiam2, flangeDiam2, _
                                       parFlowDiameter) + 2 * parInsulationThickness
    Set ObjInsHorCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHorCylinder
    Set ObjInsHorCylinder = Nothing
    
' Insert your code for output 2(Insulation for Vertical Cylinder)
    stPoint.Set 0, parFlowDiameter / 2, 0
    enPoint.Set 0, parInstrumentHeight + parInsulationThickness, 0
    parInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    Set ObjInsVerCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVerCylinder
    Set ObjInsVerCylinder = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
