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
'   Creation Date:  Friday, Sep 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR9,
'   PDMS Category Name=IFNO). Physical aspect consist of one vertical cylinder, one horizontal cylinder
'   on right side of the vertical cylinder and two nozzles both facing -negative direction direction of X-axis.
'   Insulation aspect consist of two cylinders one along X-axis and the other along Y-axis.
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
    Dim ObjInsVerCylinder As Object
    Dim ObjInsHoriCylinder As Object
    
    Dim parInstrumentHeight As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight1 As Double
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)    'D2
    parFlowDiameter = arrayOfInputs(3)        'D3
    parInstrumentHeight1 = arrayOfInputs(4)   'D4
    parFacetoFace = arrayOfInputs(5)          'D5
    parInstrumentLength = arrayOfInputs(6)    'D6
    parInstrumentDiameter = arrayOfInputs(7)  'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Vertical Cylinder)
    stPoint.Set 0, -parInstrumentHeight1 - parInsulationThickness, 0
    enPoint.Set 0, parInstrumentHeight - parInstrumentHeight1 + parInsulationThickness, 0
    parInsulationDiameter = parFlowDiameter + 2 * parInsulationThickness
    Set ObjInsVerCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVerCylinder
    Set ObjInsVerCylinder = Nothing
    
' Insert your code for output 2(Insulation for Horizontal Cylinder)
    Dim radialPlaceDist As Double
    radialPlaceDist = Sqr(Abs((parFlowDiameter / 2) ^ 2 - (parInstrumentDiameter / 2) ^ 2))

    stPoint.Set radialPlaceDist, 0, 0
    enPoint.Set parInstrumentLength + parInsulationThickness, 0, 0
    parInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    Set ObjInsHoriCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHoriCylinder
    Set ObjInsHoriCylinder = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
