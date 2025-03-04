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
'   Author:          MS
'   Creation Date:  Monday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IWFE, PDMS Category Name=IFED)
'   The symbol consists of eight outputs, five physical and three insulation outputs.
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjFlowElementBodyIns As Object
    Dim ObjNozz3Ins As Object
    Dim ObjNozz4Ins As Object
    
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parInstrumentHeight As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)        'D2
    parFace1toCenter = arrayOfInputs(3)     'D3
    parNozzleOffset = arrayOfInputs(4)        'D4
    parNozzleOffset1 = arrayOfInputs(5)       'D5
    parInstrumentHeight = arrayOfInputs(6)  'D6
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 2(Flow Element Body Insulation)
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    parInsulationDiameter = maxDia + 2 * parInsulationThickness
    Set ObjFlowElementBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowElementBodyIns
    Set ObjFlowElementBodyIns = Nothing
    
 ' Insert your code for output 3(Nozz3 Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set -parNozzleOffset1 + parNozzleOffset, 0, 0
    enPoint.Set -parNozzleOffset1 + parNozzleOffset, parInstrumentHeight, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjNozz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz3Ins
    Set ObjNozz3Ins = Nothing
    
 ' Insert your code for output 4(Nozz4 Insulation)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set parNozzleOffset, 0, 0
    enPoint.Set parNozzleOffset, parInstrumentHeight, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjNozz4Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz4Ins
    Set ObjNozz4Ins = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
