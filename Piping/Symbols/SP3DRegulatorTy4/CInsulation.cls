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
'   Creation Date:  Monday, Sep 16 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG4, PDMS Category Name=IVRI)
'   The symbol consists of  eight outputs, seven physical and one insulation outputs.
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjConeBody1 As Object
    Dim ObjConeBody2 As Object
    Dim ObjConeBody3 As Object
    Dim ObjConeBody4 As Object
    Dim ObjTopCylBody As Object
    Dim ObjRegulatorBodyIns As Object
    Dim ObjConeBody4Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight2 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFacetoFace = arrayOfInputs(3)        'D3
    parInstrumentHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight1 = arrayOfInputs(5)  'D5
    parInstrumentDiameter = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)  'D9
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'The origin is taken to be along the flow axis and mid of parFacetoFace
 ' Insert your code for output 6(Regulator Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    stPoint.Set -parFace1toCenter, _
                        -maxDia / 2 - parInsulationThickness, _
                        maxDia / 2 + parInsulationThickness

    enPoint.Set parFacetoFace - parFace1toCenter, _
                        parInstrumentHeight2 + parInsulationThickness, _
                        -maxDia / 2 - parInsulationThickness

    Set ObjRegulatorBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRegulatorBodyIns
    Set ObjRegulatorBodyIns = Nothing
    
 ' Insert your code for output 7(Cone Body4Ins)
    Dim coneApexHalfAngle As Double
    coneApexHalfAngle = Atn((parInstrumentDiameter / 2) / parInstrumentHeight1)
    stPoint.Set 0, -parInstrumentHeight1 - parInsulationThickness, 0
    enPoint.Set 0, 0, 0
    Set ObjConeBody4Ins = PlaceCone(m_OutputColl, stPoint, enPoint, _
                                                parInstrumentDiameter / 2 + (parInsulationThickness / Cos(coneApexHalfAngle)) + (parInsulationThickness * Tan(coneApexHalfAngle)), _
                                                                            (parInsulationThickness / Cos(coneApexHalfAngle)), True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeBody4Ins
    Set ObjConeBody4Ins = Nothing
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
