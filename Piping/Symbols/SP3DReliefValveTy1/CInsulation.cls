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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT1, PDMS Category Name=IVPG)
'   The symbol consists of eight outputs,  five physical and three insulation outputs.
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
    Dim ObjValveBody As Object
    Dim ObjTopCylBody1 As Object
    Dim ObjTopCylBody2 As Object
    Dim ObjValveBodyIns As Object
    Dim ObjNozz1BodyIns As Object
    Dim ObjNozz2BodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim Diameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentHeight2 = arrayOfInputs(10)    'D10
    parInstrumentHeight3 = arrayOfInputs(11)    'D11
    parInstrumentDiameter = arrayOfInputs(12)   'D12
    parInstrumentDiameter1 = arrayOfInputs(13)  'D13
    parInsulationThickness = arrayOfInputs(14)
    
    iOutput = 0
'The origin is considered to be at intersection of nozzle axis
 ' Insert your code for output 4(Valve Body Insulation)
    stPoint.Set -parInstrumentLength1 - parInsulationThickness, _
                        -parInstrumentHeight - parInsulationThickness, _
                        parInstrumentWidth + parInsulationThickness

    enPoint.Set parInstrumentLength + parInsulationThickness, _
                        parInstrumentHeight1 + parInstrumentHeight3 + parInstrumentHeight2 + parInsulationThickness, _
                        -parInstrumentWidth1 - parInsulationThickness

    Set ObjValveBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyIns
    Set ObjValveBodyIns = Nothing
    
 ' Insert your code for output 5(Nozzle1 Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim InsulationDiameter As Double
    InsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDiameter = flangeDiam + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parInstrumentLength1, 0, 0
    Set ObjNozz1BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz1BodyIns
    Set ObjNozz1BodyIns = Nothing
    
 ' Insert your code for output 6(Nozzle2 Body Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    InsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDiameter = flangeDiam + parInsulationThickness * 2
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, -parInstrumentHeight, 0
    Set ObjNozz2BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz2BodyIns
    Set ObjNozz2BodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
