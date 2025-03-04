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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMRSV, PDMS Category Name=IFRD)
'   The symbol consists of eleven outputs, eight physical and three insulation outputs.
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
    Dim ObjCylinderBody1 As Object
    Dim ObjCylinderBody2 As Object
    Dim ObjReducingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjVenturiBodyIns As Object
    Dim ObjNozz3BodyIns As Object
    Dim ObjNozz4BodyIns As Object
    
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentDiameter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzleOffset2 As Double
    Dim parNozzleOffset3 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInstrumentLength4 As Double
    Dim parInstrumentLength5 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)  'D2
    parInstrumentLength1 = arrayOfInputs(3) 'D3
    parInstrumentLength2 = arrayOfInputs(4) 'D4
    parInstrumentLength3 = arrayOfInputs(5) 'D5
    parInstrumentDiameter = arrayOfInputs(6)       'D6
    parNozzleOffset = arrayOfInputs(7)     'D7
    parNozzleOffset1 = arrayOfInputs(8)     'D8
    parNozzleOffset2 = arrayOfInputs(9)        'D9
    parNozzleOffset3 = arrayOfInputs(10)  'D10
    parPortRotation1 = arrayOfInputs(11)                'D11
    parPortRotation2 = arrayOfInputs(12)         'D12
    parInstrumentLength4 = arrayOfInputs(13)        'D13
    parInstrumentLength5 = arrayOfInputs(14)        'D14
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0
'The orgin is considered to be along the flow axis and mid of parInstrumentLength2
 ' Insert your code for output 5(Venturi Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    Dim InsulationDiameter As Double
    'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2
    Dim maxDia As Double
    maxDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    InsulationDiameter = maxDia + 2 * parInsulationThickness
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength - parInstrumentLength4, 0, 0
    enPoint.Set parInstrumentLength2 / 2 + parInstrumentLength3 + parInstrumentLength5, 0, 0
    Set ObjVenturiBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVenturiBodyIns
    Set ObjVenturiBodyIns = Nothing
    
 ' Insert your code for output 6(Nozzle3 Body Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set -(parNozzleOffset1 - parNozzleOffset), _
                        (parNozzleOffset2) * Sin(parPortRotation1), _
                        (parNozzleOffset2) * Cos(parPortRotation1)
    enPoint.Set -(parNozzleOffset1 - parNozzleOffset), 0, 0
    InsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDiameter = flangeDiam + parInsulationThickness * 2
     
   Set ObjNozz3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz3BodyIns
    Set ObjNozz3BodyIns = Nothing
    
 ' Insert your code for output 7(Nozzle4 Body Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set 0, _
                        (parNozzleOffset3) * Sin(parPortRotation2), _
                        (parNozzleOffset3) * Cos(parPortRotation2)
    enPoint.Set 0, 0, 0
    InsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDiameter = flangeDiam + parInsulationThickness * 2
   Set ObjNozz4BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz4BodyIns
    Set ObjNozz4BodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
Private Function ReturnMax4(A As Double, B As Double, C As Double, D As Double) As Double
    Dim MaxValue As Double
    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

