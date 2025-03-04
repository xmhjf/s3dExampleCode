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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Monday, Sep 16 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMFM1, PDMS Category Name=IMFW)
'   The symbol consists of six outputs, four physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   08.SEP.2006     KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   03.Aug.2007     VRG                 TR 124962 Insulation aspect is corrected so that it doesn't extend beyond face to center dimensions
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
  
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Integer
    Dim ObjFlowMeterBody As Object
    Dim ObjFlowMeterTopBody As Object
    Dim ObjFlowMeterBodyIns As Object
    Dim ObjFlowMeterTopBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentWidth As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentHeight = arrayOfInputs(5)  'D5
    parInstrumentLength = arrayOfInputs(6)  'D6
    parInstrumentWidth = arrayOfInputs(7)   'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'The origin is taken to be along the flow axis and center of parFlowDiameter
 ' Insert your code for output 3(Flow Meter Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
'maxDia refers to the maximum diameter of pipe and flange for nozz 1 and nozz2 and also parFlowDiameter
    Dim maxDia As Double
    maxDia = ReturnMax5(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2, parFlowDiameter)
    InsulationDia = maxDia + 2 * parInsulationThickness
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjFlowMeterBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterBodyIns
    Set ObjFlowMeterBodyIns = Nothing
    
 ' Insert your code for output 4(Flow Meter Top Body Insulation)
     stPoint.Set -parInstrumentLength / 2 - parInsulationThickness, _
                        0, _
                        -parInstrumentWidth / 2 - parInsulationThickness

     enPoint.Set parInstrumentLength / 2 + parInsulationThickness, _
                        parInstrumentHeight + parInsulationThickness, _
                        parInstrumentWidth / 2 + parInsulationThickness


    Set ObjFlowMeterTopBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterTopBodyIns
    Set ObjFlowMeterTopBodyIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

    
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
