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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Sep 16 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMFM2,
'   PDMS Category Name=IFMX). The symbol consists of seven outputs, five physical and
'   two insulation outputs. Insulation consists of a base (cylinder) and a top (box).
'   Change History:
'   dd.mmm.yyyy     who                   change description
'   -----------     -----                 ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   08.SEP.2006     KKC                   DI-95670 Replace names with initials in all revision history sheets and symbols
'   03.Aug.2007     VRG                   TR 124962 Insulation aspect is corrected so that it doesn't extend beyond face to center dimensions
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    
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
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick      As Double
    Dim sptOffset        As Double
    Dim flangeDiam1      As Double
    Dim depth            As Double
    
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    
    Dim iOutput           As Integer
    Dim ObjInsFlowMtrBase As Object
    Dim ObjInsFlowMtrTop  As Object
    
    Dim parFace1toCenter  As Double
    Dim parFace2toCenter  As Double
    Dim parFlowDiameter   As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentWidth As Double
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
    parInstrumentLength = arrayOfInputs(7)   'D7
    parInstrumentWidth = arrayOfInputs(8)    'D8
    parInsulationThickness = arrayOfInputs(9)

    
    iOutput = 0

' Insert your code for output 1(Insulation for Flow Meter Base)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
'   Use maximum value among pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2 and parFlowDiameter
'   for computing insulation diameter
    parInsulationDiameter = ReturnMax5(pipeDiam1, flangeDiam1, pipeDiam2, flangeDiam2, parFlowDiameter) + _
                                                2 * parInsulationThickness
    Set ObjInsFlowMtrBase = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFlowMtrBase
    Set ObjInsFlowMtrBase = Nothing
    
' Insert your code for output 2(Insulation for Flow Meter Top)
'   Note: Insulation box is considered to start from the Flow Meter Base to the Box top.
'   So separate insulation is not provided for the connector.
    stPoint.Set -parInstrumentLength / 2 - parInsulationThickness, _
                 parFlowDiameter / 2, _
                -parInstrumentWidth / 2 - parInsulationThickness
    enPoint.Set parInstrumentLength / 2 + parInsulationThickness, _
                parInstrumentHeight + parInsulationThickness, _
                parInstrumentWidth / 2 + parInsulationThickness
    Set ObjInsFlowMtrTop = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFlowMtrTop
    Set ObjInsFlowMtrTop = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
