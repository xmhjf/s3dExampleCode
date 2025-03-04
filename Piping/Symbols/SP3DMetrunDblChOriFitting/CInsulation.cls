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
'   Creation Date:  Wednesday, Sep 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMDC2,
'   PDMS Category Name=IFRC). The symbol consists of 11  outputs, nine physical and
'   two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1     As Double
    Dim depth           As Double
    
    Dim pipeDiam2   As Double
    Dim flangeDiam2 As Double

    Dim iOutput     As Double
    Dim ObjInsBody As Object
    Dim ObjInsPipeLine As Object
    
    Dim parInstrumentLength      As Double
    Dim parInstrumentLength1     As Double
    Dim parInstrumentHeight      As Double
    Dim parInstrumentHeight1     As Double
    Dim parInstrumentLength2     As Double
    Dim parInstrumentLength3     As Double
    Dim parInstrumentWidth       As Double
    Dim parInstrumentWidth1      As Double
    Dim parArmLength             As Double
    Dim parNozzleOffset          As Double
    Dim parNozzleOffset1         As Double
    Dim parArmHeight             As Double
    Dim parArmHeight1            As Double
    Dim parInstrumentLength4     As Double
    Dim parInstrumentLength5     As Double
    Dim parInsulationThickness   As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)      'D2
    parInstrumentLength1 = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength2 = arrayOfInputs(6)     'D6
    parInstrumentLength3 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parArmLength = arrayOfInputs(10)            'D10
    parNozzleOffset = arrayOfInputs(11)         'D11
    parNozzleOffset1 = arrayOfInputs(12)        'D12
    parArmHeight = arrayOfInputs(13)            'D14
    parArmHeight1 = arrayOfInputs(14)           'D15
    parInstrumentLength4 = arrayOfInputs(15)    'D16
    parInstrumentLength5 = arrayOfInputs(16)    'D17
    parInsulationThickness = arrayOfInputs(17)
        
    iOutput = 0

' Insert your code for output 1(InsBody)
    stPoint.Set -parInstrumentLength2 - parInsulationThickness, _
                -parInstrumentHeight - parInsulationThickness, _
                parInstrumentWidth + parInsulationThickness

    enPoint.Set parInstrumentLength3 + parInsulationThickness, _
                parInstrumentHeight1 + parInsulationThickness, _
                -parInstrumentWidth1 - parInsulationThickness
    Set ObjInsBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    Set ObjInsBody = Nothing
    
' Insert your code for output 2(InsPipeLine)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth

    stPoint.Set -parInstrumentLength - parInstrumentLength4, 0, 0
    enPoint.Set parInstrumentLength1 + parInstrumentLength5, 0, 0
'   Assumption: Instrument Left Hand Side/Right Hand Side flanges (At distances D2 and D3 respectively)
'   diameter is assumed to be 1.2 times corresponding pipe diameter.
    Dim LHSFlangeDiameter As Double
    Dim RHSFlangeDiameter As Double
    
    LHSFlangeDiameter = 1.2 * pipeDiam1
    RHSFlangeDiameter = 1.2 * pipeDiam2
    parInsulationDiameter = ReturnMax4(flangeDiam1, LHSFlangeDiameter, RHSFlangeDiameter, flangeDiam2) + _
                            2 * parInsulationThickness
    Set ObjInsPipeLine = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPipeLine
    Set ObjInsPipeLine = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
