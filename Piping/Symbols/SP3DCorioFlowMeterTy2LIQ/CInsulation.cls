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
'   Author:          BG
'   Creation Date:  Monday, Oct 7 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM2,
'   PDMS Category Name=IFNA). Physical aspect consist of 1 main cylinder body, 2 Support
'   Cylinder and a Box. Insulation aspect consist of a Cylinder and a box
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
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjFlowMtrBody As Object
    Dim ObjCylinder1 As Object
    Dim ObjCylinder2 As Object
    Dim ObjBox As Object
    Dim ObjBodyIns As Object
    Dim ObjBoxIns As Object
    Dim ObjCylin1Ins As Object
    Dim ObjCylin2Ins As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)       'D2
    parFlowDiameter = arrayOfInputs(3)       'D3
    parInstrumentHeight = arrayOfInputs(4)       'D4
    parInstrumentLength = arrayOfInputs(5)       'D5
    parInstrumentHeight1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInstrumentWidth1 = arrayOfInputs(8)       'D8
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

 ' Insert your code for output 5(Body Insulation)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Dim insulationDia  As Double
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    insulationDia = ReturnMax5(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2, parFlowDiameter) + _
                                parInsulationThickness * 2
    startPoint.Set -parFacetoFace / 2 - parInsulationThickness, 0, 0
    endPoint.Set parFacetoFace / 2 + parInsulationThickness, 0, 0
    Set ObjBodyIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
    Set ObjBodyIns = Nothing
    
' Insert your code for output 6(Cylinder 1 Insulation)
 'Center to center between the Cylinders is assumed to be 0.7 times of the variable parInstrumentLength
 'Cylinder diameter is assumed to be 0.1 times of the variable parInstrumentLength
    Dim cylinCenter As Double
    Dim cylinDiameter As Double
    
    cylinCenter = 0.7 * parInstrumentLength
    cylinDiameter = 0.1 * parInstrumentLength
    insulationDia = cylinDiameter + 2 * parInsulationThickness
    startPoint.Set -cylinCenter / 2, -((parFlowDiameter / 2) ^ 2 - (cylinDiameter / 2) ^ 2), 0
    endPoint.Set -cylinCenter / 2, -(parInstrumentHeight - parInstrumentHeight1), 0
    Set ObjCylin1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylin1Ins
    Set ObjCylin1Ins = Nothing
    
' Insert your code for output 7(Cylinder 2 Insulation)
    startPoint.Set cylinCenter / 2, -((parFlowDiameter / 2) ^ 2 - (cylinDiameter / 2) ^ 2), 0
    endPoint.Set cylinCenter / 2, -(parInstrumentHeight - parInstrumentHeight1), 0
    Set ObjCylin2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, insulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylin2Ins
    Set ObjCylin2Ins = Nothing
       
 ' Insert your code for output 8(Box Insulation)
    startPoint.Set -parInstrumentLength / 2 - parInsulationThickness, _
                            -parInstrumentHeight - parInsulationThickness, _
                            -parInstrumentWidth1 - parInsulationThickness
    endPoint.Set parInstrumentLength / 2 + parInsulationThickness, _
                            -(parInstrumentHeight - parInstrumentHeight1) + parInsulationThickness, _
                            parInstrumentWidth + parInsulationThickness
    Set ObjBoxIns = PlaceBox(m_OutputColl, startPoint, endPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoxIns
    Set ObjBoxIns = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
