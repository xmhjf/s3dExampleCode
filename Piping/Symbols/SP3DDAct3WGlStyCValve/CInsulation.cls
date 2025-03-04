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
'   Author:        ACM
'   Creation Date:  Wednesday, Oct 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ISSD3),
'   The symbol consists of 10 outputs,8 physical outputs
'   and 2 Insulation Output.
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
    
    Dim ObjBodyCone1Ins As Object
    Dim ObjBodyCone2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parFace3toCenter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight1 = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

 ' Insert your code for output 6(InsulationBody cone 1)
   
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
 
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    Dim CylEdgeToCen As Double
    Dim CylRad As Double
    
    CylRad = parFace1toCenter / 2
    CylEdgeToCen = Sqr(parActuatorHeight1 ^ 2 + CylRad ^ 2)
    
    Dim maxRad As Double
    maxRad = ReturnMax5(pipeDiam1 / 2, pipeDiam2 / 2, flangeDiam1 / 2, flangeDiam2 / 2, CylEdgeToCen)
    
    InsulationDia = maxRad * 2 + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFacetoFace - parFace1toCenter, 0, 0
    
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 7(Body Cone 2 Insulation)
    Dim pipeDiam3        As Double
    Dim flangeDiam3        As Double
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset, depth
    
    InsulationDia = pipeDiam3 + parInsulationThickness * 2
    If flangeDiam3 > pipeDiam3 Then InsulationDia = flangeDiam3 + 2 * parInsulationThickness
       
    stPoint.Set 0, -parFace3toCenter, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2Ins
    Set ObjBodyCone2Ins = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
     
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
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

