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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Monday, Nov 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is prepared based on Resistoflex Flanged Lined Pipe & Fittings catalog.
'   The symbol consists of seven outputs, Four physical and  Three Insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added. 
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure. 
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
    Dim raisedFaceDiam As Double
    
    Dim iOutput     As Double
    Dim ObjOrificeBodyIns As Object
    Dim ObjPort3BodyIns As Object
    Dim ObjPort3Ins As Object
    
    Dim parFacetoFace As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
 ' Insert your code for output 2(Orifice Body Insulation)
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth, raisedFaceDiam
    If flangeThick = 0 Then flangeThick = 0.0001
    Dim InsulationDia As Double
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    InsulationDia = raisedFaceDiam + 2 * parInsulationThickness
    Set ObjOrificeBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOrificeBodyIns
    Set ObjOrificeBodyIns = Nothing
    
 ' Insert your code for output 3(Port3 Body Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001

    stPoint.Set 0, parFace3toCenter - flangeThick, 0
    enPoint.Set 0, raisedFaceDiam / 2, 0
    InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3BodyIns
    Set ObjPort3BodyIns = Nothing
    
 ' Insert your code for output 4(Port3 Insulation)
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, parFace3toCenter - flangeThick - parInsulationThickness, 0
    InsulationDia = flangeDiam + 2 * parInsulationThickness
    If pipeDiam > flangeDiam Then InsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjPort3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Ins
    Set ObjPort3Ins = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
