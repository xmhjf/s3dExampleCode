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
'   Author:          NN
'   Creation Date:  Thursday, June 14 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   18.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
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
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
   
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedPort3 As Object
    
    Dim parFacetoFace As Double
    Dim parFace2toBranch As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace2toBranch = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

' Insert your code for output 1(Insulated Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -(parFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
' Insert your code for output 2(Insulated Port1)
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -(parFacetoFace / 2 - flangeThick - parInsulationThickness), 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 3(Insulated Branch Port2)
    Dim BranchPoint   As New AutoMath.DPosition
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    BranchPoint.Set (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), _
            -(parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), 0
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)
  ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing

' Insert your code for output 4(Insulated Port2)
    stPoint.Set parFace2toBranch * Sqr(1 / 2), -parFace2toBranch * Sqr(1 / 2), 0
    enPoint.Set (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), _
            -(parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing

' Insert your code for output 5(Insulated Branch Port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001

    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    BranchPoint.Set (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), 0
    Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)
  ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
    Set ObjInsulatedBranch = Nothing
  
' Insert your code for output 6 (Insulated Port3)
    stPoint.Set parFace2toBranch * Sqr(1 / 2), parFace2toBranch * Sqr(1 / 2), 0
    enPoint.Set (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), (parFace2toBranch - flangeThick - parInsulationThickness) * Sqr(1 / 2), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
