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
'   Author:          BG Govindrajan
'   Creation Date:  Wednesday, Mar 27 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This Symbol is constructed using 8 outputs. 3 outputs for Physical Aspect and 5 for
'Insulation aspect. Physical Aspect is created using: 1. Nozzle with length for port 1,2.Nozzle for
' Port 2 and 3.Branch using Nozzle with Length for Port 3
'For Insulation Aspect Using: 1.Port 1 Insulation, 2.Main Body Insulation, 3.Port 2 Insulation
'4.Branch Insulation and 5.Port 3 Insulation.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim flangeDiam3      As Double
    Dim depth3           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjPort1Insul As Object
    Dim ObjMainBodyInsul As Object
    Dim ObjPort2Insul As Object
    Dim ObjBranchInsul As Object
    Dim ObjPort3Insul As Object
    
    Dim parInsulationThickness As Double
    Dim parInsulationDiameter As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
''''' ' Insert your code for output 4(Port1 Insulation)
'''''    Dim stPoint   As New AutoMath.DPosition
'''''    Dim enPoint   As New AutoMath.DPosition
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
'''''    stPoint.Set -parFacetoCenter, 0, 0
'''''    enPoint.Set -parFacetoCenter + flangeThick1 + parInsulationThickness, 0, 0
'''''    parInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
'''''    Set ObjPort1Insul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
'''''' Set the output
'''''    iOutput = iOutput + 1
'''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort1Insul
'''''    Set ObjPort1Insul = Nothing
'''''    Set stPoint = Nothing
'''''    Set enPoint = Nothing
    
' Insert your code for output 5(Main Body Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    Dim stPtBody   As New AutoMath.DPosition
    Dim enPtBody   As New AutoMath.DPosition
    
    stPtBody.Set -parFacetoCenter + flangeThick1, 0, 0
    enPtBody.Set parFacetoCenter - flangeThick2, 0, 0
    parInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
    Set ObjMainBodyInsul = PlaceCylinder(m_OutputColl, stPtBody, enPtBody, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainBodyInsul
    Set ObjMainBodyInsul = Nothing
    Set stPtBody = Nothing
    Set enPtBody = Nothing
    
'''''' Insert your code for output 6(Port2 Insulation)
'''''    Dim stPtP2   As New AutoMath.DPosition
'''''    Dim enPtP2   As New AutoMath.DPosition
'''''
'''''    stPtP2.Set parFacetoCenter, 0, 0
'''''    enPtP2.Set parFacetoCenter - flangeThick2 - parInsulationThickness, 0, 0
'''''    parInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
'''''    Set ObjPort2Insul = PlaceCylinder(m_OutputColl, stPtP2, enPtP2, parInsulationDiameter, True)
'''''' Set the output
'''''    iOutput = iOutput + 1
'''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort2Insul
'''''    Set ObjPort2Insul = Nothing
'''''    Set stPtP2 = Nothing
'''''    Set enPtP2 = Nothing
    
 ' Insert your code for output 7(Branch Body Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
    Dim stPtBr   As New AutoMath.DPosition
    Dim enPtBr   As New AutoMath.DPosition
    
    stPtBr.Set 0, -(pipeDiam1 * 0.65) - parInsulationThickness, 0
    enPtBr.Set 0, parFace1toCenter - flangeThick3, 0
    parInsulationDiameter = pipeDiam3 + parInsulationThickness * 2
    Set ObjBranchInsul = PlaceCylinder(m_OutputColl, stPtBr, enPtBr, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsul
    Set ObjBranchInsul = Nothing
    Set stPtBr = Nothing
    Set enPtBr = Nothing

'''''' Insert your code for output 8(Port 3 Insulation)
'''''    Dim stPtP3   As New AutoMath.DPosition
'''''    Dim enPtP3   As New AutoMath.DPosition
'''''
'''''    stPtP3.Set 0, parFace1toCenter, 0
'''''    enPtP3.Set 0, parFace1toCenter - flangeThick3 - parInsulationThickness, 0
'''''    parInsulationDiameter = flangeThick3 + parInsulationThickness * 2
'''''    Set ObjPort3Insul = PlaceCylinder(m_OutputColl, stPtP3, enPtP3, parInsulationDiameter, True)
'''''' Set the output
'''''    iOutput = iOutput + 1
'''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Insul
'''''    Set ObjPort3Insul = Nothing
'''''    Set stPtP3 = Nothing
'''''    Set enPtP3 = Nothing
'''''    Set CenterPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
