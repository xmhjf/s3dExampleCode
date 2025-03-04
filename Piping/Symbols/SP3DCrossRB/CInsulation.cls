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
'   Author:          BG
'   Creation Date:  Wednesday, Jan 2 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   19.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
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
    
    Dim iOutput     As Double
    Dim ObjPort1Insulation As Object
    Dim ObjPort2Insulation As Object
    Dim ObjPort3Insulation As Object
    Dim ObjPort4Insulation As Object
    Dim ObjMainRunInsul As Object
    Dim ObjBranch1Insul As Object
    Dim ObjBranch2Insul As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 5(Port in Main Run Insulation)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort1Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort1Insulation
    Set ObjPort1Insulation = Nothing
    
 ' Insert your code for output 9(Main Run Body Insulation)
    stPoint.Set (parFace1toCenter - flangeThick), 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainRunInsul
    Set ObjMainRunInsul = Nothing
    
 ' Insert your code for output 6(Port in Main Run Insulation)
    stPoint.Set parFace1toCenter, 0, 0
    enPoint.Set (parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort2Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort2Insulation
    Set ObjPort2Insulation = Nothing
    
 ' Insert your code for output 7(Port in Branch Run Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, parFace2toCenter, 0
    enPoint.Set 0, (parFace2toCenter - flangeThick - parInsulationThickness), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort3Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Insulation
    Set ObjPort3Insulation = Nothing
    
' Insert your code for output 10(Brach1 Run Body Insulation )
    stPoint.Set 0, 0, 0
    enPoint.Set 0, (parFace2toCenter - flangeThick), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjBranch1Insul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranch1Insul
    Set ObjBranch1Insul = Nothing

' Insert your code for output 8(Port in Branch Run Insulation)
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, -(parFace2toCenter - flangeThick - parInsulationThickness), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort4Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4Insulation
    Set ObjPort4Insulation = Nothing
     
 ' Insert your code for output 11(Brach2 Run Body Insulation )
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -(parFace2toCenter - flangeThick), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjBranch2Insul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranch2Insul
    Set ObjBranch2Insul = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
