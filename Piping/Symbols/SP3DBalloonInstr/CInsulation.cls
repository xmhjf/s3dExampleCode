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
'   Creation Date:  Thursday, May 30 2002
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
    Dim ObjInstrBody As Object
    Dim ObjStem As Object
    Dim objCone As Object
    Dim ObjNozzl1Ins As Object
    Dim ObjNozzle2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
    Dim parConeHeight As Double
    Dim parConeDiameter1 As Double
    Dim parConeDiameter2 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parConeHeight = arrayOfInputs(5)
    parConeDiameter1 = arrayOfInputs(6)
    parConeDiameter2 = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

 ' Insert your code for output 1(Instruments Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    Dim startPoint As AutoMath.DPosition
    Dim endPoint  As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    endPoint.Set parFace2toCenter - flangeThick- parInsulationThickness, 0, 0
    Set ObjInstrBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, pipeDiam + parInsulationThickness * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstrBody
    Set ObjInstrBody = Nothing
    
 ' Insert your code for output 4(Nozzle 1 Insulation)
    Dim parInsulationDiameter As Double
    
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjNozzl1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzl1Ins
    Set ObjNozzl1Ins = Nothing
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
 
 ' Insert your code for output 5(Nozzle 2 Insulation)
    startPoint.Set parFace2toCenter, 0, 0
    endPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    Set ObjNozzle2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2Ins
    Set ObjNozzle2Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
