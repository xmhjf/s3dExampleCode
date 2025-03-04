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
'   Author:          MS
'   Creation Date:  Tuesday, Oct 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVCU)
'  The symbol consists of Twelve outputs; Ten physical outputs and Two Insulation Output
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1Ins As Object
    Dim ObjBodyCone2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parMotor1Diameter As Double
    Dim parMotor2Diameter As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Length As Double
    Dim parMotor1Length As Double
    Dim parActuatorOffset As Double
    Dim parActuatorHeight1 As Double
    Dim parHandWheelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)      'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parActuatorHeight = arrayOfInputs(5)      'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorLength = arrayOfInputs(7)      'D7
    parActuatorLength1 = arrayOfInputs(8)     'D8
    parMotor1Diameter = arrayOfInputs(9)      'D9
    parMotor2Diameter = arrayOfInputs(10)     'D10
    parMotor2Offset = arrayOfInputs(11)   'D11
    parMotor1Offset = arrayOfInputs(12)   'D12
    parMotor2Length = arrayOfInputs(13)   'D13
    parMotor1Length = arrayOfInputs(14)   'D14
    parActuatorOffset = arrayOfInputs(15)     'D15
    parActuatorHeight1 = arrayOfInputs(16)     'D16
    parHandWheelDiameter = arrayOfInputs(17)      'D17
    parInsulationThickness = arrayOfInputs(18)
    
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
 ' Insert your code for output 9(Body Cone 1 Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 10(Body Cone 2 Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, parCylHeight + parInsulationThickness, 0
    Set ObjBodyCone2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

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
