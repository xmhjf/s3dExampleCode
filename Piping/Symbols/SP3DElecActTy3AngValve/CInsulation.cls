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
'   Creation Date:  Wednesday, Oct 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVCT).
'   The symbol consists of Fourteen outputs, Thirteen physical outputs and One Insulation Output.
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
    Dim parMotor4Offset As Double
    Dim parMotor4Length As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parHandWheelDiameter As Double
    Dim parMotor1Diameter As Double
    Dim parMotor4Diameter As Double
    Dim parMotor3Length As Double
    Dim parMotor1Offset As Double
    Dim parMotor2Offset As Double
    Dim parMotor1Length As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parMotor3toCenter As Double
    Dim parActuatorHeight3 As Double
    Dim parInsulationThickness As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parMotor4Offset = arrayOfInputs(5)  'D5
    parMotor4Length = arrayOfInputs(6)  'D6
    parActuatorHeight = arrayOfInputs(7)    'D7
    parActuatorHeight1 = arrayOfInputs(8)   'D8
    parActuatorHeight2 = arrayOfInputs(9)   'D9
    parHandWheelDiameter = arrayOfInputs(10)    'D10
    parMotor1Diameter = arrayOfInputs(11)   'D11
    parMotor4Diameter = arrayOfInputs(12)   'D12
    parMotor3Length = arrayOfInputs(13) 'D13
    parMotor1Offset = arrayOfInputs(14) 'D14
    parMotor2Offset = arrayOfInputs(15) 'D15
    parMotor1Length = arrayOfInputs(16) 'D16
    parActuatorWidth = arrayOfInputs(17)    'D17
    parActuatorWidth1 = arrayOfInputs(18)   'D18
    parMotor3toCenter = arrayOfInputs(19)   'D19
    parActuatorHeight3 = arrayOfInputs(20)  'D20
    parInsulationThickness = arrayOfInputs(21)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Valve Horizontal Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    InsulationDia = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDia = flangeDiam + parInsulationThickness * 2
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1Ins
    Set ObjBodyCone1Ins = Nothing
    
 ' Insert your code for output 2(Body Cone 2 Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    InsulationDia = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then InsulationDia = flangeDiam + parInsulationThickness * 2
    stPoint.Set 0, -parFace2toCenter, 0
    enPoint.Set 0, parCylHeight + parInsulationThickness, 0
    Set ObjBodyCone2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
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
