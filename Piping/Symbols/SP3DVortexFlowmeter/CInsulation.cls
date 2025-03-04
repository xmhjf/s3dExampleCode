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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Jul 4 2006
'   Description:
'        Vortex Flowmeter (Source: Model 8800C in Vortex Flow Meter.pdf at www.emersonprocess.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    
    Dim iOutput             As Double
    Dim ObjInsFlowmeterBody  As Object
    Dim ObjInsConnectorBody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object

    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)    'A
    parInstrumentHeight = arrayOfInputs(3)   'C
    parInstrumentDiameter = arrayOfInputs(4)
    parInstrumentWidth = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Flowmeter  Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ConnectorDiameter   As Double
    Dim dInsulationDiameter   As Double
    
    stPoint.Set 0, parInstrumentHeight, -parInstrumentWidth / 2 - parInsulationThickness
    enPoint.Set 0, parInstrumentHeight, parInstrumentWidth / 2 + parInsulationThickness
    dInsulationDiameter = parInstrumentDiameter + 2 * parInsulationThickness
    Set ObjInsFlowmeterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFlowmeterBody
    Set ObjInsFlowmeterBody = Nothing
    
' Insert your code for output 2(Insulation for Connector  Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parInstrumentHeight, 0
'   Set ConnectorDiameter equal to 0.2 * minimum of pipeDiam and parInstrumentDiameter
    ConnectorDiameter = 0.2 * parInstrumentDiameter
    If CmpDblLessThan(parInstrumentDiameter, pipeDiam) Then ConnectorDiameter = 0.2 * pipeDiam
    dInsulationDiameter = ConnectorDiameter + 2 * parInsulationThickness
    Set ObjInsConnectorBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsConnectorBody
    Set ObjInsConnectorBody = Nothing
    
' Insert your code for output 3(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
    dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 4(Insulation for body Port1 Side)
    stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 5(Insulation Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
    dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 6(Insulation for body Port2 Side)
    stPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
