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
'   Author:         MS
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No.: S49AC, Page No. D-105 of PDS Piping Component Data Reference Guide.
'   This Symbol is created with Thirteen Outputs, Five of which are physical aspect outputs.
'   Branch Body, Strainer Body using 'PlaceCylinder' function,
'   Three ObjNozzle object by using 'CreateNozzle' function.
'   And all the Insulation aspect ouputs using 'PlaceCylinder' function.
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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

    Dim iOutput     As Double
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object
    Dim ObjPipeinNoz1Ins As Object
    Dim ObjPipeinNoz2Ins As Object
    Dim ObjPipeinNoz3Ins As Object
    Dim ObjNoz1Ins As Object
    Dim ObjNoz2Ins As Object
    Dim ObjNoz3Ins As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parStrToptoCenter As Double
    Dim parStrDiameter As Double
    Dim parBranchDiameter As Double
    Dim parStrWidth As Double
    Dim parFace3toStrTop As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parStrToptoCenter = arrayOfInputs(4)
    parStrDiameter = arrayOfInputs(5)
    parBranchDiameter = arrayOfInputs(6)
    parStrWidth = arrayOfInputs(7)
    parFace3toStrTop = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim dInsulationDia As Double

' Insert your code for output 3(Pipe portion in Noz1, Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjPipeinNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz1Ins
    Set ObjPipeinNoz1Ins = Nothing

 ' Insert your code for output 4(Noz1 Insulation)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
    Set ObjNoz1Ins = Nothing

 ' Insert your code for output 5(Pipe portion in Noz2, Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set parFace2toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjPipeinNoz2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz2Ins
    Set ObjPipeinNoz2Ins = Nothing

 ' Insert your code for output 6(Noz2 Insulation)
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
    Set ObjNoz2Ins = Nothing

' Insert your code for output 7(Pipe portion in Noz3, Insulation)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, parStrToptoCenter + parFace3toStrTop - flangeThick, 0
    enPoint.Set 0, parStrToptoCenter, 0

    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjPipeinNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz3Ins
    Set ObjPipeinNoz3Ins = Nothing

 ' Insert your code for output 8(Noz3 Insulation)
     stPoint.Set 0, parStrToptoCenter + parFace3toStrTop - flangeThick - parInsulationThickness, 0
    enPoint.Set 0, parStrToptoCenter + parFace3toStrTop, 0

    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz3Ins
    Set ObjNoz3Ins = Nothing

' Insert your code for output 9(Insulation for Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parStrToptoCenter - parStrWidth, 0
    dInsulationDia = parBranchDiameter + 2 * parInsulationThickness

    Set ObjInsBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing

 ' Insert your code for output 10(Insulation for Strainer)
     stPoint.Set 0, parStrToptoCenter - parStrWidth - parInsulationThickness, 0
    enPoint.Set 0, parStrToptoCenter + parInsulationThickness, 0
    dInsulationDia = parStrDiameter + 2 * parInsulationThickness
    Set ObjInsStrainer = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainer
    Set ObjInsStrainer = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
