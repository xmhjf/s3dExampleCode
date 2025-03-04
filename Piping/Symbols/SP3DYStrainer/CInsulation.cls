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
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No.: S51AC or S51AZ, Page No. D-105 PDS Piping Component Data Reference Guide.
'   Eight Insulation aspect outputs are created using 'PlaceCylinder' function.
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   26.OCT.2004     MS The input Offset is replaced with ComponentOffset.
'                                       Also the Nozzle 1 position in cphysical is corrected to
'                                       oPlacePoint.Set -parFace1toCenter + depth -sptOffset,0,0
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    Dim ObjInsPort3 As Object
    Dim ObjInsBody3 As Object
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parStrToptoCenter As Double
    Dim parStrDiameter As Double
    Dim parBranchDiameter As Double
    Dim parComponentOffset As Double
    Dim parStrWidth As Double
    Dim parFace3toStrTop As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsulationDia  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parStrToptoCenter = arrayOfInputs(4)
    parStrDiameter = arrayOfInputs(5)
    parBranchDiameter = arrayOfInputs(6)
    parComponentOffset = arrayOfInputs(7)
    parStrWidth = arrayOfInputs(8)
    parFace3toStrTop = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0

' Insert your code for output 1(Insulation for Port1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0

    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then dInsulationDia = flangeDiam + 2 * parInsulationThickness

    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing

' Insert your code for output 2( Insulation for Body Port1 side)
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing

' Insert your code for output 3(Insulation for Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing

' Insert your code for output 4( Insulation for Body Port2 side)
    stPoint.Set parFace2toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing

' Insert your code for output 5(Insulation for Port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set (parStrToptoCenter + parFace3toStrTop) * Cos(PI / 4) - parComponentOffset * Sin(PI / 4), _
                    (parStrToptoCenter + parFace3toStrTop) * Sin(PI / 4) + parComponentOffset * Cos(PI / 4), 0
    Dim Port3CPoint   As New AutoMath.DPosition
    Port3CPoint.Set stPoint.x, stPoint.y, 0
    enPoint.Set Port3CPoint.x - (flangeThick + parInsulationThickness) * Cos(PI / 4), _
                Port3CPoint.y - (flangeThick + parInsulationThickness) * Sin(PI / 4), 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort3
    Set ObjInsPort3 = Nothing

' Insert your code for output 6(Insulation for Body Port3 side)
    stPoint.Set Port3CPoint.x - (flangeThick + parInsulationThickness) * Cos(PI / 4), _
                Port3CPoint.y - (flangeThick + parInsulationThickness) * Sin(PI / 4), 0
    enPoint.Set parStrToptoCenter * Cos(PI / 4) - parComponentOffset * Sin(PI / 4), _
                parStrToptoCenter * Sin(PI / 4) + parComponentOffset * Cos(PI / 4), 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody3
    Set ObjInsBody3 = Nothing
    Set Port3CPoint = Nothing

' Insert your code for output 7(Insulation for Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set (parStrToptoCenter - parStrWidth) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth) * Sin(PI / 4), 0
    dInsulationDia = parBranchDiameter + 2 * parInsulationThickness
    Set ObjInsBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing

' Insert your code for output 8(Insulation for Strainer)
    stPoint.Set (parStrToptoCenter - parStrWidth - parInsulationThickness) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth - parInsulationThickness) * Sin(PI / 4), 0
    enPoint.Set (parStrToptoCenter + parInsulationThickness) * Cos(PI / 4), _
                (parStrToptoCenter + parInsulationThickness) * Sin(PI / 4), 0
    dInsulationDia = parStrDiameter + 2 * parInsulationThickness
    Set ObjInsStrainer = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainer
    Set ObjInsStrainer = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
