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
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I201CZ Page No. 364 of PDS Piping Component Data Reference Guide pcd.pdf.
'
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
    Dim ObjHoriCylinIns As Object
    Dim ObjVertCylinIns As Object
    Dim ObjPipeinNoz1Ins As Object
    Dim ObjPipeinNoz2Ins As Object
    Dim ObjNoz1Ins As Object
    Dim ObjNoz2Ins As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parCylinderDiameter As Double
    Dim parVertCylinHeightFromCen As Double
    Dim parHoriCylinLengthFromCen As Double
    Dim parInsulationThickness  As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)     'P2
    parFace1toCenter = arrayOfInputs(3)     'P3
    parCylinderDiameter = arrayOfInputs(4)     'P4
    parVertCylinHeightFromCen = arrayOfInputs(5)      'P5
    parHoriCylinLengthFromCen = arrayOfInputs(6)      'P6
    parInsulationThickness = arrayOfInputs(7)
        
    iOutput = 0

 ' Insert your code for output 3(Horizontal Cylin Insulation)
    Dim parInsulationDiameter As Double
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, 0
    endPoint.Set parHoriCylinLengthFromCen + parInsulationThickness, 0, 0
    parInsulationDiameter = parCylinderDiameter + parInsulationThickness * 2
    Set ObjHoriCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinIns
    Set ObjHoriCylinIns = Nothing
    
 ' Insert your code for output 4(Vertical Cylin Insulation)
    startPoint.Set 0, 0, -parVertCylinHeightFromCen - parInsulationThickness
    endPoint.Set 0, 0, parCylinderDiameter / 2 + parInsulationThickness
    parInsulationDiameter = parCylinderDiameter + parInsulationThickness * 2
    Set ObjVertCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylinIns
    Set ObjVertCylinIns = Nothing
    
 ' Insert your code for output 5(Pipe portion in Noz1, Insulation)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    startPoint.Set 0, 0, -parVertCylinHeightFromCen
    endPoint.Set 0, 0, -parFacetoCenter + flangeThick
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjPipeinNoz1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz1Ins
    Set ObjPipeinNoz1Ins = Nothing
    
 ' Insert your code for output 7(Noz1 Insulation)
    startPoint.Set 0, 0, -parFacetoCenter
    endPoint.Set 0, 0, -parFacetoCenter + flangeThick + parInsulationThickness
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
    Set ObjNoz1Ins = Nothing
 
 ' Insert your code for output 6(Pipe portion in Noz2, Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    startPoint.Set parFace1toCenter - flangeThick, 0, 0
    endPoint.Set parHoriCylinLengthFromCen, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjPipeinNoz2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeinNoz2Ins
    Set ObjPipeinNoz2Ins = Nothing
    
 ' Insert your code for output 8(Noz2 Insulation)
    startPoint.Set parFace1toCenter, 0, 0
    endPoint.Set parFace1toCenter - flangeThick - parInsulationThickness, 0, 0
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
    Set ObjNoz2Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
