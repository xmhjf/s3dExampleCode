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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Saturday, Feb 17 2001
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC                     DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.Aug.2007     VRG                     TR-124960 Insulation diameter was computed based on End preparation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Integer
    
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedNozzle1 As Object
    Dim ObjInsulatedNozzle2 As Object
    
    Dim parFacetoFace As Double
    Dim parY4StemHeight As Double
    Dim parY5OperatorDiameter As Double
    Dim parY6HandwheelToOpTop As Double
    Dim parY7OperatorHeight As Double
    Dim parY8HandwheelProjection As Double
    Dim parY9HandwheelDiameter As Double
    Dim parY10HandwheelOffset As Double
    Dim parY11HandwheelToStemTop As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parY4StemHeight = arrayOfInputs(3)
    parY5OperatorDiameter = arrayOfInputs(4)
    parY6HandwheelToOpTop = arrayOfInputs(5)
    parY7OperatorHeight = arrayOfInputs(6)
    parY8HandwheelProjection = arrayOfInputs(7)
    parY9HandwheelDiameter = arrayOfInputs(8)
    parY10HandwheelOffset = arrayOfInputs(9)
    parY11HandwheelToStemTop = arrayOfInputs(10)
    parHandwheelAngle = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0

 ' Insert your code for output 8(Insulated Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parInsulationDiameter As Double
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0

    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    
 ' Insert your code for output 11(Insulated Nozzle1)
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
         parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
         parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set ObjInsulatedNozzle1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedNozzle1
    Set ObjInsulatedNozzle1 = Nothing
    
    
    
 ' Insert your code for output 12(Insulated Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2 - (flangeThick + parInsulationThickness), 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
         parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
         parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Set ObjInsulatedNozzle2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedNozzle2
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set ObjInsulatedNozzle2 = Nothing
     
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
