VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No.: S49AC, Page No. D-105 of PDS Piping Component Data Reference Guide.
'   This Symbol is created with Nine Outputs, Five of which are physical aspect outputs.
'    Branch Body, Strainer Body using 'PlaceCylinder' function,
'    Three ObjNozzle object by using 'CreateNozzleWithLength' function  and
'   And all the Insulation aspect ouputs using 'PlaceCylinder' function.
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjBranchBody As Object
    Dim ObjStrainerBody As Object

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
    parFace1toCenter = arrayOfInputs(2)     'p2
    parFace2toCenter = arrayOfInputs(3)     'P3
    parStrToptoCenter = arrayOfInputs(4)    'P4
    parStrDiameter = arrayOfInputs(5)             'P5
    parBranchDiameter = arrayOfInputs(6)             'P6
    parStrWidth = arrayOfInputs(7)                'P7
    parFace3toStrTop = arrayOfInputs(8)        'P8
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

 ' Insert your code for output 1(Branch Body)
     stPoint.Set 0, 0, 0
    enPoint.Set 0, parStrToptoCenter - parStrWidth, 0
    Set ObjBranchBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBranchDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchBody
    Set ObjBranchBody = Nothing

 ' Insert your code for output 2(Strainer Body)
    stPoint.Set 0, parStrToptoCenter - parStrWidth, 0
    enPoint.Set 0, parStrToptoCenter, 0
    Set ObjStrainerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parStrDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainerBody
    Set ObjStrainerBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double

    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0

    dNozzleLength = parFace1toCenter
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    dNozzleLength = parFace2toCenter
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, parStrToptoCenter + parFace3toStrTop + sptOffset - depth, 0
    oDir.Set 0, 1, 0

    dNozzleLength = parFace3toStrTop
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub

