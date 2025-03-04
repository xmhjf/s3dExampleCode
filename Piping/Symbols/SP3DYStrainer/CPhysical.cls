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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No.: S51AC or S51AZ, Page No. D-105 PDS Piping Component Data Reference Guide.
'   Symbol is created with 13 Outputs - Eight are Insulation aspect outputs.
'   Other Five are physical aspect outputs created as follows:
'   ObjBodyBranch, ObjBodyStrainer - Using 'PlaceCylinder' function and
'   three ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'   Change History:

'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   26.OCT.2004     MS The input Offset is replaced with ComponentOffset.
'                                       Also the Nozzle 1 position in cphysical is corrected to
'                                       oPlacePoint.Set -parFace1toCenter + depth -sptOffset,0,0
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private PI       As Double

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
    Dim ObjBodyBranch As Object
    Dim ObjBodyStrainer As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parStrToptoCenter As Double
    Dim parStrDiameter As Double
    Dim parBranchDiameter As Double
    Dim parComponentOffset As Double
    Dim parStrWidth As Double
    Dim parFace3toStrTop As Double
    Dim parInsulationThickness As Double

    Dim stPoint                 As New AutoMath.DPosition
    Dim enPoint                 As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)  'P2
    parFace2toCenter = arrayOfInputs(3)  'P3
    parStrToptoCenter = arrayOfInputs(4) 'P4
    parStrDiameter = arrayOfInputs(5)    'P5
    parBranchDiameter = arrayOfInputs(6) 'P6
    parComponentOffset = arrayOfInputs(7)         'P7
    parStrWidth = arrayOfInputs(8)       'P8
    parFace3toStrTop = arrayOfInputs(9)  'P9
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0

' Insert your code for output 9(Body of Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set (parStrToptoCenter - parStrWidth) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth) * Sin(PI / 4), 0
    Set ObjBodyBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBranchDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyBranch
    Set ObjBodyBranch = Nothing

' Insert your code for output 10(Body of Striner)
    stPoint.Set (parStrToptoCenter - parStrWidth) * Cos(PI / 4), _
                (parStrToptoCenter - parStrWidth) * Sin(PI / 4), 0
    enPoint.Set parStrToptoCenter * Cos(PI / 4), parStrToptoCenter * Sin(PI / 4), 0
    Set ObjBodyStrainer = PlaceCylinder(m_OutputColl, stPoint, enPoint, parStrDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyStrainer
    Set ObjBodyStrainer = Nothing
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

    oPlacePoint.Set -parFace1toCenter + depth - sptOffset, 0, 0
    oDir.Set -1, 0, 0

    dNozzleLength = parFace1toCenter
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
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
    oPlacePoint.Set (parStrToptoCenter + parFace3toStrTop + sptOffset - depth) * Cos(PI / 4) - parComponentOffset * Sin(PI / 4), _
                    (parStrToptoCenter + parFace3toStrTop + sptOffset - depth) * Sin(PI / 4) + parComponentOffset * Cos(PI / 4), 0
    oDir.Set Cos(PI / 4), Sin(PI / 4), 0

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
