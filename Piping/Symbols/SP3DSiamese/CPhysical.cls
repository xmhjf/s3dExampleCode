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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details are taken from PDS Piping Component Data Reference Manual, at Page no D-94
'   and SN=FS60A. This Symbol consist of only Physical aspect and made up of a Nozzle with Length,
'   Body Cylinder, and Two Outlets Cylinders
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjBodyCylinder As Object
    Dim ObjOutletLet1 As Object
    Dim ObjOutletLet2 As Object
    
    Dim parFacetoCenter As Double
    Dim parBodyDiameter As Double
    Dim parOutletLength As Double
    Dim parOutletDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'D1
    parBodyDiameter = arrayOfInputs(3)  'D2
    parOutletLength = arrayOfInputs(4)  'D3
    parOutletDiameter = arrayOfInputs(5)    'D4
    
    iOutput = 0

 ' Insert your code for output 1(BodyCylinder)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    startPt.Set 0, 0, -pipeDiam / 2
    endPt.Set 0, 0, pipeDiam / 2
    Set ObjBodyCylinder = PlaceCylinder(m_OutputColl, startPt, endPt, parBodyDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinder
    Set ObjBodyCylinder = Nothing
    
 ' Insert your code for output 2(OutletLet1)
    startPt.Set 0, 0, 0
    endPt.Set parOutletLength * Cos(PI / 4), -parOutletLength * Sin(PI / 4), 0
    Set ObjOutletLet1 = PlaceCylinder(m_OutputColl, startPt, endPt, parOutletDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOutletLet1
    Set ObjOutletLet1 = Nothing
    
 ' Insert your code for output 3(OutletLet2)
    startPt.Set 0, 0, 0
    endPt.Set parOutletLength * Cos(PI / 4), parOutletLength * Sin(PI / 4), 0
    Set ObjOutletLet2 = PlaceCylinder(m_OutputColl, startPt, endPt, parOutletDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOutletLet2
    Set ObjOutletLet2 = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFacetoCenter)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

