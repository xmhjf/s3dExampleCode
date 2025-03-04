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
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbols details are taken from PDS Piping Component Data Reference Manual, at Page no D-94
'    and SN=FS66A. This Symbol consist of only Physical aspect and made up of a Nozzle with Length,
'    Body Cylinder, and Two Outlets Cylinders
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
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
    Dim ObjOutlet1 As Object
    Dim ObjOutlet2 As Object
    
    Dim parFacetoCenter As Double
    Dim parCentertoEnd As Double
    Dim parOutletLength As Double
    Dim parOutletDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)     'D1
    parCentertoEnd = arrayOfInputs(3)             'D2
    parOutletLength = arrayOfInputs(4)          'D3
    parOutletDiameter = arrayOfInputs(5)        'D4
    
    iOutput = 0
    
 ' Insert your code for output 1(Outlet1)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    startPt.Set 0, 0, 0
    endPt.Set parOutletLength * Cos(PI / 4), parOutletLength * Sin(PI / 4), 0
    Set ObjOutlet1 = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOutlet1
    Set ObjOutlet1 = Nothing
    
 ' Insert your code for output 2(Outlet2)
    startPt.Set 0, 0, 0
    endPt.Set parOutletLength * Cos(PI / 4), -parOutletLength * Sin(PI / 4), 0
    Set ObjOutlet2 = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOutlet2
    Set ObjOutlet2 = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim NozzleLength As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set 0, 0, -parFacetoCenter - sptOffset + depth
    oDir.Set 0, 0, -1
    NozzleLength = parFacetoCenter + parCentertoEnd
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
