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
'    The details of this Symbol are available in PDS Piping Component Data Reference Guide
'    at Page No D-91 and SN=FS26A. Symbol consist of only Physical aspect and made up of
'    Two Nozzles with Length and a Cylinder for Body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjCylinBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBodyToptoCenter As Double
    Dim parBodyBottoCenter As Double
    Dim parBodyDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D1
    parFace2toCenter = arrayOfInputs(3)     'D2
    parBodyToptoCenter = arrayOfInputs(4)       'D3
    parBodyBottoCenter = arrayOfInputs(5)       'D4
    parBodyDiameter = arrayOfInputs(6)          'D5
    
    iOutput = 0

 ' Insert your code for output 1(Cylinder Body)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set 0, 0, -parBodyBottoCenter
    endPt.Set 0, 0, parBodyToptoCenter
    Set ObjCylinBody = PlaceCylinder(m_OutputColl, startPt, endPt, parBodyDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinBody
    Set ObjCylinBody = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    Dim NozzleLength As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, 0, -parFace2toCenter - sptOffset + depth
    oDir.Set 0, 0, -1
    NozzleLength = parFace2toCenter - parBodyBottoCenter
    If NozzleLength < flangeThick Then NozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
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
