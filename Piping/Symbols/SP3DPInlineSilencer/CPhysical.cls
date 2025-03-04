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
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is:  I201BZ Page No. 364 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim ObjVentSilencerBody  As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parVentSilencerBodyDiameter As Double
    Dim parVentSilencerBodyLength1 As Double
    Dim parVentSilencerBodyLength2 As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2) 'P2
    parFace1toCenter = arrayOfInputs(3) 'P3
    parOffset1 = arrayOfInputs(4) 'P4
    parOffset2 = arrayOfInputs(5) 'P5
    parVentSilencerBodyDiameter = arrayOfInputs(6) 'P6
    parVentSilencerBodyLength1 = arrayOfInputs(7)  'P7
    parVentSilencerBodyLength2 = arrayOfInputs(8)  'P8
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

' Insert your code for output 4(Body of VentSilencer)
    stPoint.Set -parVentSilencerBodyLength1, 0, 0
    enPoint.Set parVentSilencerBodyLength2, 0, 0
    Set ObjVentSilencerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVentSilencerBodyDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVentSilencerBody
    Set ObjVentSilencerBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
 
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    Dim NozzleLength As Double
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, -parOffset1
    oDir.Set -1, 0, 0
    NozzleLength = parFacetoCenter - parVentSilencerBodyLength1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0, parOffset2
    oDir.Set 1, 0, 0
    NozzleLength = parFace1toCenter - parVentSilencerBodyLength2
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub

