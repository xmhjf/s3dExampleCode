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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Friday, Jun 7 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details can be get from PDS - Piping Component Data Reference Guide. Symbol Number
'   SN=FS24A page no. D-91
'   This symbol consist of Physical aspect only. One Box and a Nozzle with Length
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   17.Jan.2005     svsmylav        TR-69697: Added following check to avoid CreateNozzleWithLength function failure:
'                                           If NozzleLength < flangeThick Then NozzleLength = flangeThick
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
    Dim ObjBox As Object
    
    Dim parFacetoCenter As Double
    Dim parPortVOffset As Double
    Dim parPortHOffset As Double
    Dim parStationLength As Double
    Dim parStationHeight As Double
    Dim parStationWidth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'D1
    parPortVOffset = arrayOfInputs(3)   'D2
    parPortHOffset = arrayOfInputs(4)   'D3
    parStationLength = arrayOfInputs(5) 'D4
    parStationHeight = arrayOfInputs(6) 'D5
    parStationWidth = arrayOfInputs(7)  'D6
    
    iOutput = 0

 ' Insert your code for output 1(Box)
    Dim lowerCornerPt As AutoMath.DPosition
    Dim upperCornerPt As AutoMath.DPosition
    
    Set lowerCornerPt = New AutoMath.DPosition
    Set upperCornerPt = New AutoMath.DPosition
    
    lowerCornerPt.Set -parStationLength / 2, -parStationWidth / 2, -parStationHeight / 2
    upperCornerPt.Set parStationLength / 2, parStationWidth / 2, parStationHeight / 2
    Set ObjBox = PlaceBox(m_OutputColl, lowerCornerPt, upperCornerPt)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    Set lowerCornerPt = Nothing
    Set upperCornerPt = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    NozzleLength = (parFacetoCenter - parStationLength / 2)
    If NozzleLength < flangeThick Then NozzleLength = flangeThick
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, parPortHOffset, parPortVOffset
    oDir.Set -1, 0, 0
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


