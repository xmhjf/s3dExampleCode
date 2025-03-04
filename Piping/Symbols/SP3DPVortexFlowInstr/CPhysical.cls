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
'    Symbol Model No. is: I237AZ Page No. 367 of PDS Piping Component Data Reference Guide pcd.pdf.
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
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double

    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjConnector As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parOffset As Double
    Dim parTrapLength As Double
    Dim parTrapHeight As Double
    Dim parTrapWidth1 As Double
    Dim parTrapWidth2 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2) 'P2
    parFace1toCenter = arrayOfInputs(3) 'P3
    parOffset = arrayOfInputs(4) 'P4
    parTrapLength = arrayOfInputs(5) 'P5
    parTrapHeight = arrayOfInputs(6) 'P6
    parTrapWidth1 = arrayOfInputs(7) 'P7
    parTrapWidth2 = arrayOfInputs(8) 'P8
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

' Insert your code for output 7(Body of Trap)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ConnectorDiameter As Double
    
    stPoint.Set -parTrapLength / 2, -parTrapWidth1, parOffset - parTrapHeight
    enPoint.Set parTrapLength / 2, parTrapWidth2, parOffset
    Set ObjBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code for output 8(Connector from ports to Trap)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2
    
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parOffset - parTrapHeight
'   Set ConnectorDiameter equal to 0.2 * minimum of pipeDiam and parTrapLength
    ConnectorDiameter = 0.2 * parTrapLength
    If parTrapLength < pipeDiam1 Then ConnectorDiameter = 0.2 * pipeDiam1
    Set ObjConnector = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConnectorDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector
    Set ObjConnector = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    dNozzleLength = parFacetoCenter
    If dNozzleLength < flangeThick1 Then dNozzleLength = flangeThick1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    dNozzleLength = parFace1toCenter
    If dNozzleLength < flangeThick2 Then dNozzleLength = flangeThick2
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

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

