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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday 16, Feb 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model No. 59 in Royal Guard's
'    catalog available at URL http://www.royalguardstrainers.com/fabricated/simplex/tee.htm
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Dim m_oSymGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
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
    
    Dim dStrToptoCenter As Double
    Dim dStrDiameter As Double
    Dim dBranchDiameter As Double
    Dim dComponentOffset As Double
    Dim dStrWidth As Double
    Dim dFace3toStrTop As Double
    
    Dim parFaceToFace  As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFaceToFace = arrayOfInputs(2)  'A
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
'   Assumption: Origin of the symbol is taken at mid-point of parFacetoFace
  
'   Assumptions: Following dimensions are taken in proportion of known dimensions
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dStrToptoCenter = 0.7 * parFaceToFace  '70% of 'A'
    dStrDiameter = 0.9 * flangeDiam '90% of flangeDiam
    If CmpDblGreaterthan(pipeDiam, dStrDiameter) Then dStrDiameter = pipeDiam
    dBranchDiameter = pipeDiam  'pipeDiam
    dComponentOffset = 0.1 * pipeDiam '10% of pipeDiam
    dStrWidth = 2 * flangeThick '2 * flangeThick
    If CmpDblEqual(dStrWidth, 0) Then dStrWidth = NEGLIGIBLE_THICKNESS
    dFace3toStrTop = NEGLIGIBLE_THICKNESS
    
' Insert your code for output 7(Body of Branch)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set 0, (dStrToptoCenter - dStrWidth), 0

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dBranchDiameter

' Insert your code for output 8(Body of Striner)
    stPoint.Set 0, (dStrToptoCenter - dStrWidth), 0
    enPoint.Set 0, dStrToptoCenter, 0

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dStrDiameter
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

    oPlacePoint.Set -parFaceToFace / 2 + depth - sptOffset, 0, 0
    oDir.Set -1, 0, 0

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dNozzleLength = parFaceToFace - flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    oPlacePoint.Set parFaceToFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set dComponentOffset, (dStrToptoCenter + dFace3toStrTop + sptOffset - depth), 0
    oDir.Set 0, 1, 0

    dNozzleLength = dFace3toStrTop
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick

    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

