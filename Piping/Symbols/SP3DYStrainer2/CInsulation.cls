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
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday 9, Feb 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model No. 57FB-BVY-50 (FBIO) in Royal Guard's
'    catalog available at URL http://www.royalguardstrainers.com/fabricated/simplex/wye.htm
'
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double
Dim m_oSymGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
      PI = 4 * Atn(1)
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

'   Assumption: Origin of the symbol is taken at 35% of parFacetoFace value from port 1
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double

    dFace1toCenter = 0.35 * parFaceToFace
    dFace2toCenter = parFaceToFace - dFace1toCenter
    
'   Assumptions: Following dimensions are taken in proportion of known dimensions
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dStrToptoCenter = 0.9 * parFaceToFace  '90% of 'A'
    dStrDiameter = 0.9 * flangeDiam '90% of flangeDiam
    If CmpDblGreaterthan(pipeDiam, dStrDiameter) Then dStrDiameter = pipeDiam
    dBranchDiameter = pipeDiam  'pipeDiam
    dComponentOffset = 0.1 * pipeDiam '10% of pipeDiam
    dStrWidth = 2 * flangeThick '2 * flangeThick
    If CmpDblEqual(dStrWidth, 0) Then dStrWidth = NEGLIGIBLE_THICKNESS
    dFace3toStrTop = NEGLIGIBLE_THICKNESS

' Insert your code for output 1(Insulation for Port1)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    Dim dInsulationDia  As Double
    
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0

    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 2( Insulation for Body)
    stPoint.Set -dFace1toCenter + flangeThick, 0, 0

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    enPoint.Set dFace2toCenter - flangeThick, 0, 0
    dInsulationDia = pipeDiam + parInsulationThickness * 2

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 3(Insulation for Port2)
    stPoint.Set dFace2toCenter, 0, 0
    enPoint.Set dFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 4(Insulation for Port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set (dStrToptoCenter + dFace3toStrTop) * Cos(PI / 4) - dComponentOffset * Sin(PI / 4), _
                    (dStrToptoCenter + dFace3toStrTop) * Sin(PI / 4) + dComponentOffset * Cos(PI / 4), 0
    Dim Port3CPoint   As New AutoMath.DPosition
    Port3CPoint.Set stPoint.x, stPoint.y, 0
    enPoint.Set Port3CPoint.x - (flangeThick + parInsulationThickness) * Cos(PI / 4), _
                Port3CPoint.y - (flangeThick + parInsulationThickness) * Sin(PI / 4), 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 6(Insulation for Branch)
    stPoint.Set 0, 0, 0
    enPoint.Set (dStrToptoCenter - dStrWidth) * Cos(PI / 4), _
                (dStrToptoCenter - dStrWidth) * Sin(PI / 4), 0
    dInsulationDia = dBranchDiameter + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 8(Insulation for Strainer)
    stPoint.Set (dStrToptoCenter - dStrWidth - parInsulationThickness) * Cos(PI / 4), _
                (dStrToptoCenter - dStrWidth - parInsulationThickness) * Sin(PI / 4), 0
    enPoint.Set (dStrToptoCenter + parInsulationThickness) * Cos(PI / 4), _
                (dStrToptoCenter + parInsulationThickness) * Sin(PI / 4), 0
    dInsulationDia = dStrDiameter + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
