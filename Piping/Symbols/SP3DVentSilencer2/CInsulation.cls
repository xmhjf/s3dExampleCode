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
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model HPS(with size in inches as suffix) in the hps.pdf
'    Fluidkinetics's catalog available at URL http://www.fluidkinetics.com/hushpup.htm
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
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
    
    Dim iOutput             As Double

    Dim dInsulationDia   As Double

    Dim parVentSilencerBodyDiameter As Double
    Dim parPipingSpecialtyHeight As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVentSilencerBodyDiameter = arrayOfInputs(2) 'B
    parPipingSpecialtyHeight = arrayOfInputs(3)    'C
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl

'   Insert your code for output 1(Insulation for Silencer Body)
'   Assumptions:
'    1. The height Body of Vent Silencer is taken to be 80% of 'C' and assumed to be cylinder
'    2. The dimension of cylinder takes care of the curved portion of the body at bottom
'       (the curved portion is not shown)
'
    Dim dVentSilencerHt As Double
    Dim dNozzleLength As Double
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    dVentSilencerHt = 0.8 * parPipingSpecialtyHeight

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'   Compute nozzle length
    dNozzleLength = parPipingSpecialtyHeight - dVentSilencerHt
    If CmpDblGreaterthan(flangeThick, dNozzleLength) Then dNozzleLength = flangeThick * 1.1 '1.1 is used to avoid failure
    stPoint.Set dNozzleLength - parInsulationThickness, 0, 0
    enPoint.Set parPipingSpecialtyHeight + parInsulationThickness, 0, 0
    dInsulationDia = parVentSilencerBodyDiameter + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia

' Insert your code for output 2(Insulation for Port 1)
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set 0, 0, 0
    enPoint.Set flangeThick + parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
'   Insert your code for output 3(Insulation of connector from port1 to Silencer)
    stPoint.Set flangeThick + parInsulationThickness, 0, 0
    enPoint.Set dNozzleLength, 0, 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
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
