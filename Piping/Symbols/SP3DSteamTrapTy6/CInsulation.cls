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
'   Creation Date:  Tuesday, Feb 6 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model UNA13/15 V in the una-e.pdf and DBL_3973e-995_00_UNA13-UNA15_en.pdf
'    (Type: Ball-Float Trap) in Flowserve Gestra's catalog available at URL http://www.gestra.com/scripts/frames.php
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjInsTrapBody  As Object

    Dim dInsulationDia   As Double

    Dim parFacetoFace As Double
    Dim parPipingSpecialtyLength As Double
    Dim parTrapLength1 As Double
    Dim parPipingSpecialtyHeight As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)    'L
    parPipingSpecialtyLength = arrayOfInputs(3)
    parTrapLength1 = arrayOfInputs(4)
    parPipingSpecialtyHeight = arrayOfInputs(5)
    parPipingSpecialtyWidth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl

' Insert your code for output 1(Insulation for Trap  Body)
'    1. The Body of Trap is taken to be box
'    2. The start point of the box is taken at 40% of parPipingSpecialtyHeight
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -(parPipingSpecialtyLength - parTrapLength1) - parInsulationThickness, _
                parPipingSpecialtyHeight * 0.4 - parInsulationThickness, _
                -parPipingSpecialtyWidth / 2 - parInsulationThickness
    enPoint.Set parTrapLength1 + parInsulationThickness, _
                parPipingSpecialtyHeight + parInsulationThickness, _
                parPipingSpecialtyWidth / 2 + parInsulationThickness
    Set ObjInsTrapBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTrapBody
    Set ObjInsTrapBody = Nothing
    
' Insert your code for output 2(Insulation for Connector  Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parPipingSpecialtyHeight * 0.4 - parInsulationThickness, 0
'   Assumption: Connector diameter is taken as 80% of parPipingSpecialtyWidth
    Dim dConnectorDia As Double
    dConnectorDia = 0.8 * parPipingSpecialtyWidth
    
'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
            dConnectorDia + 2 * parInsulationThickness

' Insert your code for output 3(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 4(Insulation for pipe body)
    stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    enPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 5(Insulation Port2 Side)
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    stPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

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
