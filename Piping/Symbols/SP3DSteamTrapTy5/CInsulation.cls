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
'    This symbol constructed as per the Model B Trap sketch in seriesb-bi.pdf (Type: Float & Thermostatic Steam Trap)
'    in Armstrong's catalog available at URL http://www.armstrong-intl.com
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjInsBody   As Object
    
    Dim parPipingSpecialtyHeight As Double
    Dim parPipingSpecialtyLength As Double
    Dim parNozzleOffset As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzle1toNozzle2 As Double
    Dim parNozzleOffset2 As Double
    Dim parInsulationThickness As Double
    Dim dInsulationDia As Double
        
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPipingSpecialtyHeight = arrayOfInputs(2)  'B
    parPipingSpecialtyLength = arrayOfInputs(3)  'C
    parNozzleOffset = arrayOfInputs(4)   'D
    parPipingSpecialtyWidth = arrayOfInputs(5)   'H
    parNozzleOffset1 = arrayOfInputs(6)  'K
    parNozzle1toNozzle2 = arrayOfInputs(7)   'M
    parNozzleOffset2 = arrayOfInputs(8)  'N
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl

' Insert your code for output 1(Insulation for Trap  Body)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'   Assumption: The geometry of the body of the Trap is approximated to a box, with
'   length (C/2 -  2 * K) on port 1 & Port 2 side and 'C/2' on the opposite side
    stPoint.Set -(parPipingSpecialtyLength / 2 - 2 * parNozzleOffset1) - parInsulationThickness, _
                -(parNozzle1toNozzle2 / 2 + parNozzleOffset) - parInsulationThickness, _
                -parNozzleOffset2 - parInsulationThickness
    enPoint.Set parPipingSpecialtyLength / 2 + parInsulationThickness, _
                stPoint.y + parPipingSpecialtyHeight + 2 * parInsulationThickness, _
                stPoint.z + parPipingSpecialtyWidth + 2 * parInsulationThickness
    
    Set ObjInsBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody
    Set ObjInsBody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    Dim dFacetoCenter As Double
    Dim dNozzleLength As Double

    dFacetoCenter = parPipingSpecialtyLength / 2 - parNozzleOffset1
    dNozzleLength = parNozzleOffset1

    stPoint.Set -dFacetoCenter, parNozzle1toNozzle2 / 2, 0
    enPoint.Set stPoint.x + flangeThick + parInsulationThickness, stPoint.y, stPoint.z
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 3(Insulation for pipe Port1 Side)
    stPoint.Set -dFacetoCenter + flangeThick, _
                parNozzle1toNozzle2 / 2, 0
    enPoint.Set -dFacetoCenter + dNozzleLength, _
                parNozzle1toNozzle2 / 2, 0
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 4(Insulation Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    dFacetoCenter = parPipingSpecialtyLength / 2
    stPoint.Set -dFacetoCenter, -parNozzle1toNozzle2 / 2, 0
    enPoint.Set stPoint.x + flangeThick + parInsulationThickness, stPoint.y, stPoint.z
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 5(Insulation for pipe Port2 Side)
    dNozzleLength = 2 * parNozzleOffset1
    stPoint.Set -dFacetoCenter + flangeThick, -parNozzle1toNozzle2 / 2, 0
    enPoint.Set -dFacetoCenter + dNozzleLength, _
                stPoint.y, stPoint.z
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
