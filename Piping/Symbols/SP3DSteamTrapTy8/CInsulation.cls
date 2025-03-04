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
'   Creation Date:  Fridday, Feb 3 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model No. BPC32Y (Type: Balanced Pressure Thermostatic Steam Trap)
'    in Spirax Sarco's catalog available at URL http://www.spiraxsarco.com
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
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

    Dim objTrapBody As Object
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsulationDia   As Double
    
    Dim parTrapLength1 As Double
    Dim parFacetoFace As Double
    Dim parBodyWidth As Double
    Dim parTrapHeightToptoCen As Double
    Dim parWithdrawalDistance As Double
    Dim parTrapHeightBottoCen As Double
    Dim parWithdrawalDistance1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parTrapLength1 = arrayOfInputs(2)    'A
    parFacetoFace = arrayOfInputs(3)     'A1
    parBodyWidth = arrayOfInputs(4)  'B
    parTrapHeightToptoCen = arrayOfInputs(5)     'C
    parWithdrawalDistance = arrayOfInputs(6)     'E
    parTrapHeightBottoCen = arrayOfInputs(7)     'F
    parWithdrawalDistance1 = arrayOfInputs(8)    'G
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl

' Insert your code for output 1(Insulation for Upper Trap Body)
'   Assumptions:-
'      1. Upper trap body diameter is taken to be 60 percent of A
'      2. Upper trap body height is taken to be 40 percent of C
    Dim dInsUpperTrapBodyDia  As Double
    Dim dUpperTrapBodyHt  As Double

    dInsUpperTrapBodyDia = 0.6 * parTrapLength1 + 2 * parInsulationThickness
    dUpperTrapBodyHt = 0.4 * parTrapHeightToptoCen
    
    stPoint.Set 0, (parTrapHeightToptoCen - dUpperTrapBodyHt), 0
    enPoint.Set stPoint.x, stPoint.y + dUpperTrapBodyHt + parInsulationThickness, stPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsUpperTrapBodyDia
    
' Insert your code for output 2(Insulation for Trap Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
'   Assumption: The base for Upper Trap Body is taken as a box starting from half-of-pipeDiam
    stPoint.Set -(parTrapLength1 / 2 + parInsulationThickness), _
                -(pipeDiam / 2 + parInsulationThickness), _
                -(parBodyWidth / 2 + parInsulationThickness)
    enPoint.Set (parTrapLength1 / 2 + parInsulationThickness), _
                (parTrapHeightToptoCen - dUpperTrapBodyHt + parInsulationThickness), _
                (parBodyWidth / 2 + parInsulationThickness)
    Set objTrapBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTrapBody
    Set objTrapBody = Nothing
    
' Insert your code for output 3(Insulation for Lower Trap Body)
'   Assumptions:
'     1. Lower trap body diameter is taken to be 30 percent of A
'     2. Start point of the Lower Trap Body is taken to be at A /4 distance from the origin
    Dim dInsLowerTrapBodyDia  As Double
    
    dInsLowerTrapBodyDia = 0.3 * parTrapLength1 + 2 * parInsulationThickness
    stPoint.Set -parTrapLength1 / 4, 0, 0
    enPoint.Set stPoint.x + (parTrapHeightBottoCen + parInsulationThickness) * Cos(PI / 4), _
                stPoint.y - (parTrapHeightBottoCen + parInsulationThickness) * Sin(PI / 4), _
                stPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsLowerTrapBodyDia
    
' Insert your code for output 4(Insulation Port1)
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set -parFacetoFace / 2 + flangeThick + parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 5(Insulation for body Port1 Side)
    Dim dPipeLength As Double
    dPipeLength = (parFacetoFace - parTrapLength1) / 2 - flangeThick
    stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    enPoint.Set stPoint.x + dPipeLength, stPoint.y, stPoint.z
    dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 6(Insulation Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001

    stPoint.Set parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2 - flangeThick - parInsulationThickness, 0, 0
    dInsulationDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness

'   Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    
' Insert your code for output 7(Insulation for body Port2 Side)
    stPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
    enPoint.Set stPoint.x - dPipeLength, stPoint.y, stPoint.z
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
