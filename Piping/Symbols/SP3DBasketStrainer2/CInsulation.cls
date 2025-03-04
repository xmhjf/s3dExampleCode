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
'   Creation Date:  Tuesday 31, Jan 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.strainwell.com/conical-stariners.html
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

    Dim iOutput     As Double

    Dim parStrDiameter As Double
    Dim parEndDiameter As Double
    Dim parBasketEndDiameter As Double
    Dim parBasketLength As Double
    Dim parCollarThickness As Double
    Dim parArmLength As Double
    Dim parStrWidth As Double
    Dim parCylinderLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStrDiameter = arrayOfInputs(2) 'D1
    parEndDiameter = arrayOfInputs(3) 'D2
    parBasketEndDiameter = arrayOfInputs(4) 'D3
    parBasketLength = arrayOfInputs(5)      'L
    parCollarThickness = arrayOfInputs(6)   'T
    parArmLength = arrayOfInputs(7)
    parStrWidth = arrayOfInputs(8)
    parCylinderLength = arrayOfInputs(9) 'Body at the small end of the basket
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl

 ' Insert your code for output 1(CollarIns)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim dInsulationDia As Double
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    stPoint.Set -parCollarThickness / 2, 0, 0
    enPoint.Set parCollarThickness / 2, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dInsulationDia
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
