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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         GL/VKC
'   Creation Date:  Tuesday,August,12 2008
'   Description:    Insulation for Flanged Heater
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation" 'Used for error messages
Private PI       As Double
Private Const INCH = 0.0254

Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parC As Double
    Dim parM As Double
    Dim parE As Double
    Dim parOT As Double
    Dim parOB As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)     'W
    parDepth = arrayOfInputs(3)     'H
    parC = arrayOfInputs(4)
    parM = arrayOfInputs(5)
    parE = arrayOfInputs(6)
    parOT = arrayOfInputs(7)
    parOB = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint  As New AutoMath.DPosition
    Dim oEnPoint  As New AutoMath.DPosition
    
    'Create Insulation for the heater
    Dim Box1 As Object
    oStPoint.Set 0, 0.5 * parDepth + parInsulationThickness, 0.5 * parWidth
    oEnPoint.Set -parC, _
                 -0.5 * parDepth - parInsulationThickness, _
                 -0.5 * parWidth - parInsulationThickness
    Set Box1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), Box1
    Set Box1 = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
        
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

