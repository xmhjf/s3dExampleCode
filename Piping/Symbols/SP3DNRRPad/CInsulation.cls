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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MS
'   Creation Date:  Thursday, 14th Oct 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   18.FEB.2005     MS     TR 71938 The present intent is to only support
'                                           tangential non-radials which means that the offset will always be
'                                           1/2 the OD of the header minus 1/2 the OD of the branch pipe. Hence removed
'                                           ComponentOffset as input to symbol.
'   23.FEB.2005     MS     The branch-port is now located at the center of the header with the required offset to be non-radial
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private RAD          As Double

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    RAD = 180 / PI
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjPad As Object
    Dim ObjInsulatedPad As Object
    
    Dim parRPadWidth As Double
    Dim parRPadThickness As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parRPadWidth = arrayOfInputs(2)
    parRPadThickness = arrayOfInputs(3)
    parAngle = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 ' Insert your code for output 2(Insulated Reinforcing Pad)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim pipeDiam2        As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset, depth
    Dim parInsulationDiameter As Double
    Dim ang, angle, pr, sr, w, tana, sina, cosa, obliquebranch, XHeader, Ystart, Yend, wsr As Double
    pr = pipeDiam / 2
    sr = pipeDiam2 / 2
    
    'The offset will always be 1/2 the OD of the header minus 1/2 the OD of the branch pipe
    Dim dComponentOffset As Double
    dComponentOffset = pr - sr
    
    angle = Abs(parAngle)
    ang = angle
    tana = Tan(ang)
    sina = Sin(ang)
    cosa = Cos(ang)
    w = parRPadWidth
    obliquebranch = sr / sina
    parInsulationDiameter = 2 * (parRPadWidth + obliquebranch + parInsulationThickness)
    XHeader = pr / tana
    wsr = w + sr
    If wsr > pr Then
    Ystart = 0
    Else: Ystart = Sqr(pr * pr - wsr * wsr)
    End If
    Yend = pr + parRPadThickness + parInsulationThickness
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set XHeader, Ystart, dComponentOffset
    enPoint.Set XHeader, Yend, dComponentOffset
    Set ObjInsulatedPad = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPad
    Set ObjInsulatedPad = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
