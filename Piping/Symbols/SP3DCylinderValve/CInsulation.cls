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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Friday, May 4 2007
'   Description:
'   This symbol is created for cylinder valves of JIS B 8245 & 8246 as per CR-118765
'   Change History:
'   dd.mmm.yyyy     who                               change description
'   -----------         -----                                ------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim sptOffset2       As Double
    Dim depth2           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double

    
    Dim iOutput     As Double

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
 ' Insert your code for output 1(Side cap Cylinder)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth
    
    Dim dInsulationDia As Double
    dInsulationDia = pipeDiam2 + 2 * parInsulationThickness

'Assumptions:
'    1.The axis of the side cylinder is assumed to be at an offset of half the pipe OD at Port 2
'    2.The diameter of the side cap cylinder is assumed to be equal to Pipe OD at Port 2
'    3.The end of the side cap cylinder is assumed to be at distance equal Face 2 to Center from the center
    
    oStPoint.Set -pipeDiam2 / 2, 0, 0
    oEnPoint.Set -pipeDiam2 / 2, -parFace2toCenter, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia
    
' Insert your code for output 2(Insulation Cylinder along Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
    If CmpDblGreaterthan(dInsulationDia, 2 * parFace2toCenter) Then
        dInsulationDia = 2 * parFace2toCenter
    End If

'Assumption:
'The length of the cylinder along port 1 above the center line is assumed to be equal to Face 1 to Center
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set parFace1toCenter, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia
    
' Insert your code for output 3(Insulation Cylinder along Port 2)
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
    
        dInsulationDia = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
    End If
    
    oStPoint.Set 0, parFace2toCenter, 0
    oEnPoint.Set 0, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
       
End Sub
