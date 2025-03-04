VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Wednesday, Oct 13, 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Generic Non Radial Olet symbol. It takes inputs from the Generic Data tables.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   18.FEB.2005     MS     TR 71938 The present intent is to only support
'                                           tangential non-radials which means that the offset will always be
'                                           1/2 the OD of the header minus 1/2 the OD of the branch pipe. Hence removed
'                                           ComponentOffset as input to symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "SimplePhysical:" 'Used for error messages
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
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim cptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim cptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'get branch data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, cptOffset2, depth2
    'The offset will always be 1/2 the OD of the header minus 1/2 the OD of the branch pipe
    Dim dComponentOffset As Double
    dComponentOffset = (pipeDiam1 / 2) - (pipeDiam2 / 2)
    
    
    ''Waist Dia is is assumed as 1.5 times the Branch PipeOD as per the PDS Eden 'F-173'
    Dim dWaistDiameter As Double
    dWaistDiameter = 1.5 * pipeDiam2
    
    Dim doletBody As Double
    Dim dOletTopConeLength As Double
    
'    doletBody = (Sqr(((pipeDiam2 / 2) * (pipeDiam2 / 2)) - ((pipeDiam1 / 2) * (pipeDiam1 / 2)))) / 2
    doletBody = (Sqr((pipeDiam1 ^ 2) - (pipeDiam2 ^ 2))) / 2
    dOletTopConeLength = 0.4 * doletBody
    
' Insert your code for output 1(Cone 1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set 0, parFacetoCenter - dOletTopConeLength, dComponentOffset
    enPoint.Set 0, 0, dComponentOffset + 0.0000001
    
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, dWaistDiameter, 0
     
' Insert your code for output 2(Cone 2)
    stPoint.Set 0, parFacetoCenter - dOletTopConeLength, dComponentOffset
    enPoint.Set 0, parFacetoCenter, dComponentOffset + 0.0000001
    
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, dWaistDiameter, pipeDiam2
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, parFacetoCenter + cptOffset2 - depth2, dComponentOffset
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

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

