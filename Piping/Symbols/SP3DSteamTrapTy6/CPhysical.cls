VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
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

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

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
    Dim ObjBody As Object
    Dim ObjConnector As Object
    
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

'   Assumption: Origin of the symbol is taken at mid-point of parFacetoFace

' Insert your code for output 6(Body of Trap)
'   Assumptions:
'    1. The Body of Trap is taken to be box
'    2. The start point of the box is taken at 40% of parPipingSpecialtyHeight
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -(parPipingSpecialtyLength - parTrapLength1), _
                parPipingSpecialtyHeight * 0.4, _
                -parPipingSpecialtyWidth / 2
    enPoint.Set parTrapLength1, parPipingSpecialtyHeight, parPipingSpecialtyWidth / 2
    Set ObjBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Insert your code for output 7(Connector from ports to Trap)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parPipingSpecialtyHeight * 0.4, 0
'   Assumption: Connector diameter is taken as 80% of parPipingSpecialtyWidth
    Dim dConnectorDia As Double
    dConnectorDia = 0.8 * parPipingSpecialtyWidth
    Set ObjConnector = PlaceCylinder(m_OutputColl, stPoint, enPoint, dConnectorDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector
    Set ObjConnector = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
'   Compute nozzle length
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim dNozzleLength As Double
    dNozzleLength = parFacetoFace - flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

