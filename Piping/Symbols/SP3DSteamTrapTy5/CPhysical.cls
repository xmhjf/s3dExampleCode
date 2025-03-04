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
'    This symbol constructed as per the Model B Trap sketch in seriesb-bi.pdf (Type: Float & Thermostatic Steam Trap)
'    in Armstrong's catalog available at URL http://www.armstrong-intl.com
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
    Dim ObjTrapBody  As Object
        
    Dim parPipingSpecialtyHeight As Double
    Dim parPipingSpecialtyLength As Double
    Dim parNozzleOffset As Double
    Dim parPipingSpecialtyWidth As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzle1toNozzle2 As Double
    Dim parNozzleOffset2 As Double
    Dim parInsulationThickness As Double
        
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
  
'   Assumption: Origin of the symbol is taken at intersection point of:
'   horizontal line drawn at mid-point of 'M'
'   Vertical line drawn at mid-point of 'C'

' Insert your code for output 6(Body of Trap)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'   Assumption: The geometry of the body of the Trap is approximated to a box, with
'   length (C/2 -  2 * K) on port 1 & Port 2 side and 'C/2' on the opposite side
    stPoint.Set -(parPipingSpecialtyLength / 2 - 2 * parNozzleOffset1), _
                -(parNozzle1toNozzle2 / 2 + parNozzleOffset), _
                -parNozzleOffset2
    enPoint.Set parPipingSpecialtyLength / 2, _
                stPoint.y + parPipingSpecialtyHeight, _
                stPoint.z + parPipingSpecialtyWidth
    
    Set ObjTrapBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTrapBody
    Set ObjTrapBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dFacetoCenter As Double
    Dim dNozzleLength As Double

    dFacetoCenter = parPipingSpecialtyLength / 2 - parNozzleOffset1
    oPlacePoint.Set -dFacetoCenter - sptOffset + depth, parNozzle1toNozzle2 / 2, 0
    oDir.Set -1, 0, 0
    dNozzleLength = parNozzleOffset1
    If CmpDblGreaterthan(flangeThick, dNozzleLength) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dFacetoCenter = parPipingSpecialtyLength / 2
    oPlacePoint.Set -dFacetoCenter - sptOffset + depth, -parNozzle1toNozzle2 / 2, 0
    oDir.Set -1, 0, 0
    dNozzleLength = 2 * parNozzleOffset1
    If CmpDblGreaterthan(flangeThick, dNozzleLength) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

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

