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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is prepared based on Resistoflex Flanged Lined Pipe & Fittings catalog.
'   The symbol consists of Thirteen outputs, Eight physical and  Five  Insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim flangeThick1     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjOriFlange1Body As Object
    Dim ObjOriFlange2Body As Object
    Dim ObjOriFlange1MatingFlange As Object
    Dim ObjOriFlange2MatingFlange As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parOrificeFlangeThk As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parOrificeFlangeThk = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
' The origin is taken to be along the flow axis and in between the orificeflanges.
'The Port 3 and Port4 are at parOrificeFlangeThk/2 from the origin.
'The Orifice flange Thickness(parOrificeFlangeThk) is taken to be same for both the Orifice flanges.
'The gap between the orifice flanges is taken to be 2mm.
 ' Insert your code for output 1(Orifice Flange1 Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset, depth
'The Orifice Flange Diameter is taken to be 0.7 times the flange diameter of Port1.
    Dim OrificeFlangeBodyDia As Double
    OrificeFlangeBodyDia = 0.7 * flangeDiam
    stPoint.Set -0.001, 0, 0
    enPoint.Set -0.001 - parOrificeFlangeThk, 0, 0
    Set ObjOriFlange1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, OrificeFlangeBodyDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriFlange1Body
    Set ObjOriFlange1Body = Nothing
    
 ' Insert your code for output 2(Orifice Flange2 Body)
     stPoint.Set 0.001, 0, 0
    enPoint.Set 0.001 + parOrificeFlangeThk, 0, 0
    Set ObjOriFlange2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, OrificeFlangeBodyDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriFlange2Body
    Set ObjOriFlange2Body = Nothing
    
 ' Insert your code for output 3(Flange1 Body)
 'The Flange1 Body Diameter is taken to be Flange Diameter of Port1
 'The Flange1 Body Thickness is taken to be Flange Thickness of Port1
 'The gap between the orifice flange and the mating flange is taken to be 2mm
    stPoint.Set -0.003 - parOrificeFlangeThk, 0, 0
    enPoint.Set -0.003 - parOrificeFlangeThk - flangeThick1, 0, 0
    Set ObjOriFlange1MatingFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriFlange1MatingFlange
    Set ObjOriFlange1MatingFlange = Nothing
    
 ' Insert your code for output 4(Flange2 Body)
'The Flange2 Body Diameter is taken to be Flange Diameter of Port2
 'The Flange2 Body Thickness is taken to be Flange Thickness of Port2
 'The gap between the orifice flange and the mating flange is taken to be 2mm
    stPoint.Set 0.003 + parOrificeFlangeThk, 0, 0
    enPoint.Set 0.003 + parOrificeFlangeThk + flangeThick1, 0, 0
    Set ObjOriFlange2MatingFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOriFlange2MatingFlange
    Set ObjOriFlange2MatingFlange = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1 (Nozzle with Length)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim NozzleLength As Double
    NozzleLength = parFace1toCenter - parOrificeFlangeThk - flangeThick1
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2 (Nozzle with Length)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    NozzleLength = parFace2toCenter - parOrificeFlangeThk - flangeThick1
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'Port3 and Port4 are taken to be 90 deg apart(45 deg from the vertical axis each)
' Place Nozzle 3 (Nozzle with Length)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim meetRadius As Double
    meetRadius = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam / 2) ^ 2)
    NozzleLength = parFace3toCenter - meetRadius
    oPlacePoint.Set parOrificeFlangeThk / 2, _
                                (parFace3toCenter + sptOffset - depth) * Cos(PI / 4), _
                                -(parFace3toCenter + sptOffset - depth) * Sin(PI / 4)
    oDir.Set 0, Cos(PI / 4), -Sin(PI / 4)
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 4 (Nozzle with Length)
'The Face3 to center is assumed to be equal to Face4 to center
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    meetRadius = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam / 2) ^ 2)
    oPlacePoint.Set -parOrificeFlangeThk / 2, _
                                (parFace3toCenter + sptOffset - depth) * Cos(PI / 4), _
                                (parFace3toCenter + sptOffset - depth) * Sin(PI / 4)
    oDir.Set 0, Cos(PI / 4), Sin(PI / 4)
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub



