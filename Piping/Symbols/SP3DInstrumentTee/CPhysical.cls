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
'   Author:          BG
'   Creation Date:  Friday, Nov 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the documents available at URL http://www.resistoflex.com.
'   Physical aspect consist of Tee Body as Cylinder, Two Nozzles and One Nozzle with Length along
'   Y- Axis.
'   Insulation aspect consist of Three Cylinders, one for the Main Body seceond for the Branch Nozzle
'   and the third for the Branch Body
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjTeeBody As Object
    
    Dim parFacetoFace As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    Dim InstrumentDiameter As Double
    Dim raisedFaceOrSocketDiam As Double
        
    iOutput = 0
 ' Insert your code for output 1(Instrument Body Between 2 Nozzles)
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, _
                                                                            sptOffset, depth, raisedFaceOrSocketDiam
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    'The assuption is that for end preparations of 100 series Flange Diameter will be zero and we will get
    'a positive value for raised Face Diameter, Symbol created under that condition is Instrument Tee
    'For Lugged body Instrument the body diameter is equal to Flange Diameter, so when we get some positive
    'Value for Flange, raised face diameter will be zero.
    If flangeDiam = 0 Then                  'Instrument Tee
        InstrumentDiameter = raisedFaceOrSocketDiam
    Else                                                    'Lugged Body Instrument
        InstrumentDiameter = flangeDiam
    End If
    
    Set ObjTeeBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, InstrumentDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTeeBody
    Set ObjTeeBody = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
 'Place Nozzle 3 with Length
    Dim NozzleLength As Double
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    NozzleLength = parFace3toCenter - Sqr( _
                                                                    (pipeDiam1 / 2) ^ 2 - (pipeDiam / 2) ^ 2 _
                                                                    )
    oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                                oPlacePoint, NozzleLength)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

