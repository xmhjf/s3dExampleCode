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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG 
'   Creation Date:  Monday, Nov 25 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the catalog available at URL http://www.resistoflex.com, all
'   assumptions are based on the catalog provided at the URL location.
'   Physical aspect consist of one Distant Spacer cylinder and Two nozzles. Insulation aspect consist of
'   a simple cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who                    change description
'   -----------     -----                  ------------------
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is added.
'   14.Mar.2005     svsmylav       TR-49094: Added condition to ensure flange body uses
'                                          maximum of raisedFaceOrSocketDiam, flangeDiam.
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim raisedFaceOrSocketDiam1 As Double
    Dim raisedFaceOrSocketDiam2 As Double
    Dim CollarTk As Double
    
    Dim iOutput     As Double
    Dim ObjFlange1 As Object
    Dim ObjDistBody As Object
    Dim ObjFlange2 As Object
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0
    
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition

' Insert your code for output 1(Flange Body)
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                                                flangeDiam, sptOffset1, depth1, raisedFaceOrSocketDiam1
    'Raised face immediately after the Nozzle (Collar) is assumed to be 0.75 inches in thickness as told by NN
    CollarTk = 0.01905
    startPoint.Set -parFacetoFace / 2, 0, 0
    endPoint.Set -parFacetoFace / 2 + CollarTk, 0, 0
    'Flange body diameter is the maximum of raised Face Diameter and flange Diameter
    Dim dFlangeBodyDiameter As Double
    dFlangeBodyDiameter = BodyOD(raisedFaceOrSocketDiam1, flangeDiam)
    Set ObjFlange1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, dFlangeBodyDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange1
    Set ObjFlange1 = Nothing
    
' Insert your code for output 1(Distance Piece Body)
    startPoint.Set -parFacetoFace / 2 + CollarTk, 0, 0
    endPoint.Set parFacetoFace / 2 - CollarTk, 0, 0
    Set ObjDistBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, pipeDiam1, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDistBody
    Set ObjDistBody = Nothing

' Insert your code for output 1(Flange 2)
    RetrieveParametersForThruBoltedEnds 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, _
                                                                flangeDiam, sptOffset2, depth2, raisedFaceOrSocketDiam2
    startPoint.Set parFacetoFace / 2, 0, 0
    endPoint.Set parFacetoFace / 2 - CollarTk, 0, 0
    dFlangeBodyDiameter = BodyOD(raisedFaceOrSocketDiam2, flangeDiam)
    Set ObjFlange2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, dFlangeBodyDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange2
    Set ObjFlange2 = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
  
 ' Place Nozzle 2
    oPlacePoint.Set parFacetoFace / 2 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
