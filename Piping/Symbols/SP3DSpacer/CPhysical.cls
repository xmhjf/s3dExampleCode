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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Monday, Nov 25 2002
'   Description:
'  This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol constructed as per the catalog available at URL http://www.resistoflex.com, all
'   assumptions are based on the catalog provided at the URL location.
'   Physical aspect consist of one  Spacer cylinder and Two nozzles. Insulation aspect consist of
'   a simple cylinder.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------    -----                    ------------------
'  09.Jul.2003    SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006    KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  29.Jan.2007    dkl  CR 112819 Updated the computation of Spacer diameter in case of non-flanged End preparation.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim raisedFaceOrSocketDiam As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

' Insert your code for output 1(Spacer Body)
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    Dim dSpacerDia As Double
    Dim objSpacerBody As Object
    
    RetrieveParametersForThruBoltedEnds 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                   flangeDiam, sptOffset, depth, raisedFaceOrSocketDiam
    'If the end is flanged or threaded i.e. raised face/socket diameter or
    'flange diameter are non zero then, spacer diameter is the maximum of
    'raised Face/Socket Diameter and flange Diameter.
     If CmpDblGreaterthan(raisedFaceOrSocketDiam, LINEAR_TOLERANCE) Or _
                            CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
        dSpacerDia = BodyOD(raisedFaceOrSocketDiam, flangeDiam)
    'Else spacer diameter is equal to pipe diameter (End preparation is not flanged).
    Else
        dSpacerDia = pipeDiam
    End If
    
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEndPoint.Set parFacetoFace / 2, 0, 0
    
    Set objSpacerBody = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, dSpacerDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSpacerBody
    Set objSpacerBody = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    
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
  
 ' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
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
