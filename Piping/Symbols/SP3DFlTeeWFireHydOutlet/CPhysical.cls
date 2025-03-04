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
'   Author:          SSP
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Flanged T with fire hydrant outlet
'   The symbol consists of seven outputs, four physical and three insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim flangeThick1     As Double
    Dim flangeThick2     As Double
    Dim flangeThick3     As Double
    Dim flangeThick4    As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim sptOffset3       As Double
    Dim sptOffset4      As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3           As Double
    Dim depth4          As Double
    
    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace3toCenter = arrayOfInputs(3)
    parFace4toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

' Place Nozzle 1
' Origin of the symbol is intersection of Nozzle axis and vertical axis
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
' Assumption
' nozzle1 and nozzle2 are at equi distant from center.
    
    NozzleLength = 2 * parFace1toCenter - flangeThick2

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    Dim MeetPoint As Double
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam, sptOffset3, depth3
    
    oPlacePoint.Set 0, 0, -parFace3toCenter - sptOffset3 + depth3
    oDir.Set 0, 0, -1
    
    MeetPoint = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam3 / 2) ^ 2)
    
    NozzleLength = parFace3toCenter - MeetPoint

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam, sptOffset4, depth4
    
    oPlacePoint.Set 0, parFace4toCenter + sptOffset4 - depth4, 0
    oDir.Set 0, 1, 0
    
    MeetPoint = Sqr((pipeDiam1 / 2) ^ 2 - (pipeDiam4 / 2) ^ 2)
    
    NozzleLength = parFace4toCenter - MeetPoint

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

