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
'   Author:          NN
'   Creation Date:  Sunday, May 12 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
  
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace2toCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    
    iOutput = 0

 
'THE Body of the Cross is part of the Port1

    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim FaceToCenter  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - cptoffset + depth, 0, 0
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    Dim Port1Length As Double
    Port1Length = parFacetoCenter * 2 - flangeThick
    oDir.Set -1, 0, 0
   
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
  
' Place Nozzle 2
  
    oPlacePoint.Set parFacetoCenter + cptoffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
 
    ' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    Dim Port3Length As Double
    Port3Length = parFace2toCenter * 2 - flangeThick
    oPlacePoint.Set 0, parFace2toCenter + cptoffset - depth, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port3Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    ' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    oPlacePoint.Set 0, -(parFace2toCenter + cptoffset - depth), 0
    oDir.Set 0, -1, 0

    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

