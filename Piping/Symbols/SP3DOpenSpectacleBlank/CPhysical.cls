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
'   Creation Date:  Thursday, Jan 25 2001
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjInlineOpenSpect As Object
    Dim ObjInterSpectCyl As Object
    Dim ObjBlankDisc As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parFig8InterAxis As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFig8InterAxis = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0

 ' Insert your code for output 1(In line Open Spectacle)
    Dim parDiscDiameter As Double
    parDiscDiameter = parFig8InterAxis * 0.8
    Dim Port1Position   As New AutoMath.DPosition
    Port1Position.Set -parFacetoFace / 2, 0, 0
    Dim Port2Position   As New AutoMath.DPosition
    Port2Position.Set parFacetoFace / 2, 0, 0

    Set ObjInlineOpenSpect = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parDiscDiameter, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineOpenSpect
    Set ObjInlineOpenSpect = Nothing
    
 ' Insert your code for output 2(Cylinder bet Discs)
    Dim parInterDiscsDiameter As Double
    parInterDiscsDiameter = parFacetoFace / 2
    Dim CylStPosition   As New AutoMath.DPosition
    CylStPosition.Set 0, parFig8InterAxis * 0.3, 0
    Dim CylEndPosition   As New AutoMath.DPosition
    CylEndPosition.Set 0, parFig8InterAxis * 0.7, 0
    Set ObjInterSpectCyl = PlaceCylinder(m_OutputColl, CylStPosition, CylEndPosition, parInterDiscsDiameter, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInterSpectCyl
    Set ObjInterSpectCyl = Nothing
    
 ' Insert your code for output 3(Blank Disc)
    Dim BlankStPosition   As New AutoMath.DPosition
    BlankStPosition.Set -parFacetoFace / 2, parFig8InterAxis, 0
    Dim BlankEndPosition   As New AutoMath.DPosition
    BlankEndPosition.Set parFacetoFace / 2, parFig8InterAxis, 0

    Set ObjBlankDisc = PlaceCylinder(m_OutputColl, BlankStPosition, BlankEndPosition, parDiscDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBlankDisc
    Set ObjBlankDisc = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim FacetoFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    FacetoFace = arrayOfInputs(2)
    oPlacePoint.Set -FacetoFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set FacetoFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim cptoffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - cptoffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + cptoffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
