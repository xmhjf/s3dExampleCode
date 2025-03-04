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
'   Creation Date:  Wednesday, Jan 24 2001
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

    Dim iOutput          As Double
    Dim ObjBodyBall      As Object
    Dim ObjBodyCone1     As Object
    Dim ObjBodyCone2     As Object
    Dim ObjStem          As Object
    Dim ObjHandWheel     As Object
    Dim ObjHWSpokesX     As Object
    Dim ObjHWSpokesZ     As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parOperatorHeight = arrayOfInputs(3)
    parOperatorDiameter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 ' Insert your code for output 1(Ball of Body of Globe Valve)
    Set ObjBodyBall = PlaceSphere(m_OutputColl, CenterPos, parFacetoFace / 4)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyBall
    Set ObjBodyBall = Nothing
    
 ' Insert your code for output 2(Cone Port1 of Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim BodyRadius        As Double
    BodyRadius = pipeDiam / 2

    Dim stPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
    Dim CenPos   As New AutoMath.DPosition
    CenPos.Set -parFacetoFace / 8, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, CenPos, BodyRadius, BodyRadius / 3)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 3(Cone Port2 of Body)
    Dim flangeThick2      As Double
    Dim enPoint   As New AutoMath.DPosition
    Dim CenPoint   As New AutoMath.DPosition
    CenPoint.Set parFacetoFace / 8, 0, 0
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, cptOffset, depth
    enPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, CenPoint, enPoint, BodyRadius / 3, BodyRadius)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 4(Stem of Operator)
    Dim StemEndPoint   As New AutoMath.DPosition
    StemEndPoint.Set 0, parOperatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, CenterPos, StemEndPoint, parFacetoFace / 8, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 5(Hand Wheel of Operator)
    Dim HandWheelTorusCenter   As New AutoMath.DPosition
    HandWheelTorusCenter.Set 0, parOperatorHeight * 0.75, 0
    Dim vNormal         As New AutoMath.DVector
    vNormal.Set 0, 1, 0
    Set ObjHandWheel = PlaceTorus(m_OutputColl, HandWheelTorusCenter, vNormal, parOperatorDiameter / 2, parFacetoFace / 16)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
 
 ' Insert your code for output 6(Hand Wheel Spokes along Body X)
    Dim parHWSpokeStposition   As New AutoMath.DPosition
    Dim parHWSpokeEndposition   As New AutoMath.DPosition
    
    parHWSpokeStposition.Set -parOperatorDiameter / 2, parOperatorHeight * 0.75, 0
    parHWSpokeEndposition.Set parOperatorDiameter / 2, parOperatorHeight * 0.75, 0

    Set ObjHWSpokesX = PlaceCylinder(m_OutputColl, parHWSpokeStposition, parHWSpokeEndposition, parFacetoFace / 16, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHWSpokesX
    Set ObjHWSpokesX = Nothing
 
 ' Insert your code for output 7(Hand Wheel Spokes along Z)
    Dim parHWSpokeZStposition   As New AutoMath.DPosition
    Dim parHWSpokeZEndposition   As New AutoMath.DPosition

    parHWSpokeZStposition.Set 0, parOperatorHeight * 0.75, -parOperatorDiameter / 2
    parHWSpokeZEndposition.Set 0, parOperatorHeight * 0.75, parOperatorDiameter / 2

    Set ObjHWSpokesZ = PlaceCylinder(m_OutputColl, parHWSpokeZStposition, parHWSpokeZEndposition, parFacetoFace / 16, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHWSpokesZ
    Set ObjHWSpokesZ = Nothing

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
