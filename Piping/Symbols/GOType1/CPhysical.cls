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
'   Creation Date:  Saturday, Feb 17 2001
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
Private PI           As Double
Private RAD          As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

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
    Dim ObjBody As Object
    Dim ObjStem As Object
    Dim ObjOperator As Object
    Dim ObjAxis As Object
    Dim ObjWheel As Object
    Dim ObjSpoke As Object
    Dim ObjSpokeA As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parY4StemHeight As Double
    Dim parY5OperatorDiameter As Double
    Dim parY6HandwheelToOpTop As Double
    Dim parY7OperatorHeight As Double
    Dim parY8HandwheelProjection As Double
    Dim parY9HandwheelDiameter As Double
    Dim parY10HandwheelOffset As Double
    Dim parY11HandwheelToStemTop As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parY4StemHeight = arrayOfInputs(3)
    parY5OperatorDiameter = arrayOfInputs(4)
    parY6HandwheelToOpTop = arrayOfInputs(5)
    parY7OperatorHeight = arrayOfInputs(6)
    parY8HandwheelProjection = arrayOfInputs(7)
    parY9HandwheelDiameter = arrayOfInputs(8)
    parY10HandwheelOffset = arrayOfInputs(9)
    parY11HandwheelToStemTop = arrayOfInputs(10)
    parHandwheelAngle = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    iOutput = 0

 ' Insert your code for output 1(Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Set ObjBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
 ' Insert your code for output 2(Stem)
    Dim StemStartPoint   As New AutoMath.DPosition
    Dim StemEndPoint   As New AutoMath.DPosition
    StemStartPoint.Set 0, pipeDiam / 4, 0
    StemEndPoint.Set 0, parY4StemHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, StemStartPoint, StemEndPoint, parY5OperatorDiameter / 3, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 3(Operator)
    Dim operatorStartPoint   As New AutoMath.DPosition
    Dim operatorEndPoint   As New AutoMath.DPosition
    operatorStartPoint.Set 0, parY4StemHeight - parY11HandwheelToStemTop + parY6HandwheelToOpTop, 0
    operatorEndPoint.Set 0, operatorStartPoint.y - parY7OperatorHeight, 0
    Set ObjOperator = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parY5OperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperator
    Set ObjOperator = Nothing
    
 ' Insert your code for output 4(Axis)
'''    Dim parHandwheelAngleradian As Double
    Dim OpHwAxisStartPoint   As New AutoMath.DPosition
    Dim OpHwAxisEndPoint   As New AutoMath.DPosition
'''    parHandwheelAngleradian = parHandwheelAngle / RAD
    Dim parCosHWA As Double
    Dim parSinHWA As Double
'''    parCosHWA = Cos(parHandwheelAngleradian)
'''    parSinHWA = Sin(parHandwheelAngleradian)
    parCosHWA = Cos(parHandwheelAngle)
    parSinHWA = Sin(parHandwheelAngle)

    Dim ParSinY8 As Double
    Dim ParCosY8 As Double
    Dim ParSinY10 As Double
    Dim ParCosY10 As Double
    ParSinY8 = parSinHWA * parY8HandwheelProjection
    ParCosY8 = parCosHWA * parY8HandwheelProjection
    ParSinY10 = parSinHWA * parY10HandwheelOffset
    ParCosY10 = parCosHWA * parY10HandwheelOffset
    OpHwAxisStartPoint.Set -ParSinY10, parY4StemHeight - parY11HandwheelToStemTop, -ParCosY10
    OpHwAxisEndPoint.Set ParCosY8 - ParSinY10, OpHwAxisStartPoint.y, -(ParSinY8 + ParCosY10)
    Dim OpHwAxisExtraEndPoint   As New AutoMath.DPosition
    OpHwAxisExtraEndPoint.Set OpHwAxisEndPoint.x + (parCosHWA * parY5OperatorDiameter / 8), OpHwAxisEndPoint.y, OpHwAxisEndPoint.z - (parSinHWA * parY5OperatorDiameter / 8)
    Set ObjAxis = PlaceCylinder(m_OutputColl, OpHwAxisStartPoint, OpHwAxisExtraEndPoint, parY5OperatorDiameter / 8, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAxis
    Set ObjAxis = Nothing
    
 ' Insert your code for output 5(Wheel)
    Dim WheelCenter   As New AutoMath.DPosition
    Dim WheelAxis  As New AutoMath.DVector
    WheelAxis.Set parCosHWA, 0, -parSinHWA
'    WheelCenter.Set parY8HandwheelProjection - parY9HandwheelDiameter / 25, parY4StemHeight - parY11HandwheelToStemTop, -parY10HandwheelOffset
    WheelCenter.Set OpHwAxisEndPoint.x, OpHwAxisEndPoint.y, OpHwAxisEndPoint.z
    Set ObjWheel = PlaceTorus(m_OutputColl, WheelCenter, WheelAxis, parY9HandwheelDiameter / 2, parY9HandwheelDiameter / 20)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWheel
    Set ObjWheel = Nothing
 ' Insert your code for output 6(Spoke)
    Dim spokeStartPoint   As New AutoMath.DPosition
    Dim spokeEndPoint   As New AutoMath.DPosition
    spokeStartPoint.Set WheelCenter.x, WheelCenter.y - parY9HandwheelDiameter / 2, WheelCenter.z
    spokeEndPoint.Set WheelCenter.x, WheelCenter.y + parY9HandwheelDiameter / 2, WheelCenter.z
    Set ObjSpoke = PlaceCylinder(m_OutputColl, spokeStartPoint, spokeEndPoint, parY9HandwheelDiameter / 25, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpoke
    Set ObjSpoke = Nothing
 
 ' Insert your code for output 7(SpokeA)
    Dim spokeAStartPoint   As New AutoMath.DPosition
    Dim spokeAEndPoint   As New AutoMath.DPosition
    Dim ParSinHY9 As Double
    Dim ParCosHY9 As Double
    ParSinHY9 = parSinHWA * parY9HandwheelDiameter / 2
    ParCosHY9 = parCosHWA * parY9HandwheelDiameter / 2
    spokeAStartPoint.Set WheelCenter.x + ParSinHY9, WheelCenter.y, (WheelCenter.z + ParCosHY9)
    spokeAEndPoint.Set WheelCenter.x - ParSinHY9, WheelCenter.y, (WheelCenter.z - ParCosHY9)
    Set ObjSpokeA = PlaceCylinder(m_OutputColl, spokeAStartPoint, spokeAEndPoint, parY9HandwheelDiameter / 25, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeA
    Set ObjSpokeA = Nothing


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
