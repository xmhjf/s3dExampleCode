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
'   Creation Date:  Sunday, Feb 18 2001
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
    Dim ObjBody As Object
    Dim ObjStem As Object
    Dim ObjOperator As Object
    Dim ObjArm As Object
    Dim ObjHwAxis As Object
    Dim ObjWheel As Object
    Dim ObjSpokeY As Object
    Dim ObjSpokeZ As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parY4StemHeight As Double
    Dim parY5OperatorDiameter As Double
    Dim parY6HandwheelToOpTop As Double
    Dim parY7OperatorHeight As Double
    Dim parY8HandwheelProjection As Double
    Dim parY9OpArmProjection As Double
    Dim parY10OpArmWidth As Double
    Dim parY11HandwheelToArm As Double
    Dim parY12HandwheelDiameter As Double
    Dim parY13HandwheelToStemTop As Double
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
    parY9OpArmProjection = arrayOfInputs(8)
    parY10OpArmWidth = arrayOfInputs(9)
    parY11HandwheelToArm = arrayOfInputs(10)
    parY12HandwheelDiameter = arrayOfInputs(11)
    parY13HandwheelToStemTop = arrayOfInputs(12)
    parHandwheelAngle = arrayOfInputs(13)
    parInsulationThickness = arrayOfInputs(14)
    
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
    operatorStartPoint.Set 0, parY4StemHeight - parY13HandwheelToStemTop + parY6HandwheelToOpTop, 0
    operatorEndPoint.Set 0, operatorStartPoint.y - parY7OperatorHeight, 0
    Set ObjOperator = PlaceCylinder(m_OutputColl, operatorStartPoint, operatorEndPoint, parY5OperatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOperator
    Set ObjOperator = Nothing
    
 ' Insert your code for output 4(Arm)
    Dim parArmStartPoint   As New AutoMath.DPosition
    Dim parArmEndPoint   As New AutoMath.DPosition
    Dim parCosHWA As Double
    Dim parSinHWA As Double
    parCosHWA = Cos(parHandwheelAngle)
    parSinHWA = Sin(parHandwheelAngle)
    Dim ParSinY9 As Double
    Dim ParCosY9 As Double
    ParSinY9 = parSinHWA * parY9OpArmProjection
    ParCosY9 = parCosHWA * parY9OpArmProjection
    Dim ParSinOpD3 As Double
    Dim ParCosOpD3 As Double
    ParSinOpD3 = parSinHWA * parY5OperatorDiameter / 3
    ParCosOpD3 = parCosHWA * parY5OperatorDiameter / 3
    parArmStartPoint.Set ParCosOpD3, parY4StemHeight - parY13HandwheelToStemTop, -ParSinOpD3
    parArmEndPoint.Set ParCosY9, parArmStartPoint.y, -ParSinY9
    
    Set ObjArm = PlaceCylinder(m_OutputColl, parArmStartPoint, parArmEndPoint, parY10OpArmWidth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm
    Set ObjArm = Nothing
    
 ' Insert your code for output 5(HandWheel Axis)
    Dim parHwAxisStartPoint   As New AutoMath.DPosition
    Dim parHwAxisEndPoint   As New AutoMath.DPosition
    
    Dim ParSinY8 As Double
    Dim ParCosY8 As Double
    ParSinY8 = parSinHWA * parY8HandwheelProjection
    ParCosY8 = parCosHWA * parY8HandwheelProjection
    
    Dim ParSinAtW As Double
    Dim ParCosAtW As Double
    ParSinAtW = parSinHWA * parY11HandwheelToArm * 1.1
    ParCosAtW = parCosHWA * parY11HandwheelToArm * 1.1
    
    parHwAxisStartPoint.Set ParCosY8, parArmStartPoint.y, -ParSinY8
    parHwAxisEndPoint.Set ParCosY8 + ParSinAtW, parArmStartPoint.y, -(ParSinY8 - ParCosAtW)
    Set ObjHwAxis = PlaceCylinder(m_OutputColl, parHwAxisStartPoint, parHwAxisEndPoint, parY5OperatorDiameter / 8, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHwAxis
    Set ObjHwAxis = Nothing
    
 ' Insert your code for output 6(Hand Wheel)
    Dim WheelCenter   As New AutoMath.DPosition
    Dim ParSinWtArm As Double
    Dim ParCosWtArm As Double
    ParSinWtArm = parSinHWA * parY11HandwheelToArm
    ParCosWtArm = parCosHWA * parY11HandwheelToArm
    WheelCenter.Set ParCosY8 + ParSinWtArm, parArmStartPoint.y, -(ParSinY8 - ParCosWtArm)
    Dim WheelAxis  As New AutoMath.DVector
    WheelAxis.Set parSinHWA, 0, parCosHWA
    Set ObjWheel = PlaceTorus(m_OutputColl, WheelCenter, WheelAxis, parY12HandwheelDiameter / 2, parY12HandwheelDiameter / 20)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWheel
    Set ObjWheel = Nothing
    
 ' Insert your code for output 7(Wheel Spoke Y)
    Dim SpokeYStartPoint   As New AutoMath.DPosition
    Dim SpokeYEndPoint   As New AutoMath.DPosition
    SpokeYStartPoint.Set WheelCenter.x, WheelCenter.y - parY12HandwheelDiameter / 2, WheelCenter.z
    SpokeYEndPoint.Set WheelCenter.x, WheelCenter.y + parY12HandwheelDiameter / 2, WheelCenter.z
    Set ObjSpokeY = PlaceCylinder(m_OutputColl, SpokeYStartPoint, SpokeYEndPoint, parY12HandwheelDiameter / 25, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeY
    Set ObjSpokeY = Nothing
    
 ' Insert your code for output 8(Wheel Spoke Z)
    Dim ParSinWs As Double
    Dim ParCosWs As Double
    ParSinWs = parSinHWA * parY12HandwheelDiameter / 2
    ParCosWs = parCosHWA * parY12HandwheelDiameter / 2
    Dim SpokeZStartPoint   As New AutoMath.DPosition
    Dim SpokeZEndPoint   As New AutoMath.DPosition
    SpokeZStartPoint.Set WheelCenter.x - ParCosWs, WheelCenter.y, WheelCenter.z + ParSinWs
    SpokeZEndPoint.Set WheelCenter.x + ParCosWs, WheelCenter.y, WheelCenter.z - ParSinWs
    Set ObjSpokeZ = PlaceCylinder(m_OutputColl, SpokeZStartPoint, SpokeZEndPoint, parY12HandwheelDiameter / 25, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSpokeZ
    Set ObjSpokeZ = Nothing
    
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
