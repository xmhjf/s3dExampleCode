VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Monday, Mar 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using eight outputs. Out of these three are Physical aspect outputs.
' This class contains five outputs: two tangents plus two port insulations are prepared by using
' PlaceCylinder' function. Insulation for bend is prepared by using 'PlaceRevolution' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsTangentAtPort1 As Object
    Dim ObjInsulatedBody   As Object
    Dim ObjInsTangentAtPort2 As Object
    Dim ObjInsulatedPort1  As Object
    Dim ObjInsulatedPort2  As Object
    
    Dim parFacetoCenter As Double
    Dim parTangentLength   As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationRadius     As Double
    Dim parElbowRadius           As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parTangentLength = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'''    MsgBox parInsulationThickness * 1000
    iOutput = 0

' Insert your code for output 2(Insulated Body  Tangent at Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    parInsulationRadius = pipeDiam / 2 + parInsulationThickness
    parElbowRadius = parFacetoCenter - parTangentLength
'   Ensure that the hub radius is not greater than the distance from the center
'   of rotation to the elbow center line
    If parInsulationRadius > parElbowRadius Then
        parInsulationRadius = parElbowRadius - 0.001
    End If
''''    MsgBox parInsulationRadius * 1000
    
    stPoint.Set CenterPos.x - parFacetoCenter + flangeThick, _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x - parFacetoCenter + parTangentLength, _
                CenterPos.y, _
                CenterPos.z

    Set ObjInsTangentAtPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * parInsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangentAtPort1
    Set ObjInsTangentAtPort1 = Nothing
    
' Insert your code for output 3(Insulated Body of Bend)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim objCircleI   As IngrGeom3D.Circle3d
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parElbowRadius, 0, 0, _
                        1, 0, 0, _
                        parInsulationRadius)
    Set geomFactory = Nothing
'   Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector   'rotation vector for rotation
    CP.Set -parElbowRadius, parElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, PI / 2, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircleI
    Set objCircleI = Nothing
    sp3dElem1.Remove
    
' Insert your code for output 4(Insulated Body  Tangent at Port 2)
    stPoint.Set CenterPos.x, _
                CenterPos.y + parFacetoCenter - flangeThick, _
                CenterPos.z
    enPoint.Set CenterPos.x, _
                CenterPos.y + parFacetoCenter - parTangentLength, _
                CenterPos.z
    Set ObjInsTangentAtPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            2 * parInsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangentAtPort2
    Set ObjInsTangentAtPort2 = Nothing

'''''''''''''' Insert your code for output 5 (Insulated Port1)
'''''''''''''    stPoint.Set CenterPos.x - parFacetoCenter, _
'''''''''''''                CenterPos.y, _
'''''''''''''                CenterPos.z
'''''''''''''    enPoint.Set CenterPos.x - parFacetoCenter + flangeThick + parInsulationThickness, _
'''''''''''''                CenterPos.y, _
'''''''''''''                CenterPos.z
'''''''''''''    parInsDiameterPort = flangeDiam + parInsulationThickness * 2
'''''''''''''    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
''''''''''''''   Set the output
'''''''''''''    iOutput = iOutput + 1
'''''''''''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
'''''''''''''    Set ObjInsulatedPort1 = Nothing
'''''''''''''
'''''''''''''' Insert your code for output 6(Insulated Port2)
'''''''''''''    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''''''''''''
'''''''''''''    stPoint.Set CenterPos.x, _
'''''''''''''                CenterPos.y + parFacetoCenter, _
'''''''''''''                CenterPos.z
'''''''''''''    enPoint.Set CenterPos.x, _
'''''''''''''                CenterPos.y + parFacetoCenter - flangeThick - parInsulationThickness, _
'''''''''''''                CenterPos.z
'''''''''''''    parInsDiameterPort = flangeDiam + parInsulationThickness * 2
'''''''''''''    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameterPort, True)
'''''''''''''
''''''''''''''   Set the output
'''''''''''''    iOutput = iOutput + 1
'''''''''''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
'''''''''''''    Set ObjInsulatedPort2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
