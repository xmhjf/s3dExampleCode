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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   Insulation.cls
'   Author:         svsmylav
'   Creation Date:  Monday, August 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol is Victaulic pressfit swaged 45 Degree (Style 511) Elbow taken from 12.02.pdf of Victaulic.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
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
'    Dim ObjElbow As Object
    Dim ObjInsTangentAtPort1 As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsTangentAtPort2 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    
    Dim parFacetoCenter As Double
    Dim TangentLength As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ArcStPointToCenter  As Double
    Dim parInsDiameterPort  As Double
    Dim parInsulationRadius As Double
    Dim parElbowRadius      As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'C to PE
    parInsulationThickness = arrayOfInputs(2)
    
    iOutput = 0
 
' Insert your code for output 2(Insulated Tangent At Port1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    TangentLength = depth

    ArcStPointToCenter = parFacetoCenter - TangentLength
    Dim parAngle As Double
    parAngle = PI / 4
    parElbowRadius = ArcStPointToCenter / Tan(parAngle / 2)
    parInsulationRadius = pipeDiam / 2 + parInsulationThickness
'   Ensure that the hub radius is not greater than the distance from the center
'   of rotation to the elbow center line
    If parInsulationRadius > parElbowRadius Then
        parInsulationRadius = parElbowRadius - 0.0001
    End If
    stPoint.Set CenterPos.x - parFacetoCenter, _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x - ArcStPointToCenter, _
                CenterPos.y, _
                CenterPos.z

    Set ObjInsTangentAtPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * parInsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangentAtPort1
    Set ObjInsTangentAtPort1 = Nothing
    
' Insert your code for output 3(Insulated Body)
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
'   Construct a circle that will be used to create the revolution
    Dim objCircleI   As IngrGeom3D.Circle3d
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                      parInsulationRadius)
    Set geomFactory = Nothing
'   Revolve the circle about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, parElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, parAngle, True)
                
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
    
' Insert your code for output 4(Insulated Tangent At Port2)
    stPoint.Set CenterPos.x + (parFacetoCenter) * Cos(parAngle), _
                CenterPos.y + (parFacetoCenter) * Sin(parAngle), _
                CenterPos.z
    enPoint.Set CenterPos.x + ArcStPointToCenter * Cos(parAngle), _
                CenterPos.y + ArcStPointToCenter * Sin(parAngle), _
                CenterPos.z
    Set ObjInsTangentAtPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * parInsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangentAtPort2
    Set ObjInsTangentAtPort2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
