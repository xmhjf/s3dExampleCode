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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, August 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol is Victaulic pressfit swaged 90 Degree (Style 510) Elbow taken from 12.02.pdf of Victaulic.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim ObjElbow    As Object
    
    Dim parFacetoCenter As Double
    Dim TangentLength As Double
    Dim parInsulationThickness As Double

    Dim parElbowRadius   As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'C to PE
    parInsulationThickness = arrayOfInputs(2)
    
    iOutput = 0

' Insert your code for output 1(Elbow)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    TangentLength = depth


        
'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    parElbowRadius = parFacetoCenter - TangentLength
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parElbowRadius, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
    Set geomFactory = Nothing
'   Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector   'rotation vector for rotation
    CP.Set -parElbowRadius, parElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, CV, CP, PI / 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircle
    Set objCircle = Nothing
    sp3dElem1.Remove
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
'   faceToFace = arrayOfInputs(2)
'   oPlacePoint.Set -faceToFace / 2 - sptOffset + depth, 0, 0
    oPlacePoint.Set CenterPos.x - parFacetoCenter - sptOffset + depth, _
                    CenterPos.y, _
                    CenterPos.z
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, TangentLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
'   oPlacePoint.Set faceToFace / 2 + sptOffset - depth, 0, 0
    oPlacePoint.Set CenterPos.x, _
                    CenterPos.y + parFacetoCenter + sptOffset - depth, _
                    CenterPos.z
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, TangentLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set CenterPos = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

