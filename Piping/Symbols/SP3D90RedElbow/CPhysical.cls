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
'   Author:         MS
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: F84 Page No. D-56 of PDS Piping Component Data Reference Guide.
'    Symbol is created with Seven Outputs
'    The Four physical aspect outputs are created as follows:
'    ObjElbow -Using 'PlaceRevolution' function, ObjReducingBody using 'PlaceCone'
'    One ObjNozzle object by using 'CreateNozzle' function and another ObjNozzle by using CreateNozzleWithLength
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    Dim pipeDiam2        As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double

    Dim iOutput     As Double
    Dim ObjElbow As Object
    Dim ObjReducingBody As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parReducingLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parReducingLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0

 ' Insert your code for output 1(Body of Elbow)
    Dim PI           As Double
    PI = 4 * Atn(1)
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset2, depth2

' Construct a circle that will be used to create the revolution
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -parFace1toCenter + parReducingLength, 0, 0, _
                        1, 0, 0, _
                      pipeDiam2 / 2)
' Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -parFace1toCenter + parReducingLength, parFace1toCenter - parReducingLength, 0
    CV.Set 0, 0, 1

    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, CV, CP, PI / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set geomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing

' Insert your code for output 2(Reducing Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + parReducingLength, 0, 0
    Set ObjReducingBody = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, pipeDiam2 / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducingBody
    Set ObjReducingBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    Dim dNozzleLength As Double
    
    oPlacePoint.Set 0, parFace2toCenter + cptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
    dNozzleLength = parFace2toCenter - (parFace1toCenter - parReducingLength)
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

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
