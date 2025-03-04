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
'   Author:          ACM
'   Creation Date:  Thursday, Nov 14 2002
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

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)


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

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

 ' Insert your code for output 1(Body)
    Dim Angle As Double
    Angle = PI / 6
    Dim ElbowRadius As Double
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    ElbowRadius = parFacetoCenter / Tan(Angle / 2)
    
    CP.Set -parFacetoCenter, ElbowRadius, 0
    CV.Set 0, 0, 1
  
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    
' Construct a circle that will be used to create the revolution
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -parFacetoCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
    Set ObjBody = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, Angle, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Place Nozzle 1
        
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim PortToOrigin As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    PortToOrigin = parFacetoCenter + cptOffset - depth
    oPlacePoint.Set -PortToOrigin, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    PortToOrigin = parFacetoCenter + cptOffset - depth
    oPlacePoint.Set PortToOrigin * Cos(Angle), PortToOrigin * Sin(Angle), 0
    oDir.Set Cos(Angle), Sin(Angle), 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

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

