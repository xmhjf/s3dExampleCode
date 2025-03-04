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
'   Author:          svsmylav
'   Creation Date:  Friday, Feb 15 2002
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

'Private m_outputColl As Object
Private m_oErrors As IJEditErrors

Private Const MODULE = "Physical:" 'Used for error messages

Private Const E_FAIL = -2147467259

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitOD        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjConduitElbow As Object
    
    Dim parFacetoCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    
    iOutput = 0

' Insert your code for output 1(ConduitElbow)
'    RetrieveConduitParameters 1, oPartFclt, m_outputColl, ConduitOD, flangeThick, flangeDiam, cptOffset, depth
    RetrieveParameters 1, oPartFclt, m_outputColl, ConduitOD, flangeThick, flangeDiam, cptOffset, depth
    Dim ConduitOD1 As Double
    ConduitOD1 = BodyOD(flangeDiam, ConduitOD)
    RetrieveParameters 2, oPartFclt, m_outputColl, ConduitOD, flangeThick, flangeDiam, cptOffset, depth
    Dim ConduitOD2 As Double
    ConduitOD2 = BodyOD(flangeDiam, ConduitOD)
    Dim ConduitBodyOD As Double
    ConduitBodyOD = BodyOD(ConduitOD1, ConduitOD2)
    Dim parAngle As Double
    parAngle = PI / 4
    Dim parElbowRadius As Double
    parElbowRadius = parFacetoCenter / Tan(parAngle / 2)
    If ConduitBodyOD / 2 > parElbowRadius Then
        ConduitBodyOD = parElbowRadius * 2
    End If
'   Place Circle
    Dim objCirc1 As IngrGeom3D.Circle3d
    
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = CenterPos.x - parFacetoCenter
    cir1CenterY = CenterPos.y
    cir1CenterZ = CenterPos.z
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_outputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, ConduitBodyOD / 2)
                                                
' Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -parFacetoCenter, parElbowRadius, 0
    CV.Set 0, 0, 1
    
    Set ObjConduitElbow = PlaceRevolution(m_outputColl, objCirc1, _
                    CV, CP, parAngle, False)
   
   'Remove the Conduit circle
   Dim ob1 As IJDObject
   Set ob1 = objCirc1
   ob1.Remove
   Set ob1 = Nothing
   
'   Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitElbow
    Set ObjConduitElbow = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJConduitPortOcc
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoCenter - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_outputColl, oPartFclt, 1)
'   Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_outputColl, ConduitOD, flangeThick, flangeDiam, cptOffset, depth
    Dim parPorttoOrigin As Double
    parPorttoOrigin = parFacetoCenter + cptOffset - depth
    oPlacePoint.Set parPorttoOrigin * Cos(parAngle), parPorttoOrigin * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_outputColl, oPartFclt, 2)
'   Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


