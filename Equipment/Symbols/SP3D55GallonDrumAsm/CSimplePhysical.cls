VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          HL
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
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
    
    Dim CenterPos       As IJDPosition
    Dim Pos1            As IJDPosition
    Dim Axis            As IJDVector
    
    Dim iOutput         As Double
    Dim ObjBodyDrum     As Object
    Dim ObjBodyRing     As Object
    Dim objPlane        As Object
    
    Dim geomFactory     As New IngrGeom3D.GeometryFactory
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim centerPoint     As IngrGeom3D.Point3d
    
    Dim DrumHeight      As Double
    Dim DrumDiameter    As Double
    Dim RingHeight      As Double
    Dim RingDiameter    As Double
    
    Set CenterPos = New DPosition
    Set Pos1 = New DPosition
    Set Axis = New DVector
    
    CenterPos.Set 0, 0, 0
' Inputs
    
    iOutput = 0

' Insert your code for output 1(Body Drum)
' Height = 2'11" = 0.889m, Diameter = 1'11" = 0.5842m
' Bottom of the drum is the origin, CenterPos
' Everything will be measured from CenterPos
    DrumHeight = 0.889
    DrumDiameter = 0.5842
    
    Pos1.Set 0, 0, DrumHeight
    Set ObjBodyDrum = PlaceCylinder(m_OutputColl, CenterPos, Pos1, DrumDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyDrum
    Set ObjBodyDrum = Nothing
    
' Insert your code for output2 (First Ring)
' Height = 1/2" = 0.0127m, Diamter = 1'11.5" = 0.5969m
' Position = 11 1/2" = 0.2921m from CenterPos
    RingHeight = 0.0127
    RingDiameter = 0.5969
    Pos1.Set CenterPos.x, CenterPos.y, CenterPos.z + 0.2921
    Axis.Set 0, 0, 1
    Set ObjBodyRing = PlaceTorus(m_OutputColl, Pos1, Axis, RingDiameter / 2, RingHeight / 2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyRing
    Set ObjBodyRing = Nothing
    
' Insert your code for output3 (Second Ring)
' Position = 1'11 1/2" = 0.5969
    Pos1.Set CenterPos.x, CenterPos.y, CenterPos.z + 0.5969
    Axis.Set 0, 0, 1
    Set ObjBodyRing = PlaceTorus(m_OutputColl, Pos1, Axis, RingDiameter / 2, RingHeight / 2)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyRing
    Set ObjBodyRing = Nothing
    
' Insert your code for output4 (Default Surface)
    Axis.Set 0, 0, -1
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        CenterPos.x, CenterPos.y, CenterPos.z, _
                        Axis.x, Axis.y, Axis.z, DrumDiameter / 2)
    
    Set objPlane = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, objCircle)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objCircle = Nothing
    Set objPlane = Nothing
    
' Insert your code for output5 (Control Point)
' currently is standard point, to be replaced by control point
    Set centerPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, CenterPos.x, CenterPos.y, CenterPos.z)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), centerPoint
    Set centerPoint = Nothing
    Set geomFactory = Nothing
    
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
'''    Dim cptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
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
'''    Dim objNozzle   As IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - cptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzlePH(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + cptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzlePH(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
