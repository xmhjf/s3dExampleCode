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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         NNC
'   Creation Date:  Sunday, Jan 5 2003
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjSprocketRim As Object
    Dim ObjChainTop As Object
    Dim ObjChainLeft As Object
    Dim objChainRight As Object
    Dim ObjChainBottom As Object
    
    Dim parSprocketThickness As Double
    Dim parOperatorDiameter As Double
    Dim parHandwheelDiameter As Double
    Dim parOperatorHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSprocketThickness = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parHandwheelDiameter = arrayOfInputs(4)
    parOperatorHeight = arrayOfInputs(5)
   
    iOutput = 0

 ' Insert your code for output 1(Sprocket Rim)
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim objLineString   As IngrGeom3D.LineString3d
    Dim dRimH           As Double
    Dim dIRimR          As Double
    Dim lpoints         As Long
    Dim p(27)           As Double
    Dim axis            As IJDVector
    Dim pi              As Double
    Dim dChainR         As Double
    Dim dChainMR        As Double
    pi = 3.141592
    dRimH = (parOperatorDiameter - parHandwheelDiameter) / 2
    dIRimR = (parHandwheelDiameter / 2) + dRimH / 3
    lpoints = 9
    
    p(0) = 0
    p(1) = 0
    p(2) = parHandwheelDiameter / 2
    
    p(3) = p(0)
    p(4) = p(1)
    p(5) = parOperatorDiameter / 2
    
    p(6) = p(0)
    p(7) = p(1) + parSprocketThickness / 4
    p(8) = parOperatorDiameter / 2
    
    p(9) = p(0)
    p(10) = p(1) + parSprocketThickness / 4
    p(11) = dIRimR
    
    p(12) = p(0)
    p(13) = p(1) + parSprocketThickness * 3 / 4
    p(14) = dIRimR
    
    p(15) = p(0)
    p(16) = p(1) + parSprocketThickness * 3 / 4
    p(17) = parOperatorDiameter / 2
    
    p(18) = p(0)
    p(19) = p(1) + parSprocketThickness
    p(20) = parOperatorDiameter / 2
    
    p(21) = p(0)
    p(22) = p(1) + parSprocketThickness
    p(23) = parHandwheelDiameter / 2
    
    p(24) = p(0)
    p(25) = p(1)
    p(26) = p(2)
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, p)
    Set axis = New DVector
    axis.Set 0, 1, 0
    Set ObjSprocketRim = PlaceRevolution(m_OutputColl, objLineString, axis, CenterPos, 2 * 3.141586, True)
    Set objLineString = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSprocketRim
    Set ObjSprocketRim = Nothing
    
 ' Insert your code for output 2(Chain Top)
    Dim ChainCenter   As IJDPosition
    Set ChainCenter = New DPosition
    ChainCenter.Set 0, parSprocketThickness / 2, 0
    dChainR = parSprocketThickness / 4
    dChainMR = dIRimR + parSprocketThickness / 4
    Set ObjChainTop = PlaceTorusSweep(m_OutputColl, ChainCenter, axis, -dChainMR, dChainR, pi)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjChainTop
    Set ObjChainTop = Nothing

 ' Insert your code for output 3(Chain Left)
    Dim LeftChainTop   As IJDPosition
    Set LeftChainTop = New DPosition
    LeftChainTop.Set dChainMR, parSprocketThickness / 2, 0
    Dim LeftChainBottom   As IJDPosition
    Set LeftChainBottom = New DPosition
    LeftChainBottom.Set dChainMR, parSprocketThickness / 2, -(parOperatorHeight - 0.9 - dChainMR)
    Set ObjChainLeft = PlaceCylinder(m_OutputColl, LeftChainTop, LeftChainBottom, dChainR * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjChainLeft
    Set ObjChainLeft = Nothing

 ' Insert your code for output 4(Chain Right)
    Dim RightChainTop   As IJDPosition
    Set RightChainTop = New DPosition
    RightChainTop.Set -dChainMR, parSprocketThickness / 2, 0
    Dim RightChainBottom   As IJDPosition
    Set RightChainBottom = New DPosition
    RightChainBottom.Set -dChainMR, parSprocketThickness / 2, -(parOperatorHeight - 0.9 - dChainMR)
    Set objChainRight = PlaceCylinder(m_OutputColl, RightChainTop, RightChainBottom, dChainR * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objChainRight
    Set objChainRight = Nothing

 ' Insert your code for output 5(Chain Bottom)
'    Dim ChainBottomaxis            As IJDPosition
'    Set ChainBottomaxis = New DPosition
'    ChainBottomaxis.Set 0, ChainCenter.y, LeftChainBottom.z
'    Set ObjChainBottom = PlaceTorusSweep(m_OutputColl, ChainBottomaxis, axis, dChainMR, dChainR, pi)
   Dim stPoint1   As New AutoMath.DPosition
 Dim enPoint1   As New AutoMath.DPosition
 Dim ldiam     As Double
 Dim objCylinder  As Object
 stPoint1.Set 0, 0, 0
 enPoint1.Set 0, 1, 0
 ldiam = 1
 Set ObjChainBottom = PlaceCylinder(m_OutputColl, stPoint1, enPoint1, ldiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjChainBottom
    Set ObjChainBottom = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
