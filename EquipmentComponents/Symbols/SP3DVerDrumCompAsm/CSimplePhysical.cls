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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:02 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, May 16, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjCylinder As Object
    Dim ObjDish1 As Object
    Dim ObjDish2 As Object
    Dim objSphere As Object
    Dim objPlane As Object
    
    Dim strPosition As String
    Dim strDirection As String
    Dim dblDiameter As Double
    Dim dblLength As Double
    
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
        
    iOutput = 0
    
 ' Insert your code for output 1(Drum Body)
    strPosition = "E 0 N 0 U " & parVesselTantoTan / 2
    strDirection = "N 0 U 90"
    dblDiameter = parVesselDiameter
    dblLength = parVesselTantoTan
 
    Set ObjCylinder = CreateCylinder(m_outputColl, strPosition, dblDiameter, dblLength, strDirection)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    
 ' Insert your code for output 2(Drum End1)
    strPosition = "E 0 N 0 U " & parVesselTantoTan
    strDirection = "N 0 U 90"
    dblDiameter = parVesselDiameter
    
    Set ObjDish1 = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish1
    Set ObjDish1 = Nothing
    
 ' Insert your code for output 3(Drum End2)
    strPosition = "E 0 N 0 U 0"
    strDirection = "N 0 D -90"
    dblDiameter = parVesselDiameter
    Set ObjDish2 = Create2to1EllipticalDish(m_outputColl, strPosition, strDirection, dblDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDish2
    Set ObjDish2 = Nothing
    
' Insert your code for output 4 (Point)
    strPosition = "E 0 N 0 U 0"
    dblDiameter = 0.075
    Set objSphere = CreateSphere(m_outputColl, strPosition, dblDiameter)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objSphere
    Set objSphere = Nothing
    
'Insert your Code for the Plane
    Dim oPlane As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim circleCenter As IJDPosition
    Dim circleNormal As DVector
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set circleCenter = New DPosition
    Set circleNormal = New DVector
    
    circleNormal.Set 0, 0, -1
    circleCenter.Set 0, 0, 0

    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            circleCenter.x, circleCenter.y, circleCenter.z, _
                            circleNormal.x, circleNormal.y, circleNormal.z, _
                            0.1 * parVesselDiameter / 2)
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, oCircle)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPlane
    
    Set oPlane = Nothing
    Set oCircle = Nothing
    Set circleNormal = Nothing
    Set circleCenter = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
