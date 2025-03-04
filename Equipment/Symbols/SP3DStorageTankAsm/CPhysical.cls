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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:39 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Feb 6 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This Symbol is a replicate of "Tank.TankServices" Symbol, except that the symbol is in line with
' Symbol created using wizard.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   24.Oct.2003     SymbolTeam(India)       TR 51102 Removed duplicated Header/Copyright Information.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   18.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
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
    
    Dim iOutput     As Double
    Dim ObjTankBodyCylinder As IngrGeom3D.Projection3d
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
    Dim ObjDefaultSurface As IngrGeom3D.Plane3d
    Dim ObjConnectPort1 As IngrGeom3D.Point3d
    Dim ObjConnectPort2 As IngrGeom3D.Point3d
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim ObjEndCap As IngrGeom3D.Revolution3d
        
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportThickness As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSupportLength = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)

    iOutput = 0
    
 ' Insert your code for output 1(TankBodyCylinder)
    Dim ellipse As Object
    Dim projVect As AutoMath.DVector
    Dim projection As IngrGeom3D.Projection3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set projVect = New AutoMath.DVector
    
    Dim axesRatio As Double
    Dim circleCenterX As Double
    Dim circleCenterY As Double
    Dim circleCenterZ As Double
    Dim circleNormalX As Double
    Dim circleNormalY As Double
    Dim circleNormalZ As Double
    Dim circlePointVecX As Double
    Dim circlePointVecY As Double
    Dim circlePointVecZ As Double
    
    circleCenterX = 0#
    circleCenterY = 0#
    circleCenterZ = parSupportHeight + parVesselDiameter * 0.5

    circleNormalX = 1#
    circleNormalY = 0#
    circleNormalZ = 0#
    
    circlePointVecX = 0#
    circlePointVecY = parVesselDiameter * 0.5
    circlePointVecZ = 0#
    
    axesRatio = 1#
    Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                circleNormalX, circleNormalY, circleNormalZ, _
                                circlePointVecX, circlePointVecY, circlePointVecZ, _
                                axesRatio)

    projVect.Set 1, 0, 0
    Set ObjTankBodyCylinder = PlaceProjection(m_OutputColl, ellipse, projVect, parVesselLength, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTankBodyCylinder
    Set ObjTankBodyCylinder = Nothing
    Set projVect = Nothing
    
    'Delete ellipse which was created for creating projection.
    Dim gscadElem As IJDObject
    Set gscadElem = ellipse
    Debug.Assert Not gscadElem Is Nothing
    Set ellipse = Nothing
    gscadElem.Remove
    
' Insert your code for output 3(Elliptical Head)
    Dim revVect As AutoMath.DVector
    Dim centerPoint As AutoMath.DPosition
    Dim ellipseCenter As AutoMath.DPosition
    Dim ellipseNormal As AutoMath.DPosition
    Dim majorAxis As AutoMath.DPosition
    
    Set majorAxis = New AutoMath.DPosition
    Set ellipseNormal = New AutoMath.DPosition
    Set ellipseCenter = New AutoMath.DPosition
    Set revVect = New AutoMath.DVector
    Set centerPoint = New AutoMath.DPosition
    
    ellipseCenter.Set parVesselLength, 0, parSupportHeight + parVesselDiameter * 0.5
    majorAxis.Set 0, 0, parVesselDiameter / 2
    ellipseNormal.Set 0, -1, 0
    
    axesRatio = 0.5
    Set ellipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                axesRatio, 1.5 * PI, PI / 2)
    revVect.Set 1, 0, 0
    centerPoint.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    Set ObjEndCap = PlaceRevolution(m_OutputColl, ellipse, revVect, centerPoint, PI * 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndCap
    Set ObjEndCap = Nothing
    
    'Delete ellipse which was created for creating projection.
    Set gscadElem = ellipse
    Debug.Assert Not gscadElem Is Nothing
    Set ellipse = Nothing
    gscadElem.Remove
 
' Insert your code for output 4(Elliptical Head)
'Elliptical head minor radius is 0.5 times the Tank Radius
    
    ellipseCenter.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    majorAxis.Set 0, 0, parVesselDiameter / 2
    ellipseNormal.Set 0, 1, 0
    axesRatio = 0.5
    Set ellipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                ellipseCenter.x, ellipseCenter.y, ellipseCenter.z, _
                                ellipseNormal.x, ellipseNormal.y, ellipseNormal.z, _
                                majorAxis.x, majorAxis.y, majorAxis.z, _
                                axesRatio, 1.5 * PI, PI / 2)
    revVect.Set 1, 0, 0
    centerPoint.Set 0, 0, parSupportHeight + parVesselDiameter * 0.5
    Set ObjEndCap = PlaceRevolution(m_OutputColl, ellipse, revVect, centerPoint, PI * 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndCap
    Set ObjEndCap = Nothing
    Set revVect = Nothing
    Set centerPoint = Nothing
    
    'Delete ellipse which was created for creating projection.
    Set gscadElem = ellipse
    Debug.Assert Not gscadElem Is Nothing
    Set ellipse = Nothing
    gscadElem.Remove
    
    Set majorAxis = Nothing
    Set ellipseNormal = Nothing
    Set ellipseCenter = Nothing
    Set revVect = Nothing
    Set centerPoint = Nothing
    
    
 ' Insert your code for output 8(Support1)
    Dim lowerCorner As AutoMath.DPosition
    Dim upperCorner As AutoMath.DPosition
    
    Set lowerCorner = New AutoMath.DPosition
    Set upperCorner = New AutoMath.DPosition
    
    lowerCorner.Set parVesselLength * 0.25 - parSupportThickness * 0.5, _
                                -0.5 * parSupportLength, _
                                0
    upperCorner.Set lowerCorner.x + parSupportThickness, _
                                0.5 * parSupportLength, _
                                parSupportHeight + parVesselDiameter * 0.25
    Set ObjSupport1 = PlaceBox(m_OutputColl, lowerCorner, upperCorner)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1
    Set ObjSupport1 = Nothing
    
 ' Insert your code for output 9(Support2)
    lowerCorner.Set parVesselLength * 0.75 - parSupportThickness * 0.5, _
                                -0.5 * parSupportLength, _
                                0
    upperCorner.Set lowerCorner.x + parSupportThickness, _
                                0.5 * parSupportLength, _
                                parSupportHeight + parVesselDiameter * 0.25
    Set ObjSupport2 = PlaceBox(m_OutputColl, lowerCorner, upperCorner)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2
    Set ObjSupport2 = Nothing
    Set lowerCorner = Nothing
    Set upperCorner = Nothing
    
' Insert your code for output 10(DefaultSurface1)
'This is constructed at the bottom of the support1

    Dim Points(0 To 11) As Double

    Points(0) = parVesselLength * 0.25 - parSupportThickness * 0.5
    Points(1) = -0.5 * parSupportLength
    Points(2) = 0#
    
    Points(3) = Points(0)
    Points(4) = Points(1) + parSupportLength
    Points(5) = 0#
    
    Points(6) = Points(0) + parSupportThickness
    Points(7) = Points(1) + parSupportLength
    Points(8) = 0#
    
    Points(9) = Points(0) + parSupportThickness
    Points(10) = Points(1)
    Points(11) = 0#

    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, Points)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
 ' Insert your code for output 11(DefaultSurface2)
    'This is constructed at the bottom of the support2
    
    Points(0) = parVesselLength * 0.75 - parSupportThickness * 0.5
    Points(1) = -0.5 * parSupportLength
    Points(2) = 0#
    
    Points(3) = Points(0)
    Points(4) = Points(1) + parSupportLength
    Points(5) = 0#
    
    Points(6) = Points(0) + parSupportThickness
    Points(7) = Points(1) + parSupportLength
    Points(8) = 0#
    
    Points(9) = Points(0) + parSupportThickness
    Points(10) = Points(1)
    Points(11) = 0#

    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, Points)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
 ' Insert your code for output 12(ConnectPort1)
    Set ObjConnectPort1 = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                    parVesselLength * 0.25 - parSupportThickness * 0.5, _
                                                    -0.5 * parSupportLength, 0)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnectPort1
    Set ObjConnectPort1 = Nothing
    
 ' Insert your code for output 13(ConnectPort2)
    Set ObjConnectPort2 = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, parVesselLength * 0.75 + parSupportThickness * 0.5, _
                                                    -0.5 * parSupportLength, 0)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnectPort2
    Set ObjConnectPort2 = Nothing
    Set geomFactory = Nothing
    
 ' Insert your code for output 14(FoundationPort1)
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(1 To 3) As Double
    Dim dXaxis(1 To 3) As Double
    Dim dZaxis(1 To 3) As Double
    
    'Coordinate sysytem is at the bottom of support.
    '
    '               |-------|
    '   Y           |       |
    '   ^           |       |
    '   |           |       |
    '   |           |       |
    '   |           |       |
    '   -----> X    |       |
    '   Symbol CS   |       |
    '               |       |
    '               |   X   |
    '               |   ^   |
    '               |   |   |
    '               |   |   |
    '               |---|---|----> Y
    '                   Port CS
    
    dOrigin(1) = parVesselLength * 0.25
    dOrigin(2) = -0.5 * parSupportLength
    dOrigin(3) = 0#
    
    dXaxis(1) = 0#
    dXaxis(2) = 1#
    dXaxis(3) = 0#
    
    dZaxis(1) = 0#
    dZaxis(2) = 0#
    dZaxis(3) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort.PutCS(dOrigin(1), dOrigin(2), dOrigin(3), _
                            dXaxis(1), dXaxis(2), dXaxis(3), _
                            dZaxis(1), dZaxis(2), dZaxis(3))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    
 ' Insert your code for output 15(FoundationPort2)
    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort2", _
                                                    False, m_OutputColl.ResourceManager)
    dOrigin(1) = parVesselLength * 0.75
    Call ObjFoundationPort.PutCS(dOrigin(1), dOrigin(2), dOrigin(3), _
                            dXaxis(1), dXaxis(2), dXaxis(3), _
                            dZaxis(1), dZaxis(2), dZaxis(3))

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
     
End Sub
