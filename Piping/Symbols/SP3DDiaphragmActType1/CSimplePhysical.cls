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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         dkl
'   Creation Date:  Monday, Apr 23 2007
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      This is a diaphragm Actuator symbol created to be associated with Pressure Reducing valves.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   23.Apr.2007     dkl           CR-118766  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Dim PI As Double

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub
Private Sub Class_terminate()
    Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parBaseHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
    parOperatorDiameter = arrayOfInputs(3)
    parBaseHeight = arrayOfInputs(4) 'OperatorHeight1

    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0
    
    'The Actuator is located at a distance of Base height from the origin along Y-Axis.
    Dim oCenterPos As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    oCenterPos.Set 0, parBaseHeight, 0
    
' Insert your code for output 1, Base Cylinder.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim dHeightofActuator As Double 'Height of actuator from base Cylinder bottom to
                                    'the top of the actuator.
    dHeightofActuator = parOperatorHeight - parBaseHeight
    
    'Assumption: Height of the Base-Cylinder is Height of actuator / 12.
    Dim dHeightofBaseCylinder As Double
    dHeightofBaseCylinder = dHeightofActuator / 12
    
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oStPoint.x, oStPoint.y + dHeightofBaseCylinder, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, parOperatorDiameter

'Create Output for Bell portion.
    Dim dRadiusofBottomArc As Double
    dRadiusofBottomArc = parOperatorDiameter / 12
    
    Dim dRadiusofTopArc As Double
    dRadiusofTopArc = parOperatorDiameter / 10
    
    Dim dWidthofBell As Double
    dWidthofBell = 0.35 * parOperatorDiameter
    
    'Assumption: 1. Height of the Top-cylinder is Height of actuator / 6
    Dim dHeightofTopCylinder As Double
    dHeightofTopCylinder = dHeightofActuator / 6
    
    Dim dHeightofBell As Double
    dHeightofBell = dHeightofActuator - dHeightofBaseCylinder - dHeightofTopCylinder
    
    Dim dPoints() As Double
    ReDim dPoints(0 To 20) 'representing points in the Y-Z plane
    'Point 1
    dPoints(0) = 0  'X
    dPoints(1) = oCenterPos.y + dHeightofBaseCylinder   'Y
    dPoints(2) = 0  'Z
    'Point 2
    dPoints(3) = -dWidthofBell
    dPoints(4) = dPoints(1)
    dPoints(5) = dPoints(2)
    'Point 3
    dPoints(6) = -dWidthofBell + dRadiusofBottomArc
    dPoints(7) = dPoints(4) + dRadiusofBottomArc
    dPoints(8) = dPoints(2)
    'Point 4
    dPoints(9) = dPoints(6) + dRadiusofBottomArc
    dPoints(10) = dPoints(7) + dRadiusofBottomArc
    dPoints(11) = dPoints(2)
    'Point 5
    dPoints(12) = dPoints(9)
    dPoints(13) = dPoints(1) + dHeightofBell - dRadiusofTopArc
    dPoints(14) = dPoints(2)
    'Point 6
    dPoints(15) = dPoints(12) + dRadiusofTopArc
    dPoints(16) = dPoints(13) + dRadiusofTopArc
    dPoints(17) = dPoints(2)
    'Point 7
    dPoints(18) = dPoints(0)
    dPoints(19) = dPoints(16)
    dPoints(20) = dPoints(2)
      
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oLine As IngrGeom3D.Line3d
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPoints(0), dPoints(1), dPoints(2), _
                dPoints(3), dPoints(4), dPoints(5))
    
    Dim objCollection As Collection
    Set objCollection = New Collection
    objCollection.Add oLine
    
    Dim oArc As IngrGeom3D.Arc3d
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                dPoints(6), dPoints(4), dPoints(5), _
                dPoints(3), dPoints(4), dPoints(5), _
                dPoints(6), dPoints(7), dPoints(8))
    objCollection.Add oArc
    
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                dPoints(6), dPoints(10), dPoints(8), _
                dPoints(6), dPoints(7), dPoints(8), _
                dPoints(9), dPoints(10), dPoints(11))
    objCollection.Add oArc
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPoints(9), dPoints(10), dPoints(11), _
                dPoints(12), dPoints(13), dPoints(14))
    objCollection.Add oLine
    
    'The Top Arc
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                dPoints(15), dPoints(13), dPoints(14), _
                dPoints(12), dPoints(13), dPoints(14), _
                dPoints(15), dPoints(16), dPoints(17))
    objCollection.Add oArc
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPoints(15), dPoints(16), dPoints(17), _
                dPoints(18), dPoints(19), dPoints(20))
    objCollection.Add oLine
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                dPoints(18), dPoints(19), dPoints(20), _
                dPoints(0), dPoints(1), dPoints(2))
    objCollection.Add oLine
    
    oStPoint.Set dPoints(0), dPoints(1), dPoints(2)
    Dim objBellLeftHalfOutline As IngrGeom3D.ComplexString3d
    Set objBellLeftHalfOutline = PlaceTrCString(oStPoint, objCollection)
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 0, 1, 0
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0.0001, 0, 0
    
    Dim objBell As Object
    Set objBell = PlaceRevolution(m_OutputColl, objBellLeftHalfOutline, oAxisVec, _
                                    oCenPoint, 2 * PI, True)
                            
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBell
    Set oLine = Nothing
    Set oArc = Nothing
    Set objCollection = Nothing
    Set objBell = Nothing
    Set oGeomFactory = Nothing
    Set objBellLeftHalfOutline = Nothing
    Set oAxisVec = Nothing
    Set oCenPoint = Nothing

' Insert your code for output 3, Top-cylinder
    'Assumption: Diameter of the Top-cylinder is Operator Diameter / 12.
    oStPoint.Set oCenterPos.x, oCenterPos.y + dHeightofActuator, oCenterPos.z
    oEnPoint.Set oStPoint.x, oStPoint.y - dHeightofTopCylinder, oStPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                                oEnPoint, parOperatorDiameter / 12
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
    
End Sub

