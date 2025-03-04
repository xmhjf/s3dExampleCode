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
'   dd.mmm.yyyy     who           change description
'   -----------     -----        ------------------
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
    parBaseHeight = arrayOfInputs(4)

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
    'Assumption: 1. Height of the Top-cylinder is Height of actuator / 6
    Dim dHeightofTopCylinder As Double
    dHeightofTopCylinder = dHeightofActuator / 6
      
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dMinorAxis As Double
    dMinorAxis = dHeightofActuator - dHeightofBaseCylinder - dHeightofTopCylinder
    
    Dim oEllipticalArc As Object
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, oCenterPos.y + dHeightofBaseCylinder, 0, _
                    0, 0, 1, _
                    0, dMinorAxis, 0, _
                    (parOperatorDiameter / 2) / dMinorAxis, _
                    0, PI / 2)

    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    oAxisVec.Set 0, 1, 0
    Dim oCenPoint As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0.0001, 0, 0
    
    Dim objBell As Object
    Set objBell = PlaceRevolution(m_OutputColl, oEllipticalArc, oAxisVec, _
                                    oCenPoint, 2 * PI, True)
                            
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBell
    Set oEllipticalArc = Nothing
    Set objBell = Nothing
    Set oGeomFactory = Nothing
    Set oAxisVec = Nothing
    Set oCenPoint = Nothing

' Insert your code for output 3, Top-cylinder
    'Assumption: Diameter of the Top-cylinder is Operator Diameter / 12.
    oStPoint.Set oCenterPos.x, oCenterPos.y + dHeightofActuator, oCenterPos.z
    oEnPoint.Set oStPoint.x, oCenterPos.y, oStPoint.z

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

