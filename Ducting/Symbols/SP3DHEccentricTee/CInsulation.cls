VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DHEccentricTee.EccentricTee
'   Author:         KKK
'   Creation Date:  Thursday, Jan 19 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'   HVAC Eccentric Tee. TU - (Tangential –with hand made T-piece TSTU)
'   Symbol Taken from the http://www.lindab.com/Sweden/products/pdf/vent_eng/kapitel/eng_safe.pdf - page no.88
'
'   Change History:

'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   11.Sep.2006     dkl  TR-105007 In the Else condition for testing if 0.25<parWidth<0.4, replaced & with AND.
'   21.Jan-2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim ObjElbowIns As Object
    Dim ObjHVACNoz1Ins As Object
    Dim ObjHVACNoz2Ins As Object
    Dim ObjHVACNoz3Ins As Object
         
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parLength As Double
    Dim parBLength As Double
    Dim parInsulationThickness As Double
    Dim dblInsulationRadius As Double
    Dim dblBranchInsRadius As Double
    Dim dblOffset As Double
    
    Dim parAngle As Double
    
    parAngle = PI / 2

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parBLength = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
    dblOffset = parWidth / 2 - parBWidth / 2
    
' Insert Code for OUTPUT 1 - Body One

    dblInsulationRadius = parWidth / 2 + parInsulationThickness
    dblBranchInsRadius = parBWidth / 2 + parInsulationThickness

' Insert your code for output 3(Nozzle 1 insulation)
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parLength / 2, 0, 0
    enPoint.Set parLength / 2, 0, 0

    Set ObjHVACNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * dblInsulationRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz1Ins
    Set ObjHVACNoz1Ins = Nothing

' Insert your code for output 4 (Nozzle 2 insulation)
    stPoint.Set 0, 0, dblOffset
    enPoint.Set 0, parBLength, dblOffset
        
    Set ObjHVACNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * dblBranchInsRadius, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz3Ins
    
    Set ObjHVACNoz3Ins = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

'''<{(Cylinder begin)}>
Public Function PlaceCylinder(ByVal objOutputColl As Object, _
                                lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                lDiameter As Double, _
                                isCapped As Boolean) As Object

''' This function creates persistent projetion of circle
''' based on two points (axis of cylinder) and diameter
''' Example of call:
''' Dim stPoint   As new AutoMath.DPosition
''' Dim enPoint   As new AutoMath.DPosition
''' Dim ldiam     as long
''' Dim objCylinder  As object
''' stPoint.set 0, 0, 0
''' enPoint.set 0, 0, 1
''' lDiam = 1.5
''' set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, lDiam, True)
''' m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
''' Set objCylinder = Nothing

    Const METHOD = "PlaceCylinder:"
    On Error GoTo ErrorHandler
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.x - lStartPoint.x, _
                     lEndPoint.y - lStartPoint.y, _
                     lEndPoint.z - lStartPoint.z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
    
' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, isCapped)
    
    Set objCircle = Nothing
    
    Set PlaceCylinder = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
        
End Function
'''<{(Cylinder end)}>
