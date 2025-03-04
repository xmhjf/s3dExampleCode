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
'   ProgID:         SP3DHCleaningBend.CleanBend
'   Author:         KKK
'   Creation Date:  Thursday, Jan 19 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for the simple Physical aspect of
'   HVAC Cleaning Bend. (BBKCU 90 Deg.)
'   Symbol Taken from the http://www.lindab.com/Sweden/products/pdf/vent_eng/kapitel/eng_safe.pdf
'
'   Change History:

'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.Jan-2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:"    'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim objCircle As IngrGeom3D.Circle3d
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
    Dim parInsulationRadius As Double

    Dim CP As New AutoMath.DPosition    'arc center point
    Dim CV As New AutoMath.DVector    'rotation vector for rotation
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


    ' Insert Code for OUTPUT 1 - Body One
    parInsulationRadius = parWidth / 2 + parInsulationThickness
    If CmpDblGreaterthan(parInsulationRadius, parLength) Then
        parInsulationRadius = parLength * 0.99
    End If

    ' Insert your code for Body Insulation
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                     -parLength, 0, 0, _
                                                                     1, 0, 0, _
                                                                     parInsulationRadius)

    CP.Set -parLength, parLength, 0
    CV.Set 0, 0, 1

    Set ObjElbowIns = PlaceRevolution(m_OutputColl, objCircle, _
                                      CV, CP, parAngle, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbowIns
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing
    Set ObjElbowIns = Nothing

    ' Insert your code for output 3(Nozzle 1 insulation)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -(parLength + 0.0001), 0, 0
    enPoint.Set -parLength, 0, 0

    Set ObjHVACNoz1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInsulationRadius, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz1Ins
    Set ObjHVACNoz1Ins = Nothing


    ' Insert your code for output 5 (Nozzle 2 insulation)

    stPoint.Set 0, parLength + 0.0001, 0
    enPoint.Set 0, parLength, 0

    Set ObjHVACNoz2Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInsulationRadius, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHVACNoz2Ins
    Set ObjHVACNoz2Ins = Nothing

    ' Insert your code for output 4 (Nozzle 3 insulation)

    stPoint.Set -parLength + (parLength + parBLength) * Cos(PI / 4), _
                parLength - (parLength + parBLength) * Sin(PI / 4), 0
    enPoint.Set -parWidth / 2, parWidth / 2, 0

    Set ObjHVACNoz3Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInsulationRadius, True)

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

    Dim circleCenter As AutoMath.DPosition
    Dim circleNormal As AutoMath.DVector
    Dim objCircle As IngrGeom3D.Circle3d
    Dim dblCylWidth As Double
    Dim objProjection As IngrGeom3D.Projection3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

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


