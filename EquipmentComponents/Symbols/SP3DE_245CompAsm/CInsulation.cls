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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:01 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:        svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Simple Horizontal Cylindrical Vessel Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E245 Symbol in Page no 298.
'   Symbol is created using the following Outputs:
'   i)  Seven standard outputs Consisting of the following:
'       a) Three Insulation aspect outputs,
'       b) Three Physical aspect outputs:
'           Vessel created using 'PlaceCylinder' and two domes using 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: default surface and a control point.
'
'   ii) Variable number of Supports (Maximum two supports) and surfaces are
'       computed as per the user input.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
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
    Dim ObjInsulatedVessel As Object
    Dim ObjInsulatedVesselLEH As Object
    Dim ObjInsulatedVesselREH As Object

    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)      'P2
    parVesselDiameter = arrayOfInputs(3)    'P3
    parVesselCenterHeight = arrayOfInputs(4)    'P4
    parFirstSupportLocation = arrayOfInputs(5)  'P5
    parSecondSupportLocation = arrayOfInputs(6) 'P6
    parSupportThickness = arrayOfInputs(7)  'P7
    parSupportLength = arrayOfInputs(8)     'P8
    parInsulationThickness = arrayOfInputs(9)

    iOutput = 0

'   Origin is taken at left dome center point (i.e. Placement point 1)
'   Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis

' Insert your code for output 5(InsulatedVessel)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parVesselLength, 0, 0

'   Create Insulated Horizontal Drum
    Dim dInsulationDiameter As Double

    dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing

' Insert your code for output 6(InsulatedVesselLEH)
'   Revolve Left Elliptical Dome Curve about X-Axis
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double

    Dim oEqpLeftArcI As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArcI As IngrGeom3D.EllipticalArc3d

    cenX = 0
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = dInsulationDiameter / 2
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (dInsulationDiameter / 2) '0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0

    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set oEqpLeftArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    axisVect.Set 1, 0, 0
    revCenPt.Set parVesselLength / 2, 0, 0
    Set ObjInsulatedVesselLEH = PlaceRevolution(m_OutputColl, oEqpLeftArcI, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVesselLEH
    Set ObjInsulatedVesselLEH = Nothing
    Set oEqpLeftArcI = Nothing

' Insert your code for output 6(InsulatedVesselREH)
'   Revolve Right Elliptical Dome Curve about X-Axis
    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = dInsulationDiameter / 2
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (dInsulationDiameter / 2) '0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpRightArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    Set ObjInsulatedVesselREH = PlaceRevolution(m_OutputColl, oEqpRightArcI, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVesselREH
    Set ObjInsulatedVesselREH = Nothing
    Set geomFactory = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set oEqpRightArcI = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
