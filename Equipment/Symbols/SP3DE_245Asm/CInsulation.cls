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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Apr 12 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using four Insulation aspect outputs which are as follows:
'   i) Insulation for the vessel and ii) to iv) are Insulation for the three nozzles with length.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
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
    Dim ObjInsulatedVessel As Object
    Dim ObjInsulatedVesselLEH As Object
    Dim ObjInsulatedVesselREH As Object

    Dim ObjInsNozzle1WithLength As Object
    Dim ObjInsNozzle2WithLength As Object
    Dim ObjInsNozzle3WithLength As Object
    Dim ObjInsNozzle4WithLength As Object
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parVesselCenterHeight = arrayOfInputs(4)
    parFirstSupportLocation = arrayOfInputs(5)
    parSecondSupportLocation = arrayOfInputs(6)
    parSupportThickness = arrayOfInputs(7)
    parSupportLength = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

' Insert your code for output 1(InsulatedVessel)
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim majorX As Double, majorY As Double
    Dim majorZ As Double, MMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLineI As IngrGeom3D.Line3d
    Dim oEqpLeftArcI As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArcI As IngrGeom3D.EllipticalArc3d
'   Vessel is formed by creating a Cylinder and two Elliptical Heads rotating about X-Axis
'   Left Elliptical Dome Curve
    parInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    cenX = 0
    cenY = 0
    cenZ = 0
    majorX = 0
    majorY = 0
    majorZ = parInsulationDiameter / 2
    MMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set oEqpLeftArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
'   Right Elliptical Dome Curve
    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    majorX = 0
    majorY = 0
    majorZ = parInsulationDiameter / 2
    MMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpRightArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim startPoint As New AutoMath.DPosition
    Dim EndPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    
    startPoint.Set 0, 0, 0
    EndPoint.Set parVesselLength, 0, 0

    axisVect.Set 1, 0, 0
    revCenPt.Set parVesselLength / 2, 0, 0
'   Create Insulated Horizontal Drum
    Set ObjInsulatedVessel = PlaceCylinder(m_OutputColl, startPoint, EndPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing
    
'   Revolve LEH about X-Axis
    Set ObjInsulatedVesselLEH = PlaceRevolution(m_OutputColl, oEqpLeftArcI, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVesselLEH
    Set ObjInsulatedVesselLEH = Nothing
    
'   Revolve REH about X-Axis
    Set ObjInsulatedVesselREH = PlaceRevolution(m_OutputColl, oEqpRightArcI, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVesselREH
    Set ObjInsulatedVesselREH = Nothing
    Set geomFactory = Nothing
    Set startPoint = Nothing
    Set EndPoint = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing

        
    Set stPoint = Nothing
    Set enPoint = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
