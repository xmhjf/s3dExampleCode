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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:24 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Apr 5 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using following five Insulation aspect outputs:
'   i) Insulation for  vessel,
'   ii) Insulation for Boot, and
'   iii) to v) are insulation for nozzle with length
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
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

    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsulatedVessel As Object
    Dim ObjInsulatedBoot As Object
    
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parBootHorLocation As Double
    Dim parBootVertLocation As Double
    Dim parBootDiameter As Double
    Dim parBootAngularLocation As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parThirdSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parStiffenerRadius As Double
    Dim parInsulationThickness As Double

    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parBootHorLocation = arrayOfInputs(4)
    parBootVertLocation = arrayOfInputs(5)
    parBootDiameter = arrayOfInputs(6)
    parBootAngularLocation = arrayOfInputs(7)
    parVesselCenterHeight = arrayOfInputs(8)
    parFirstSupportLocation = arrayOfInputs(9)
    parSecondSupportLocation = arrayOfInputs(10)
    parThirdSupportLocation = arrayOfInputs(11)
    parSupportThickness = arrayOfInputs(12)
    parSupportLength = arrayOfInputs(13)
    parStiffenerRadius = arrayOfInputs(14)
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0

' Insert your code for output 12(InsulatedVessel)
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim majorX As Double, majorY As Double
    Dim majorZ As Double, MMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLineI As IngrGeom3D.Line3d
    Dim oEqpLeftArcI As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArcI As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d
'   Vessel is formed by creating a Complex string and rotating about X-Axis
'   Left Convex Dome Curve
    parInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
    cenX = CenterPos.x
    cenY = CenterPos.y
    cenZ = CenterPos.z
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
'   Horizontal Line
    Dim Linepts(0 To 5) As Double
    Linepts(0) = CenterPos.x
    Linepts(1) = CenterPos.y
    Linepts(2) = CenterPos.z - parInsulationDiameter / 2
    Linepts(3) = CenterPos.x + parVesselLength
    Linepts(4) = CenterPos.y
    Linepts(5) = CenterPos.z - parInsulationDiameter / 2
    Set oEqpLineI = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
'   Right Dome Curve
    cenX = CenterPos.x + parVesselLength
    cenY = CenterPos.y
    cenZ = CenterPos.z
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
                                                            
    Dim EleEqpCollection   As Collection
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    Set EleEqpCollection = New Collection
    EleEqpCollection.Add oEqpLeftArcI
    EleEqpCollection.Add oEqpLineI
    EleEqpCollection.Add oEqpRightArcI
    
    startPoint.Set CenterPos.x - parVesselDiameter / 4 - parInsulationThickness, _
                   CenterPos.y, _
                   CenterPos.z
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleEqpCollection)
    axisVect.Set 1, 0, 0
    revCenPt.Set CenterPos.x + parVesselLength / 2, CenterPos.y, CenterPos.z
'   Revolve it about X-Axis
    Set ObjInsulatedVessel = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing
    
'   Remove the complex string
    Dim ObjVesselCurves As IJDObject
    Set ObjVesselCurves = oEqpComplexStr
    ObjVesselCurves.Remove
    Set ObjVesselCurves = Nothing
    
' Insert your code for output 13(InsulatedBoot)
'   Boot bottom curve
    cenX = CenterPos.x + parBootHorLocation
    cenY = CenterPos.y + parBootVertLocation * Sin(parBootAngularLocation)
    cenZ = CenterPos.z + parBootVertLocation * Cos(parBootAngularLocation)
    parInsulationDiameter = parBootDiameter + 2 * parInsulationThickness
    majorX = parInsulationDiameter / 2
    majorY = 0
    majorZ = 0
    MMRatio = 0.5
    StartAngle = -0.5 * PI
    SweepAngle = PI / 2
    norX = 0
    norY = Cos(parBootAngularLocation)
    norZ = Sin(parBootAngularLocation)

    Dim oBootBottomArcI As IngrGeom3D.EllipticalArc3d
    Set oBootBottomArcI = New IngrGeom3D.EllipticalArc3d
    Set oBootBottomArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
'   Vertical Line
    Linepts(0) = CenterPos.x + parBootHorLocation + parInsulationDiameter / 2
    Linepts(1) = CenterPos.y + parBootVertLocation * Sin(parBootAngularLocation)
    Linepts(2) = CenterPos.z + parBootVertLocation * Cos(parBootAngularLocation)
    Linepts(3) = CenterPos.x + parBootHorLocation + parInsulationDiameter / 2
    Linepts(4) = CenterPos.y
    Linepts(5) = CenterPos.z
    Dim oBootLineI As IngrGeom3D.Line3d
    Set oBootLineI = New IngrGeom3D.Line3d
    Set oBootLineI = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set geomFactory = Nothing
'   Prepare complex string
    Dim EleBootCollection   As Collection
    Set EleBootCollection = New Collection
    EleBootCollection.Add oBootBottomArcI
    EleBootCollection.Add oBootLineI

    startPoint.Set CenterPos.x + parBootHorLocation, _
                   CenterPos.y + (parBootVertLocation + parBootDiameter / 4) * Sin(parBootAngularLocation), _
                   CenterPos.z + (parBootVertLocation + parBootDiameter / 4 + parInsulationThickness) * _
                                        Cos(parBootAngularLocation)
    Dim oBootComplexStr As IngrGeom3D.ComplexString3d
    Set oBootComplexStr = New IngrGeom3D.ComplexString3d
    Set oBootComplexStr = PlaceTrCString(startPoint, EleBootCollection)
    axisVect.Set 0, -Sin(parBootAngularLocation), -Cos(parBootAngularLocation)
    revCenPt.Set CenterPos.x + parBootHorLocation, _
                 CenterPos.y + (parBootVertLocation + parBootDiameter / 4) * Sin(parBootAngularLocation), _
                 CenterPos.z + (parBootVertLocation + parBootDiameter / 4) * Cos(parBootAngularLocation)
    Set ObjInsulatedBoot = PlaceRevolution(m_OutputColl, oBootComplexStr, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBoot
    Set ObjInsulatedBoot = Nothing
    
'   Remove the complex string
    Dim ObjBootCurves As IJDObject
    Set ObjBootCurves = oBootComplexStr
    ObjBootCurves.Remove
    Set ObjBootCurves = Nothing
    
  
    Dim iCount As Integer
    For iCount = 1 To EleEqpCollection.Count
        EleEqpCollection.Remove 1
    Next iCount
    Set EleEqpCollection = Nothing

    For iCount = 1 To EleBootCollection.Count
        EleBootCollection.Remove 1
    Next iCount
    Set EleBootCollection = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
