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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:23 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Apr 5 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using the following Outputs:
'   Five Insulation aspect outputs, and
'   14 Physical outputs viz. Vessel, Boot, 3 supports, 3 stiffeneres, 3 defaultsurfaces
'   and 3 Nozzles placed with length
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
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
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjVessel As Object
    Dim ObjBoot As Object
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
    Dim ObjSupport3 As Object
    Dim ObjStiffner1 As Object
    Dim ObjStiffner2 As Object
    Dim ObjStiffner3 As Object
    Dim ObjDefaultSurface1 As Object
    Dim ObjDefaultSurface2 As Object
    Dim ObjDefaultSurface3 As Object
    
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
    

' Insert your code for output 1(Vessel)
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim majorX As Double, majorY As Double
    Dim majorZ As Double, MMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpLeftArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d
'   Vessel is formed by creating a Complex string and rotating about X-Axis
'   Left Convex Dome Curve

    cenX = CenterPos.x
    cenY = CenterPos.y
    cenZ = CenterPos.z
    majorX = 0
    majorY = 0
    majorZ = parVesselDiameter / 2
    MMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set oEqpLeftArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
'   Horizontal Line
    Dim Linepts(0 To 5) As Double
    Linepts(0) = CenterPos.x
    Linepts(1) = CenterPos.y
    Linepts(2) = CenterPos.z - parVesselDiameter / 2
    Linepts(3) = CenterPos.x + parVesselLength
    Linepts(4) = CenterPos.y
    Linepts(5) = CenterPos.z - parVesselDiameter / 2
    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
'   Right Dome Curve
    cenX = CenterPos.x + parVesselLength
    cenY = CenterPos.y
    cenZ = CenterPos.z
    majorX = 0
    majorY = 0
    majorZ = parVesselDiameter / 2
    MMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpRightArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim EleEqpCollection   As Collection
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim VesselRevCenPt As New AutoMath.DPosition
    
    Set EleEqpCollection = New Collection
    EleEqpCollection.Add oEqpLeftArc
    EleEqpCollection.Add oEqpLine
    EleEqpCollection.Add oEqpRightArc
    
    startPoint.Set CenterPos.x - parVesselDiameter / 4, CenterPos.y, CenterPos.z
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleEqpCollection)
    axisVect.Set 1, 0, 0
    VesselRevCenPt.Set CenterPos.x + parVesselLength / 2, _
                 CenterPos.y, _
                 CenterPos.z
'   Revolve it about X-Axis
    Set ObjVessel = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing

'   Remove the complex string
    Dim ObjVesselCurves As IJDObject
    Set ObjVesselCurves = oEqpComplexStr
    ObjVesselCurves.Remove
    Set ObjVesselCurves = Nothing

' Insert your code for output 2(Boot)
'   Boot bottom curve
    cenX = CenterPos.x + parBootHorLocation
    cenY = CenterPos.y + parBootVertLocation * Sin(parBootAngularLocation)
    cenZ = CenterPos.z + parBootVertLocation * Cos(parBootAngularLocation)
    majorX = parBootDiameter / 2
    majorY = 0
    majorZ = 0
    MMRatio = 0.5
    StartAngle = -0.5 * PI
    SweepAngle = PI / 2
    norX = 0
    norY = Cos(parBootAngularLocation)
    norZ = Sin(parBootAngularLocation)

    Dim oBootBottomArc As IngrGeom3D.EllipticalArc3d
    Set oBootBottomArc = New IngrGeom3D.EllipticalArc3d
    Set oBootBottomArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, majorX, majorY, majorZ, MMRatio, _
                                                        StartAngle, SweepAngle)
'   Vertical Line
    Linepts(0) = CenterPos.x + parBootHorLocation + parBootDiameter / 2
    Linepts(1) = CenterPos.y + parBootVertLocation * Sin(parBootAngularLocation)
    Linepts(2) = CenterPos.z + parBootVertLocation * Cos(parBootAngularLocation)
    Linepts(3) = CenterPos.x + parBootHorLocation + parBootDiameter / 2
    Linepts(4) = CenterPos.y
    Linepts(5) = CenterPos.z
    Dim oBootLine As IngrGeom3D.Line3d
    Set oBootLine = New IngrGeom3D.Line3d
    Set oBootLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Dim EleBootCollection   As Collection
    Set EleBootCollection = New Collection
    EleBootCollection.Add oBootBottomArc
    EleBootCollection.Add oBootLine

    startPoint.Set CenterPos.x + parBootHorLocation, _
                   CenterPos.y + (parBootVertLocation + parBootDiameter / 4) * Sin(parBootAngularLocation), _
                   CenterPos.z + (parBootVertLocation + parBootDiameter / 4) * Cos(parBootAngularLocation)

    Dim oBootComplexStr As IngrGeom3D.ComplexString3d
    Set oBootComplexStr = New IngrGeom3D.ComplexString3d
    Set oBootComplexStr = PlaceTrCString(startPoint, EleBootCollection)
    axisVect.Set 0, -Sin(parBootAngularLocation), -Cos(parBootAngularLocation)
    Dim BootRevCenPt As New AutoMath.DPosition
    BootRevCenPt.Set CenterPos.x + parBootHorLocation, _
                 CenterPos.y + (parBootVertLocation + parBootDiameter / 4) * Sin(parBootAngularLocation), _
                 CenterPos.z + (parBootVertLocation + parBootDiameter / 4) * Cos(parBootAngularLocation)
    Set ObjBoot = PlaceRevolution(m_OutputColl, oBootComplexStr, axisVect, BootRevCenPt, 2 * PI, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoot
    Set ObjBoot = Nothing
    
'   Remove the complex string
    Dim ObjBootCurves As IJDObject
    Set ObjBootCurves = oBootComplexStr
    ObjBootCurves.Remove
    Set ObjBootCurves = Nothing

' Insert your code for output 3(Support1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set CenterPos.x + parFirstSupportLocation - parSupportThickness / 2, _
                CenterPos.y - parSupportLength / 2, _
                CenterPos.z - parVesselCenterHeight
    enPoint.Set stPoint.x + parSupportThickness, _
                stPoint.y + parSupportLength, _
                stPoint.z + parVesselCenterHeight
    Set ObjSupport1 = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1
    Set ObjSupport1 = Nothing
    
' Insert your code for output 4(Support2)
    stPoint.Set CenterPos.x + parFirstSupportLocation + parSecondSupportLocation _
                        - parSupportThickness / 2, _
                CenterPos.y - parSupportLength / 2, _
                CenterPos.z - parVesselCenterHeight
    enPoint.Set stPoint.x + parSupportThickness, _
                stPoint.y + parSupportLength, _
                stPoint.z + parVesselCenterHeight
    Set ObjSupport2 = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2
    Set ObjSupport2 = Nothing
    
' Insert your code for output 5(Support3)
    stPoint.Set CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation - parSupportThickness / 2, _
                CenterPos.y - parSupportLength / 2, _
                CenterPos.z - parVesselCenterHeight
    enPoint.Set stPoint.x + parSupportThickness, _
                stPoint.y + parSupportLength, _
                stPoint.z + parVesselCenterHeight
    Set ObjSupport3 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport3
    Set ObjSupport3 = Nothing
    
' Insert your code for output 6(Stiffner1)
    Dim StiffnerPoints(0 To 14)  As Double
    
'   Consider a section of Stiffener in XZ plane
'   Left hand side Bottom point
    StiffnerPoints(0) = CenterPos.x + parFirstSupportLocation - parSupportThickness / 2
    StiffnerPoints(1) = CenterPos.y
    StiffnerPoints(2) = CenterPos.z + parVesselDiameter / 2

'   Right hand side Bottom point
    StiffnerPoints(3) = CenterPos.x + parFirstSupportLocation + parSupportThickness / 2
    StiffnerPoints(4) = CenterPos.y
    StiffnerPoints(5) = CenterPos.z + parVesselDiameter / 2
    
'   Right hand side Top point
    StiffnerPoints(6) = CenterPos.x + parFirstSupportLocation + parSupportThickness / 2
    StiffnerPoints(7) = CenterPos.y
    StiffnerPoints(8) = CenterPos.z + parStiffenerRadius
        
'   Left hand side Top point
    StiffnerPoints(9) = CenterPos.x + parFirstSupportLocation - parSupportThickness / 2
    StiffnerPoints(10) = CenterPos.y
    StiffnerPoints(11) = CenterPos.z + parStiffenerRadius
    
'   Left hand side Bottom point
    StiffnerPoints(12) = StiffnerPoints(0)
    StiffnerPoints(13) = StiffnerPoints(1)
    StiffnerPoints(14) = StiffnerPoints(2)
        
'   Prepare profile
    axisVect.Set 1, 0, 0
    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = New IngrGeom3D.LineString3d
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
    Set ObjStiffner1 = PlaceRevolution(m_OutputColl, oLineString, axisVect, VesselRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffner1
    Set ObjStiffner1 = Nothing

'   Remove  linestring
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
' Insert your code for output 7(Stiffner2)
'   Left hand side Bottom point
    StiffnerPoints(0) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation - _
                            parSupportThickness / 2
    StiffnerPoints(1) = CenterPos.y
    StiffnerPoints(2) = CenterPos.z + parVesselDiameter / 2

'   Right hand side Bottom point
    StiffnerPoints(3) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parSupportThickness / 2
    StiffnerPoints(4) = CenterPos.y
    StiffnerPoints(5) = CenterPos.z + parVesselDiameter / 2
    
'   Right hand side Top point
    StiffnerPoints(6) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parSupportThickness / 2
    StiffnerPoints(7) = CenterPos.y
    StiffnerPoints(8) = CenterPos.z + parStiffenerRadius
        
'   Left hand side Top point
    StiffnerPoints(9) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation - _
                            parSupportThickness / 2
    StiffnerPoints(10) = CenterPos.y
    StiffnerPoints(11) = CenterPos.z + parStiffenerRadius
    
'   Left hand side Bottom point
    StiffnerPoints(12) = StiffnerPoints(0)
    StiffnerPoints(13) = StiffnerPoints(1)
    StiffnerPoints(14) = StiffnerPoints(2)

    Set oLineString = New IngrGeom3D.LineString3d
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
    Set ObjStiffner2 = PlaceRevolution(m_OutputColl, oLineString, axisVect, VesselRevCenPt, 2 * PI, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffner2
    Set ObjStiffner2 = Nothing

'   Remove  linestring
'    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    
' Insert your code for output 8(Stiffner3)
'   Left hand side Bottom point
    StiffnerPoints(0) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation - parSupportThickness / 2
    StiffnerPoints(1) = CenterPos.y
    StiffnerPoints(2) = CenterPos.z + parVesselDiameter / 2

'   Right hand side Bottom point
    StiffnerPoints(3) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation + parSupportThickness / 2
    StiffnerPoints(4) = CenterPos.y
    StiffnerPoints(5) = CenterPos.z + parVesselDiameter / 2
    
'   Right hand side Top point
    StiffnerPoints(6) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation + parSupportThickness / 2
    StiffnerPoints(7) = CenterPos.y
    StiffnerPoints(8) = CenterPos.z + parStiffenerRadius
        
'   Left hand side Top point
    StiffnerPoints(9) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation - parSupportThickness / 2
    StiffnerPoints(10) = CenterPos.y
    StiffnerPoints(11) = CenterPos.z + parStiffenerRadius
    
'   Left hand side Bottom point
    StiffnerPoints(12) = StiffnerPoints(0)
    StiffnerPoints(13) = StiffnerPoints(1)
    StiffnerPoints(14) = StiffnerPoints(2)

    Set oLineString = New IngrGeom3D.LineString3d
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
    Set ObjStiffner3 = PlaceRevolution(m_OutputColl, oLineString, axisVect, VesselRevCenPt, 2 * PI, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffner3
    Set ObjStiffner3 = Nothing
    
'   Remove  linestring
'    Dim objLineString As IJDObject
    Set objLineString = oLineString
    Set oLineString = Nothing
    objLineString.Remove
    Set objLineString = Nothing
    
' Insert your code for output 9(DefaultSurface1)
    Dim lineStrPoints(0 To 11)  As Double
    lineStrPoints(0) = CenterPos.x + parFirstSupportLocation - parSupportThickness / 2
    lineStrPoints(1) = CenterPos.y - parSupportLength / 2
    lineStrPoints(2) = CenterPos.z - parVesselCenterHeight

    lineStrPoints(3) = lineStrPoints(0) + parSupportThickness
    lineStrPoints(4) = lineStrPoints(1)
    lineStrPoints(5) = lineStrPoints(2)
    
    lineStrPoints(6) = lineStrPoints(3)
    lineStrPoints(7) = lineStrPoints(4) + parSupportLength
    lineStrPoints(8) = lineStrPoints(5)
    
    lineStrPoints(9) = lineStrPoints(6) - parSupportThickness
    lineStrPoints(10) = lineStrPoints(7)
    lineStrPoints(11) = lineStrPoints(8)

    Set ObjDefaultSurface1 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface1
    Set ObjDefaultSurface1 = Nothing
    
' Insert your code for output 10(DefaultSurface2)
    lineStrPoints(0) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation _
                        - parSupportThickness / 2
    lineStrPoints(1) = CenterPos.y - parSupportLength / 2
    lineStrPoints(2) = CenterPos.z - parVesselCenterHeight

    lineStrPoints(3) = lineStrPoints(0) + parSupportThickness
    lineStrPoints(4) = lineStrPoints(1)
    lineStrPoints(5) = lineStrPoints(2)
    
    lineStrPoints(6) = lineStrPoints(3)
    lineStrPoints(7) = lineStrPoints(4) + parSupportLength
    lineStrPoints(8) = lineStrPoints(5)
    
    lineStrPoints(9) = lineStrPoints(6) - parSupportThickness
    lineStrPoints(10) = lineStrPoints(7)
    lineStrPoints(11) = lineStrPoints(8)

    Set ObjDefaultSurface2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface2
    Set ObjDefaultSurface2 = Nothing

' Insert your code for output 11(DefaultSurface3)
    lineStrPoints(0) = CenterPos.x + parFirstSupportLocation + parSecondSupportLocation + _
                            parThirdSupportLocation - parSupportThickness / 2
    lineStrPoints(1) = CenterPos.y - parSupportLength / 2
    lineStrPoints(2) = CenterPos.z - parVesselCenterHeight

    lineStrPoints(3) = lineStrPoints(0) + parSupportThickness
    lineStrPoints(4) = lineStrPoints(1)
    lineStrPoints(5) = lineStrPoints(2)
    
    lineStrPoints(6) = lineStrPoints(3)
    lineStrPoints(7) = lineStrPoints(4) + parSupportLength
    lineStrPoints(8) = lineStrPoints(5)
    
    lineStrPoints(9) = lineStrPoints(6) - parSupportThickness
    lineStrPoints(10) = lineStrPoints(7)
    lineStrPoints(11) = lineStrPoints(8)

    Set ObjDefaultSurface3 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface3
    Set ObjDefaultSurface3 = Nothing
    Set geomFactory = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To EleEqpCollection.Count
        EleEqpCollection.Remove 1
    Next iCount
    Set EleEqpCollection = Nothing

    For iCount = 1 To EleBootCollection.Count
        EleBootCollection.Remove 1
    Next iCount
    Set EleBootCollection = Nothing

'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "ComplexHorCylVesselControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

