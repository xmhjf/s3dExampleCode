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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Wednesday, Oct 9 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDMS Category Name=IVJU).
'    this Symbol is same as (PDMS Category Name=IVJS).'   This symbol is Custom INtools Instrument with
'   Nozzles are fully Parametric.Physical aspect consist 2 cones, 2 Cylinders of stem Dia, 2 Cylinder , 1 box,
'   1 torus and a cone with elliptical dome.  2 Nozzles. Insulation Aspect consist of a simple Cylinder for the valve Body.
'   Operation Aspect consist of a simple Cylinder
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   28.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
    PI = 4 * Atn(1)

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    
    Dim iOutput     As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjStem As Object
    Dim ObjBox As Object
    Dim ObjCylinder As Object
    Dim ObjStem1 As Object
    Dim ObjHandWheel As Object
    Dim ObjActuatorCone As Object
    
    Dim parFace1toCenter As Double         'D2
    Dim parFacetoFace     As Double     'D3
    Dim parCylHeight      As Double     'D4
    Dim parActuatorWidth     As Double      'D5
    Dim parActuatorDiameter     As Double       'D6
    Dim parActuatorHeight     As Double          'D7
    Dim parHandWheelLength    As Double         'D8
    Dim parHandWheelDiameter      As Double         'D9
    Dim parActuatorHeight1     As Double            'D10
    Dim parActuatorHeight2     As Double            'D11
    Dim parActuatorLength     As Double             'D12
    Dim parActuatorLength1    As Double             'D14
    Dim parActuatorHeight3     As Double            'D16
    Dim parActuatorHeight4     As Double            'D19
    Dim parHandWheelOffset   As Double          'D20
    Dim parActuatorWidth1       As Double           'D25
    Dim parActuatorWidth2     As Double             'D26
    Dim parActuatorHeight5     As Double            'D27
    Dim parActuatorLength2  As Double           'D28
    Dim parInsulationThickness As Double
    Dim parOperationDiameter As Double      'D17
    Dim parOperationDiameter1 As Double      'D18
    
    'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parPortIndex2 As Long
    Dim parNpd2 As Double
    Dim parEndPreparation2 As Long
    Dim parScheduleThickness2 As Long
    Dim parEndStandard2 As Long
    Dim parPressureRating2 As Long
    Dim parFlowDirection2 As DistribFlow
    Dim parId2 As String
    Dim parNPDUnitType2 As String
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)         'D2
    parFacetoFace = arrayOfInputs(3)         'D3
    parCylHeight = arrayOfInputs(4)         'D4
    parActuatorWidth = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)    'D6
    parActuatorHeight = arrayOfInputs(7)    'D7
    parHandWheelLength = arrayOfInputs(8)    'D8
    parHandWheelDiameter = arrayOfInputs(9)     'D9
    parActuatorHeight1 = arrayOfInputs(10)    'D10
    parActuatorHeight2 = arrayOfInputs(11)    'D11
    parActuatorLength = arrayOfInputs(12)    'D12
    parActuatorLength1 = arrayOfInputs(13)    'D14
    parActuatorHeight3 = arrayOfInputs(14)    'D16
    parActuatorHeight4 = arrayOfInputs(15)    'D19
    parHandWheelOffset = arrayOfInputs(16)    'D20
    parActuatorWidth1 = arrayOfInputs(17)    'D25
    parActuatorWidth2 = arrayOfInputs(18)    'D26
    parActuatorHeight5 = arrayOfInputs(19)    'D27
    parActuatorLength2 = arrayOfInputs(20)    'D28
    parInsulationThickness = arrayOfInputs(21)
    parNPD = arrayOfInputs(22)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(23)
    parScheduleThickness = arrayOfInputs(24)
    parEndStandard = arrayOfInputs(25)
    parPressureRating = arrayOfInputs(26)
    parFlowDirection = arrayOfInputs(27)
    parPortIndex1 = arrayOfInputs(28)
    parNpd1 = arrayOfInputs(29)
    parEndPreparation1 = arrayOfInputs(30)
    parScheduleThickness1 = arrayOfInputs(31)
    parEndStandard1 = arrayOfInputs(32)
    parPressureRating1 = arrayOfInputs(33)
    parFlowDirection1 = arrayOfInputs(34)
    parPortIndex2 = arrayOfInputs(35)
    parNpd2 = arrayOfInputs(36)
    parEndPreparation2 = arrayOfInputs(37)
    parScheduleThickness2 = arrayOfInputs(38)
    parEndStandard2 = arrayOfInputs(39)
    parPressureRating2 = arrayOfInputs(40)
    parFlowDirection2 = arrayOfInputs(41)
    parOperationDiameter = arrayOfInputs(42)
    parOperationDiameter = arrayOfInputs(43)
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    iOutput = 0
    
    ' Place Nozzle 1 for output 1 (Nozzle1)
        If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

'The origin of the Symbol is assumed at intersection of Stem axis and Nozzle axis
  ' Insert your code for output 1(Body Cone1 of Valve)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Dim stemDia As Double
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
'    ' Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex2, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNPDUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex2)

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Insert your code for output 2(Body Cone2 of Valve)
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
' Insert your code for output 4(Stem)
'Stem Diameter is taken as 0.05
    stPoint.Set 0, 0, 0
    enPoint.Set 0, (parActuatorHeight - parActuatorHeight2), 0           'D7-D11
    stemDia = 0.05
    Set ObjStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 5(Box)
    Dim projDist As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Dim projVector As AutoMath.DVector
    Dim oBoxLineStr As IngrGeom3D.LineString3d
    Dim Boxpts(0 To 14) As Double
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set projVector = New AutoMath.DVector
    
    'Point 1 +ive X-axis and front side -ive z-side
    Boxpts(0) = parActuatorLength2 + parActuatorLength1
    Boxpts(1) = parActuatorHeight + parActuatorHeight1
    Boxpts(2) = parActuatorWidth2
    
    'Point 2 +ive X-axis and back side -ive z-side
    Boxpts(3) = parActuatorLength2 + parActuatorLength1
    Boxpts(4) = parActuatorHeight - parActuatorHeight2
    Boxpts(5) = parActuatorWidth2
    
    'Point 3 +ive X-axis and Back side +ive z-side
    Boxpts(6) = parActuatorLength2 + parActuatorLength1
    Boxpts(7) = parActuatorHeight - parActuatorHeight2
    Boxpts(8) = -parActuatorWidth1
    
    'Point 4 +ive X-axis and front side +ive z-side
    Boxpts(9) = parActuatorLength2 + parActuatorLength1
    Boxpts(10) = parActuatorHeight + parActuatorHeight1
    Boxpts(11) = -parActuatorWidth1
    
    'Point 5 same as Point 1
    Boxpts(12) = parActuatorLength2 + parActuatorLength1
    Boxpts(13) = parActuatorHeight + parActuatorHeight1
    Boxpts(14) = parActuatorWidth2
    Set oBoxLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, Boxpts)
    projDist = parActuatorLength1 + parActuatorLength
    projVector.Set -1, 0, 0
    Set ObjBox = PlaceProjection(m_OutputColl, oBoxLineStr, projVector, projDist, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox

'   Remove line string
    Dim objLineString As IJDObject
    Set objLineString = oBoxLineStr
    Set oBoxLineStr = Nothing
    objLineString.Remove
    Set ObjBox = Nothing
    Set projVector = Nothing
    
 ' Insert your code for output 6(Cylinder)
    Dim cylinRadius As Double
    cylinRadius = Abs((parHandWheelOffset - parActuatorLength2) + stemDia / 2)
    stPoint.Set 0, (parActuatorHeight + parActuatorHeight1), 0
    enPoint.Set 0, (parActuatorHeight4 + parActuatorHeight5), 0

    Set ObjCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, cylinRadius * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    
 ' Insert your code for output 7(Stem1)
    Dim cylinderTk As Double
    cylinderTk = (parActuatorHeight4 + parActuatorHeight5) - (parActuatorHeight + parActuatorHeight1)
    stPoint.Set -(parHandWheelOffset - parActuatorLength2), _
                            parActuatorHeight4 + parActuatorHeight5 - cylinderTk / 2, 0
    enPoint.Set -(parHandWheelOffset - parActuatorLength2), _
                            parActuatorHeight4 + parActuatorHeight5 - cylinderTk / 2, _
                            parHandWheelLength
    Set ObjStem1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, stemDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem1
    Set ObjStem1 = Nothing
    
 ' Insert your code for output 8(HandWheel)
 'hand wheel Diameter is assumed as 0.05
    Dim centerPoint As AutoMath.DPosition
    Dim axisVector As AutoMath.DVector
    Dim majorRadius As Double
    Dim minorRadius As Double
    Dim handwheelCSDia  As Double
    Set centerPoint = New AutoMath.DPosition
    Set axisVector = New AutoMath.DVector
    
    handwheelCSDia = 0.05
    centerPoint.Set -(parHandWheelOffset - parActuatorLength2), _
                            parActuatorHeight4 + parActuatorHeight5 - cylinderTk / 2, _
                            parHandWheelLength - handwheelCSDia / 2
    axisVector.Set 0, 0, 1
    majorRadius = (parHandWheelDiameter - handwheelCSDia) / 2
    minorRadius = handwheelCSDia / 2
    Set ObjHandWheel = PlaceTorus(m_OutputColl, centerPoint, axisVector, majorRadius, minorRadius)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheel
    Set ObjHandWheel = Nothing
    
 ' Insert your code for output 9(ActuatorCone)
 'height of circular Dome is assumed as 0.2 times of the diameter - parActuatorDiameter
 'Place Elliptical Arc on Right Side
    
    Dim oConeLine As IngrGeom3D.Line3d
    Dim oDomeArc As IngrGeom3D.Arc3d
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim EleCollection           As Collection
    Dim Linepts(0 To 5) As Double
    Dim arcRadius As Double
        
    arcRadius = ((0.2 * parActuatorDiameter) ^ 2 + (parActuatorDiameter / 2) ^ 2) / _
                            (2 * 0.2 * parActuatorDiameter)
    centerPoint.Set parActuatorLength2, _
                            parActuatorHeight, _
                            -parActuatorWidth + arcRadius
    stPoint.Set parActuatorLength2, _
                        parActuatorHeight, _
                        -parActuatorWidth
    enPoint.Set parActuatorLength2 + parActuatorDiameter / 2, _
                        parActuatorHeight, _
                        -parActuatorWidth + 0.2 * parActuatorDiameter
    Set oDomeArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        centerPoint.x, centerPoint.y, centerPoint.z, _
                                                        stPoint.x, stPoint.y, stPoint.z, _
                                                        enPoint.x, enPoint.y, enPoint.z)
    'Side of cone
    Linepts(0) = parActuatorLength2 + parActuatorDiameter / 2
    Linepts(1) = parActuatorHeight
    Linepts(2) = -parActuatorWidth + 0.2 * parActuatorDiameter
    
    Linepts(3) = parActuatorLength2
    Linepts(4) = parActuatorHeight
    Linepts(5) = -parActuatorWidth1
    
    Set oConeLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set EleCollection = New Collection
    EleCollection.Add oDomeArc
    EleCollection.Add oConeLine
    
    Set oComplexStr = PlaceTrCString(stPoint, EleCollection)
    axisVector.Set 0, 0, 1
    centerPoint.Set parActuatorLength2, parActuatorHeight, -(parActuatorWidth + parActuatorWidth1) / 2
    'Revolve it about y-Axiz
    Set ObjActuatorCone = PlaceRevolution(m_OutputColl, oComplexStr, axisVector, centerPoint, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCone
    Set ObjActuatorCone = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oComplexStr
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set oConeLine = Nothing
    Set oDomeArc = Nothing
    
    Dim count As Integer
    For count = 1 To EleCollection.count
        EleCollection.Remove 1
    Next count
    Set EleCollection = Nothing
    Set oComplexStr = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set centerPoint = Nothing
    Set axisVector = Nothing
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub


