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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          KKC
'   Creation Date:  Tuesday, June 19 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(SDG1.pdf)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   8-Nov-07        KKC/RRK     DI-130618: Changed the parameter names of parPositionerOrientation and parHandWheelOrientation
'                                          to parPositionerRotAngle and parHandWheelRotAngle respectively
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
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
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjActuatorBody As Object
    Dim ObjPositionerBody As Object
    Dim ObjHandwheelStem As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parHandWheelDiameter As Double
    Dim parPositionerHeight As Double
    Dim parPositionerOffset As Double
    Dim parPositionerOffset1 As Double
    Dim parPositionerOffset2 As Double
    Dim parValveOutertoHandWheel As Double
    Dim parValveOutertoCenter As Double
    Dim parPositionerLength As Double
    Dim parPositionerRotAngle As Double
    Dim parInsulationThickness As Double
    Dim parHandWheelRotAngle As Double
    
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
    
    Dim parOperationHeight As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'V1
    parFacetoFace = arrayOfInputs(3)     'V2
    parCylHeight = arrayOfInputs(4)     'V3
    parActuatorHeight = arrayOfInputs(5)    'A1
    parActuatorDiameter = arrayOfInputs(6)  'A2
    parHandWheelOffset = arrayOfInputs(7)   'H2
    parHandWheelDiameter = arrayOfInputs(8) 'H3
    parPositionerHeight = arrayOfInputs(9)      'P4
    parPositionerOffset = arrayOfInputs(10)      'P3
    parPositionerOffset1 = arrayOfInputs(11)     'P6
    parPositionerOffset2 = arrayOfInputs(12)     'P5
    parValveOutertoHandWheel = arrayOfInputs(13) 'H4
    parValveOutertoCenter = arrayOfInputs(14)       'V4
    parPositionerLength = arrayOfInputs(15)     'P2
    parPositionerRotAngle = arrayOfInputs(16)             'P1 'Rotation of Operator
    parInsulationThickness = arrayOfInputs(17)
    'Input parameters for Dynamic Nozzles
    parNPD = arrayOfInputs(18)
    parEndPreparation = arrayOfInputs(19)
    parScheduleThickness = arrayOfInputs(20)
    parEndStandard = arrayOfInputs(21)
    parPressureRating = arrayOfInputs(22)
    parFlowDirection = arrayOfInputs(23)
    parPortIndex1 = arrayOfInputs(24)
    parNpd1 = arrayOfInputs(25)
    parEndPreparation1 = arrayOfInputs(26)
    parScheduleThickness1 = arrayOfInputs(27)
    parEndStandard1 = arrayOfInputs(28)
    parPressureRating1 = arrayOfInputs(29)
    parFlowDirection1 = arrayOfInputs(30)
    parPortIndex2 = arrayOfInputs(31)
    parNpd2 = arrayOfInputs(32)
    parEndPreparation2 = arrayOfInputs(33)
    parScheduleThickness2 = arrayOfInputs(34)
    parEndStandard2 = arrayOfInputs(35)
    parPressureRating2 = arrayOfInputs(36)
    parFlowDirection2 = arrayOfInputs(37)
    parOperationHeight = arrayOfInputs(38)     'I1
    parOperationDiameter = arrayOfInputs(39) 'I2
    parOperationDiameter1 = arrayOfInputs(40) 'I3
    parHandWheelRotAngle = arrayOfInputs(41)          'H1
    
'   Text Inputs
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parNPDUnitType = arrayOfInputs(44)
    parNPDUnitType1 = arrayOfInputs(45)
    parNPDUnitType2 = arrayOfInputs(46)

    iOutput = 0
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Dim oPlacePoint As AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    
'CreateRetrieveDynamicNozzle
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
    
    If CmpDblEqual(parValveOutertoCenter, 0) Then parValveOutertoCenter = 0.00001
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2

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

' Place Nozzle 2
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
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    
 'The HandWheel stem Dia and HandWheel thickness are taken to be 0.05
    Dim StemDia As Double
    Dim HandWheelThk As Double
    StemDia = 0.05
    HandWheelThk = 0.05
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
' Insert your code for output 3(Body Cone1 of Valve)
    oStPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0

    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 4(Body Cone2 of Valve)
    
    oStPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0

    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 5(Valve Cylinder)
    Dim CylDia As Double
    oStPoint.Set 0, -parCylHeight, 0
    oEnPoint.Set 0, parValveOutertoCenter, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
      
 ' Insert your code for output 6(Actuator Body)
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    
    Dim oAxisVect As New AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 11) As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim dSphereHeight As Double
    dSphereHeight = 0.2 * parActuatorHeight
'Create a lineString and Revolve it
 'Point 1
    LineStringPoints(0) = 0
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parActuatorDiameter / 2
    LineStringPoints(4) = parValveOutertoCenter + 0.8 * parActuatorHeight
    LineStringPoints(5) = 0

'Point 3
    LineStringPoints(6) = 0
    LineStringPoints(7) = parValveOutertoCenter + 0.8 * parActuatorHeight
    LineStringPoints(8) = 0

    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 3, LineStringPoints)
    oAxisVect.Set 0, 1, 0
    oCenterPoint.Set 0, (parValveOutertoCenter + 0.8 * parActuatorHeight) / 2, 0
    Set ObjActuatorBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    Set oCenterPoint = Nothing
    Set oLineString = Nothing
 ' Insert your code for output 7(Top Cap)
    Dim oRevVect As AutoMath.DVector
    Dim oEllipseCenter As AutoMath.DPosition
    Dim oEllipseNormal As AutoMath.DPosition
    Dim oMajorAxis As AutoMath.DPosition
    Dim oEllipse As Object
    Dim dAxesRatio As Double
    Dim ObjEndCap As Object
    
    Set oMajorAxis = New AutoMath.DPosition
    Set oEllipseNormal = New AutoMath.DPosition
    Set oEllipseCenter = New AutoMath.DPosition
    Set oRevVect = New AutoMath.DVector
    Set oCenterPoint = New AutoMath.DPosition
    
    oEllipseCenter.Set 0, parValveOutertoCenter + 0.8 * parActuatorHeight, 0
    oMajorAxis.Set parActuatorDiameter / 2, 0, 0
    oEllipseNormal.Set 0, 0, -1
    
    dAxesRatio = 0.4 * (parActuatorHeight / parActuatorDiameter)
    Set oEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                oEllipseCenter.x, oEllipseCenter.y, oEllipseCenter.z, _
                                oEllipseNormal.x, oEllipseNormal.y, oEllipseNormal.z, _
                                oMajorAxis.x, oMajorAxis.y, oMajorAxis.z, _
                                dAxesRatio, 1.5 * PI, PI / 2)
    oRevVect.Set 0, 1, 0
    oCenterPoint.Set 0, parValveOutertoCenter + 0.8 * parActuatorHeight, 0
    Set ObjEndCap = PlaceRevolution(m_OutputColl, oEllipse, oRevVect, oCenterPoint, PI * 2, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndCap
    Set ObjEndCap = Nothing
    
 ' Insert your code for output 7(Positioner Body)
'Create a line string and project it
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d

'Point 1
    LineStrPts(0) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(1) = parPositionerOffset2 - parPositionerHeight
    LineStrPts(2) = parPositionerOffset1
'Point 2
    LineStrPts(3) = parPositionerOffset
    LineStrPts(4) = parPositionerOffset2 - parPositionerHeight
    LineStrPts(5) = parPositionerOffset1
'Point 3
    LineStrPts(6) = parPositionerOffset
    LineStrPts(7) = parPositionerOffset2
    LineStrPts(8) = parPositionerOffset1
'Point 4
    LineStrPts(9) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(10) = parPositionerOffset2
    LineStrPts(11) = parPositionerOffset1
'Point 5
    LineStrPts(12) = LineStrPts(0)
    LineStrPts(13) = LineStrPts(1)
    LineStrPts(14) = LineStrPts(2)
    
    Dim dProjLength As Double
    dProjLength = parPositionerOffset1 - (((parActuatorDiameter / 2) * (parPositionerOffset2 - parPositionerHeight)) _
                            / (parValveOutertoCenter + parActuatorHeight * 0.8))

    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
     
    oAxisVect.Set 0, 0, -1
    Set ObjPositionerBody = PlaceProjection(m_OutputColl, oLString, oAxisVect, dProjLength, True)
    
'Transform the output  (Rotation about stem axis)
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parPositionerRotAngle, oAboutVec
    ObjPositionerBody.Transform oTransMat
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositionerBody
    Set ObjPositionerBody = Nothing
    Set oTransMat = Nothing
    
' Insert your code for output 8(HandWheel)
   '   Place Torus
    Dim ObjHandwheel    As IngrGeom3D.Torus3d
    Dim TorSecRad As Double
    Dim TorPriRad As Double
    Dim oTorOrigin As AutoMath.DPosition
    Set oTorOrigin = New AutoMath.DPosition
    Dim oTorAxis As AutoMath.DVector
    Set oTorAxis = New AutoMath.DVector
    TorSecRad = HandWheelThk / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    oTorOrigin.Set 0, parValveOutertoCenter + parValveOutertoHandWheel, _
                                -parHandWheelOffset + HandWheelThk / 2
    oTorAxis.Set 0, 0, 1
    Set ObjHandwheel = PlaceTorus(m_OutputColl, oTorOrigin, oTorAxis, TorPriRad, TorSecRad)
    
    'Transform the output  (Rotation about stem axis)
    Set oTransMat = New AutoMath.DT4x4
    oTransMat.Rotate -parHandWheelRotAngle, oAboutVec
    ObjHandwheel.Transform oTransMat

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set oTorOrigin = Nothing
    Set oTorAxis = Nothing
    
 ' Insert your code for output 9(HandwheelStem)
    Dim dMeetRadius As Double
    dMeetRadius = ((parActuatorDiameter / 2) / (parValveOutertoCenter + parActuatorHeight * 0.8)) * _
                                (parValveOutertoCenter + parValveOutertoHandWheel - StemDia / 2)
    oStPoint.Set 0, parValveOutertoCenter + parValveOutertoHandWheel, -dMeetRadius
    oEnPoint.Set 0, parValveOutertoCenter + parValveOutertoHandWheel, -parHandWheelOffset
    Set ObjHandwheelStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, StemDia, True)
    
    'Transform the output  (Rotation about stem axis)
    ObjHandwheelStem.Transform oTransMat
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheelStem
    Set ObjHandwheelStem = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxisVect = Nothing
    Set oGeomFactory = Nothing
    Set oLString = Nothing
    Set oTransMat = Nothing
    Set oAboutVec = Nothing
    Set oEllipse = Nothing
       
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
