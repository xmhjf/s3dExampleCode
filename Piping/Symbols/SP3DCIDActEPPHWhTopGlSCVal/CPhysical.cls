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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Friday, 3rd Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components( PDMS Category Name=IVHH)
'   The symbol consists of 11 outputs: Nine Physical outputs , One Insulation Output and one Operation
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  08.Jun.2008     KKC  CR-137825  Piping instrument catalog should be updated to current Intools Catalog
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
    Dim parValveOutertoCenter As Double
    Dim parPositionerLength As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double


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
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFacetoFace = arrayOfInputs(3)     'D3
    parCylHeight = arrayOfInputs(4)     'D4
    parActuatorHeight = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)  'D6
    parHandWheelOffset = arrayOfInputs(7)   'D8
    parHandWheelDiameter = arrayOfInputs(8) 'D9
    parPositionerHeight = arrayOfInputs(9)      'D10
    parPositionerOffset = arrayOfInputs(10)      'D11
    parPositionerOffset1 = arrayOfInputs(11)     'D12
    parPositionerOffset2 = arrayOfInputs(12)     'D14
    parValveOutertoCenter = arrayOfInputs(13)       'D19
    parPositionerLength = arrayOfInputs(14)     'D20
    parInsulationThickness = arrayOfInputs(15)
    'Input parameters for Dynamic Nozzles
    parNPD = arrayOfInputs(16)
    parEndPreparation = arrayOfInputs(17)
    parScheduleThickness = arrayOfInputs(18)
    parEndStandard = arrayOfInputs(19)
    parPressureRating = arrayOfInputs(20)
    parFlowDirection = arrayOfInputs(21)
    parPortIndex1 = arrayOfInputs(22)
    parNpd1 = arrayOfInputs(23)
    parEndPreparation1 = arrayOfInputs(24)
    parScheduleThickness1 = arrayOfInputs(25)
    parEndStandard1 = arrayOfInputs(26)
    parPressureRating1 = arrayOfInputs(27)
    parFlowDirection1 = arrayOfInputs(28)
    parPortIndex2 = arrayOfInputs(29)
    parNpd2 = arrayOfInputs(30)
    parEndPreparation2 = arrayOfInputs(31)
    parScheduleThickness2 = arrayOfInputs(32)
    parEndStandard2 = arrayOfInputs(33)
    parPressureRating2 = arrayOfInputs(34)
    parFlowDirection2 = arrayOfInputs(35)
    parOperationHeight = arrayOfInputs(36)     'D7
    parOperationDiameter = arrayOfInputs(37) 'D17
    parOperationDiameter1 = arrayOfInputs(38) 'D18
    
    parId1 = arrayOfInputs(39)
    parId2 = arrayOfInputs(40)
    parNPDUnitType = arrayOfInputs(41)
    parNPDUnitType1 = arrayOfInputs(42)
    parNPDUnitType2 = arrayOfInputs(43)
   
    If UBound(arrayOfInputs) > 43 Then
        parAngle = arrayOfInputs(44)
    End If
    iOutput = 0
    
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
        ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
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
    Dim stPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
' Insert your code for output 3(Body Cone1 of Valve)
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 4(Body Cone2 of Valve)
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 5(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
      
 ' Insert your code for output 6(Actuator Body)
     Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 8) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
'Create a lineString and Revolve it
 'Point 1
    LineStringPoints(0) = 0
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
'Point 2
    LineStringPoints(3) = parActuatorDiameter / 2
    LineStringPoints(4) = parValveOutertoCenter + parActuatorHeight
    LineStringPoints(5) = 0
'Point 3
    LineStringPoints(6) = 0
    LineStringPoints(7) = parValveOutertoCenter + parActuatorHeight
    LineStringPoints(8) = 0

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 3, LineStringPoints)
    axisVect.Set 0, 1, 0
    centerPoint.Set 0, (parValveOutertoCenter + parActuatorHeight) / 2, 0
    Set ObjActuatorBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    Set centerPoint = Nothing
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
    
 ' Insert your code for output 7(Positioner Body)
'Create a line string and project it
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d

'Point 1
    LineStrPts(0) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(1) = parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight
    LineStrPts(2) = parPositionerOffset1
'Point 2
    LineStrPts(3) = parPositionerOffset
    LineStrPts(4) = parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight
    LineStrPts(5) = parPositionerOffset1
'Point 3
    LineStrPts(6) = parPositionerOffset
    LineStrPts(7) = parValveOutertoCenter + parPositionerOffset2
    LineStrPts(8) = parPositionerOffset1
'Point 4
    LineStrPts(9) = -(parPositionerLength - parPositionerOffset)
    LineStrPts(10) = parValveOutertoCenter + parPositionerOffset2
    LineStrPts(11) = parPositionerOffset1
'Point 5
    LineStrPts(12) = LineStrPts(0)
    LineStrPts(13) = LineStrPts(1)
    LineStrPts(14) = LineStrPts(2)
    
    Dim ProjLength As Double
    ProjLength = parPositionerOffset1 - (((parActuatorDiameter / 2) * (parValveOutertoCenter + parPositionerOffset2 - parPositionerHeight)) _
                            / (parValveOutertoCenter + parActuatorHeight))

    Set oLString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
     
    axisVect.Set 0, 0, -1
    Set ObjPositionerBody = PlaceProjection(m_OutputColl, oLString, axisVect, ProjLength, True)
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parAngle, oAxisVect
    ObjPositionerBody.Transform oTransMatrix
       
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositionerBody
    Set ObjPositionerBody = Nothing
    
' Insert your code for output 8(HandWheel)
   '   Place Torus
    Dim ObjHandwheel    As IngrGeom3D.Torus3d
    Dim TorSecRad As Double
    Dim TorPriRad As Double
    Dim TorOrigin       As New AutoMath.DPosition
    Dim TorAxis         As New AutoMath.DVector
    TorSecRad = HandWheelThk / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    TorOrigin.Set 0, parValveOutertoCenter + parActuatorHeight + parHandWheelOffset - HandWheelThk / 2, 0
    TorAxis.Set 0, 1, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
    
 ' Insert your code for output 9(HandwheelStem)
    stPoint.Set 0, parValveOutertoCenter + parActuatorHeight, 0
    enPoint.Set 0, parValveOutertoCenter + parActuatorHeight + parHandWheelOffset, 0
    Set ObjHandwheelStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheelStem
    Set ObjHandwheelStem = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = oLString
    sp3dElem1.Remove
    Set oLString = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    Set oAxisVect = Nothing
    Set oTransMatrix = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

