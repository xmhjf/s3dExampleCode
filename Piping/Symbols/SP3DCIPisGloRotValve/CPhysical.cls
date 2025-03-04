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
'   Author       : KKC
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on (SP1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------   -----       ------------------
'   8-Nov-07        svsmylav/KKC     DI-130618: Changed the parameter names of parPositionerOrientation and parHandWheelOrientation
'                                          to parPositionerRotAngle and parHandWheelRotAngle respectively
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
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
    Dim ObjStemBody As Object
    Dim ObjLargeBoxBody As Object
    Dim ObjSmallBoxBody As Object
    Dim ObjCylinder1Body As Object
    Dim ObjCylinder2Body As Object
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2  As Object
    Dim ObjValCyl  As Object
        
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parActuatorCyl1Diameter As Double
    Dim parActuatorCyl1Length As Double
    Dim parHandWheelLength As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parPositionerLength As Double
    Dim parPositionerHeight As Double
    Dim parPostionerOffset As Double
    Dim parPostionerOffset1 As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parPositionerRotAngle As Double
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

    Dim dStemdia As Double
    Dim parOperationHeight As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'V1
    parFacetoFace = arrayOfInputs(3)          'V2
    parCylHeight = arrayOfInputs(4)           'V3
    parActuatorHeight = arrayOfInputs(5)      'A6
    parActuatorHeight1 = arrayOfInputs(6)     'A5
    parActuatorWidth = arrayOfInputs(7)       'A7
    parActuatorWidth1 = arrayOfInputs(8)      'A11
    parActuatorLength = arrayOfInputs(9)      'A3
    parActuatorLength1 = arrayOfInputs(10)    'A4
    parActuatorHeight2 = arrayOfInputs(11)    'A9
    parActuatorLength2 = arrayOfInputs(12)    'A10
    parActuatorWidth2 = arrayOfInputs(13)     'A8
    parActuatorOffset = arrayOfInputs(14)     'C1
    parActuatorCylDiameter = arrayOfInputs(15) 'C11
    parActuatorCylLength = arrayOfInputs(16)   'C12
    parActuatorHeight3 = arrayOfInputs(17)     'A2
    parActuatorCyl1Diameter = arrayOfInputs(18)    'C21
    parActuatorCyl1Length = arrayOfInputs(19)      'C22
    parHandWheelLength = arrayOfInputs(20)         'H2
    parHandWheelDiameter = arrayOfInputs(21)         'H3
    parHandWheelOffset = arrayOfInputs(22)          'H4
    parPositionerHeight = arrayOfInputs(23)        'P4
    parPositionerLength = arrayOfInputs(24)         'P2
    parPostionerOffset = arrayOfInputs(25)         'P3
    parPostionerOffset1 = arrayOfInputs(26)         'P5
    parInsulationThickness = arrayOfInputs(27)
    parRotation = arrayOfInputs(28)
    parNPD = arrayOfInputs(29)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(30)
    parScheduleThickness = arrayOfInputs(31)
    parEndStandard = arrayOfInputs(32)
    parPressureRating = arrayOfInputs(33)
    parFlowDirection = arrayOfInputs(34)
    parPortIndex1 = arrayOfInputs(35)
    parNpd1 = arrayOfInputs(36)
    parEndPreparation1 = arrayOfInputs(37)
    parScheduleThickness1 = arrayOfInputs(38)
    parEndStandard1 = arrayOfInputs(39)
    parPressureRating1 = arrayOfInputs(40)
    parFlowDirection1 = arrayOfInputs(41)
    parPortIndex2 = arrayOfInputs(42)
    parNpd2 = arrayOfInputs(43)
    parEndPreparation2 = arrayOfInputs(44)
    parScheduleThickness2 = arrayOfInputs(45)
    parEndStandard2 = arrayOfInputs(46)
    parPressureRating2 = arrayOfInputs(47)
    parFlowDirection2 = arrayOfInputs(48)
    parOperationHeight = arrayOfInputs(49)     'I1
    parPositionerRotAngle = arrayOfInputs(50)       'P1
    parHandWheelRotAngle = arrayOfInputs(51)        'H1
'   Text Inputs
    parId1 = arrayOfInputs(52)
    parId2 = arrayOfInputs(53)
    parNPDUnitType = arrayOfInputs(54)
    parNPDUnitType1 = arrayOfInputs(55)
    parNPDUnitType2 = arrayOfInputs(56)

    iOutput = 0
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parActuatorLength1, 0) Then parActuatorLength1 = parActuatorLength / 2
    If CmpDblEqual(parActuatorHeight, 0) Then parActuatorHeight = parActuatorHeight1 / 2
    If CmpDblEqual(parActuatorWidth2, 0) Then parActuatorWidth2 = parActuatorWidth / 2
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

' Origin is assumed to be at the center of the valveBody along the flow axis
  ' Insert your code for output 1(Body Cone1 of Valve)

    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    dStemdia = 0.05
    
    oStPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
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
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

 ' Insert your code for output 2(Body Cone 2 of Valve)
    oStPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim CylDia As Double
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
        
' Insert your code for output 4(Stem Body)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, parActuatorHeight3 - parActuatorHeight1, 0
    Set ObjStemBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemdia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemBody
    Set ObjStemBody = Nothing
    
 ' Insert your code for output 5(Large Box Body)
 
    Dim dLineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim oAxisVector As AutoMath.DVector
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
    
' Point 1 (Bottom Left)
    dLineStrPts(0) = -(parActuatorLength - parActuatorLength1)
    dLineStrPts(1) = parActuatorHeight3 - parActuatorHeight1
    dLineStrPts(2) = parActuatorWidth - parActuatorWidth2
' Point 2 (Bottom Right)
    dLineStrPts(3) = parActuatorLength1
    dLineStrPts(4) = parActuatorHeight3 - parActuatorHeight1
    dLineStrPts(5) = parActuatorWidth - parActuatorWidth2
' Point 3 (Top Right)
    dLineStrPts(6) = parActuatorLength1
    dLineStrPts(7) = parActuatorHeight3 - parActuatorHeight1
    dLineStrPts(8) = -parActuatorWidth2
' Point 4 (Top Left)
    dLineStrPts(9) = -(parActuatorLength - parActuatorLength1)
    dLineStrPts(10) = parActuatorHeight3 - parActuatorHeight1
    dLineStrPts(11) = -parActuatorWidth2
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dProjLength As Double
    dProjLength = parActuatorHeight1
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    oAxisVector.Set 0, 1, 0
    Set ObjLargeBoxBody = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
' Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
' Axis of Rotation is Stem Axis
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjLargeBoxBody.Transform oTransMat
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLargeBoxBody
    Set ObjLargeBoxBody = Nothing
    
 ' Insert your code for output 6(Small Box Body)
    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
   
    
' Point 1 (Bottom Left)
    dLineStrPts(0) = -parActuatorLength2 / 2
    dLineStrPts(1) = parActuatorHeight3
    dLineStrPts(2) = parActuatorWidth1 / 2
' Point 2 (Bottom Right)
    dLineStrPts(3) = parActuatorLength2 / 2
    dLineStrPts(4) = parActuatorHeight3
    dLineStrPts(5) = parActuatorWidth1 / 2
' Point 3 (Top Right)
    dLineStrPts(6) = parActuatorLength2 / 2
    dLineStrPts(7) = parActuatorHeight3
    dLineStrPts(8) = -parActuatorWidth1 / 2
' Point 4 (Top Left)
    dLineStrPts(9) = -parActuatorLength2 / 2
    dLineStrPts(10) = parActuatorHeight3
    dLineStrPts(11) = -parActuatorWidth1 / 2
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    
    dProjLength = parActuatorHeight2
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    oAxisVector.Set 0, 1, 0
    Set ObjSmallBoxBody = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
' Rotation
    ObjSmallBoxBody.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSmallBoxBody
    Set ObjSmallBoxBody = Nothing
    
 ' Insert your code for output 7(Cylinder 1 Body)
    oStPoint.Set -(parActuatorCylLength + parActuatorLength - parActuatorLength1), _
                              parActuatorHeight3 - parActuatorHeight1 + parActuatorHeight, -parActuatorOffset
    oEnPoint.Set -(parActuatorLength - parActuatorLength1), _
                              parActuatorHeight3 - parActuatorHeight1 + parActuatorHeight, -parActuatorOffset
    Set ObjCylinder1Body = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parActuatorCylDiameter, True)
    ObjCylinder1Body.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1Body
    Set ObjCylinder1Body = Nothing
    
 ' Insert your code for output 8(Cylinder2 Body)
    oStPoint.Set parActuatorCyl1Length + parActuatorLength1, _
                              parActuatorHeight3 - parActuatorHeight1 + parActuatorHeight, -parActuatorOffset
    oEnPoint.Set parActuatorLength1, _
                              parActuatorHeight3 - parActuatorHeight1 + parActuatorHeight, -parActuatorOffset
    Set ObjCylinder2Body = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parActuatorCyl1Diameter, True)
    ObjCylinder2Body.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2Body
    Set ObjCylinder2Body = Nothing
' Insert your code for output 9(HandWheel)
   '   Place Torus
    Dim ObjHandwheel    As IngrGeom3D.Torus3d
    Dim dTorSecRad As Double
    Dim dTorPriRad As Double
    Dim oTorOrigin As AutoMath.DPosition
    Set oTorOrigin = New AutoMath.DPosition
    Dim oTorAxis As AutoMath.DVector
    Set oTorAxis = New AutoMath.DVector
    Dim dHandWheelThk As Double
    dHandWheelThk = 0.02
    dTorSecRad = dHandWheelThk / 2
    dTorPriRad = parHandWheelDiameter / 2 - dTorSecRad
    oTorOrigin.Set -parHandWheelLength, _
                           parHandWheelOffset, 0

    oTorAxis.Set 1, 0, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, oTorOrigin, oTorAxis, dTorPriRad, dTorSecRad)
    
    'Transform the output  (Rotation about stem axis)
    Dim oTransMat1 As IJDT4x4
    Set oTransMat1 = New AutoMath.DT4x4
    oTransMat1.Rotate parHandWheelRotAngle, oAboutVec
    ObjHandwheel.Transform oTransMat1
    ObjHandwheel.Transform oTransMat

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set oTorOrigin = Nothing
    Set oTorAxis = Nothing
    
' Insert your code for output 10(Hand Wheel Stem)
    oStPoint.Set -parHandWheelLength + dHandWheelThk, _
                              parHandWheelOffset, 0
    oEnPoint.Set -0, parHandWheelOffset, 0
    Dim ObjHandWheelStem As Object
    Set ObjHandWheelStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemdia, True)
    ObjHandWheelStem.Transform oTransMat1
    ObjHandWheelStem.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelStem
    Set ObjHandWheelStem = Nothing
        
' Insert your code for output 11(Positioner Body)

    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
   
' Point 1 (Bottom Left)
    dLineStrPts(0) = -parPostionerOffset
    dLineStrPts(1) = parPostionerOffset1 - parPositionerHeight
    dLineStrPts(2) = parPositionerLength / 2
' Point 2 (Bottom Right)
    dLineStrPts(3) = -parPostionerOffset + parPositionerLength
    dLineStrPts(4) = parPostionerOffset1 - parPositionerHeight
    dLineStrPts(5) = parPositionerLength / 2
' Point 3 (Top Right)
    dLineStrPts(6) = -parPostionerOffset + parPositionerLength
    dLineStrPts(7) = parPostionerOffset1 - parPositionerHeight
    dLineStrPts(8) = 0
' Point 4 (Top Left)
    dLineStrPts(9) = -parPostionerOffset
    dLineStrPts(10) = parPostionerOffset1 - parPositionerHeight
    dLineStrPts(11) = 0
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    
    dProjLength = parPositionerHeight
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)
    
    oAxisVector.Set 0, 1, 0
    Dim ObjPositioner As Object
    Set ObjPositioner = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
'Rotation
    Dim oTransMat2 As IJDT4x4
    Set oTransMat2 = New AutoMath.DT4x4
    oTransMat2.Rotate parPositionerRotAngle, oAboutVec
    ObjPositioner.Transform oTransMat2
    ObjPositioner.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositioner
    Set ObjPositioner = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
