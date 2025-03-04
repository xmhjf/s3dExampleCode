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
'   Creation Date:  Tuesday 10 Jul 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on (SDR1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----     ------------------
'   08.Nov.07     RRK/KKC    DI-130618: Changed the parameter names of parPositionerOrientation and
'                            parHandWheelOrientation to parPositionerRotAngle and parHandWheelStepRotAngle respectively
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI As Double
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
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjValCylinder As Object
    Dim ObjStem As Object
    Dim ObjActuatorBox As Object
    Dim ObjActuatorCone As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double             'V1
    Dim parFacetoFace As Double                'V2
    Dim parCylHeight As Double                 'V3
    Dim parActuatorWidth As Double             'A2
    Dim parActuatorDiameter As Double          'A3
    Dim parActuatorHeight As Double            'A4
    Dim parActuatorHeight1 As Double           'A5
    Dim parActuatorHeight2 As Double           'A6
    Dim parActuatorLength As Double            'A9
    Dim parActuatorLength1 As Double           'A10
    Dim parActuatorWidth1 As Double            'A7
    Dim parActuatorWidth2 As Double            'A8
    Dim parActuatorLength2 As Double           'A11
    Dim parHandWheelOffset As Double           'H4
    Dim parHandWhOutertoCenter As Double           'H2
    Dim parHandWheelDiameter As Double         'H3
    Dim parHandWheelOffset1 As Double           'H5
    Dim parPositionerHeight As Double          'P2
    Dim parPositionerWidth As Double           'P4
    Dim parPositionerOffset As Double          'P3
    Dim parPositionerOffset1 As Double         'P6
    Dim parPositionerOffset2 As Double         'P5
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parOperationHeight As Double
    Dim parPositionerRotAngle As Double
    Dim parHandWheelStepRotAngle As Double
    
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
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorWidth = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight = arrayOfInputs(7)
    parActuatorHeight1 = arrayOfInputs(8)
    parActuatorHeight2 = arrayOfInputs(9)
    parActuatorLength = arrayOfInputs(10)
    parActuatorLength1 = arrayOfInputs(11)
    parActuatorWidth1 = arrayOfInputs(12)
    parActuatorWidth2 = arrayOfInputs(13)
    parActuatorLength2 = arrayOfInputs(14)
    parHandWheelOffset = arrayOfInputs(15)
    parHandWheelOffset1 = arrayOfInputs(16)
    parHandWhOutertoCenter = arrayOfInputs(17)
    parHandWheelDiameter = arrayOfInputs(18)
    parPositionerHeight = arrayOfInputs(19)
    parPositionerWidth = arrayOfInputs(20)
    parPositionerOffset = arrayOfInputs(21)
    parPositionerOffset1 = arrayOfInputs(22)
    parPositionerOffset2 = arrayOfInputs(23)
    parInsulationThickness = arrayOfInputs(24)
    parRotation = arrayOfInputs(25)
    parNPD = arrayOfInputs(26)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(27)
    parScheduleThickness = arrayOfInputs(28)
    parEndStandard = arrayOfInputs(29)
    parPressureRating = arrayOfInputs(30)
    parFlowDirection = arrayOfInputs(31)
    parPortIndex1 = arrayOfInputs(32)
    parNpd1 = arrayOfInputs(33)
    parEndPreparation1 = arrayOfInputs(34)
    parScheduleThickness1 = arrayOfInputs(35)
    parEndStandard1 = arrayOfInputs(36)
    parPressureRating1 = arrayOfInputs(37)
    parFlowDirection1 = arrayOfInputs(38)
    parPortIndex2 = arrayOfInputs(39)
    parNpd2 = arrayOfInputs(40)
    parEndPreparation2 = arrayOfInputs(41)
    parScheduleThickness2 = arrayOfInputs(42)
    parEndStandard2 = arrayOfInputs(43)
    parPressureRating2 = arrayOfInputs(44)
    parFlowDirection2 = arrayOfInputs(45)
    parOperationDiameter = arrayOfInputs(46)
    parOperationDiameter1 = arrayOfInputs(47)
    parOperationHeight = arrayOfInputs(48)
    parPositionerRotAngle = arrayOfInputs(49)
    parHandWheelStepRotAngle = arrayOfInputs(50)
'   Text Inputs
    parId1 = arrayOfInputs(51)
    parId2 = arrayOfInputs(52)
    parNPDUnitType = arrayOfInputs(53)
    parNPDUnitType1 = arrayOfInputs(54)
    parNPDUnitType2 = arrayOfInputs(55)
    
    iOutput = 0
    
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parActuatorHeight2, 0) Then parActuatorHeight2 = parActuatorHeight1 / 2
    If CmpDblEqual(parActuatorWidth2, 0) Then parActuatorWidth2 = parActuatorWidth1 / 2
    If CmpDblEqual(parActuatorLength1, 0) Then parActuatorLength1 = parActuatorLength / 2
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
'  Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

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
    
' Insert your code for output 3(ConvergingCone)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition
    
    pPos1.Set -parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0
    
    Set ObjConvergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
' Insert your code for output 4(DivergingCone)
    pPos1.Set parFacetoFace - parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0

    Set ObjDivergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing

' Insert your code for output 5(ValCylinder)
    pPos1.Set 0, 0, 0
    pPos2.Set 0, -parCylHeight, 0
    
    Dim ValCylDiam As Double
' Assumption
' ValCylDiam = parFace1toCenter is taken
    ValCylDiam = parFace1toCenter
    
    Set ObjValCylinder = PlaceCylinder(m_OutputColl, pPos1, pPos2, ValCylDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCylinder
    Set ObjValCylinder = Nothing
    
' Insert your code for output 6(Stem)
    Dim dStemDia As Double
    pPos1.Set 0, 0, 0
    pPos2.Set 0, (parActuatorHeight - parActuatorHeight2), 0
' Assumption
' dStemDia = 0.05 is taken
    dStemDia = 0.02
    Set ObjStem = PlaceCylinder(m_OutputColl, pPos1, pPos2, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 7(ActuatorBox)
    Dim dLineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Dim oAxisVector As AutoMath.DVector
    Set oAxisVector = New AutoMath.DVector
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
' Point 1 (Bottom Left)
    dLineStrPts(0) = -parActuatorLength1
    dLineStrPts(1) = (parActuatorHeight - parActuatorHeight2)
    dLineStrPts(2) = -parActuatorWidth2
' Point 2 (Bottom Right)
    dLineStrPts(3) = -parActuatorLength1
    dLineStrPts(4) = (parActuatorHeight - parActuatorHeight2)
    dLineStrPts(5) = parActuatorWidth1 - parActuatorWidth2
' Point 3 (Top Right)
    dLineStrPts(6) = -parActuatorLength1
    dLineStrPts(7) = (parActuatorHeight - parActuatorHeight2 + parActuatorHeight1)
    dLineStrPts(8) = parActuatorWidth1 - parActuatorWidth2
' Point 4 (Top Left)
    dLineStrPts(9) = -parActuatorLength1
    dLineStrPts(10) = (parActuatorHeight - parActuatorHeight2 + parActuatorHeight1)
    dLineStrPts(11) = -parActuatorWidth2
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    
    Dim dProjLength As Double
    dProjLength = parActuatorLength
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    oAxisVector.Set 1, 0, 0
    Set ObjActuatorBox = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
' Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
' Axis of Rotation is Stem Axis
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjActuatorBox.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBox
    Set ObjActuatorBox = Nothing
    
' Insert your code for output 8(ActuatorCone)
' height of circular Dome is assumed as 0.2 times of the diameter - parActuatorDiameter
' Place Elliptical Arc on Right Side
    Dim oConeLine       As IngrGeom3D.Line3d
    Dim oDomeArc        As IngrGeom3D.Arc3d
    Dim oComplexStr     As IngrGeom3D.ComplexString3d
    Dim oEleCollection   As Collection
    Dim dLinepts(0 To 5) As Double
    Dim dArcRadius       As Double
    Dim oCenterPoint   As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    Dim oStartPoint   As AutoMath.DPosition
    Set oStartPoint = New AutoMath.DPosition
    Dim oEndPoint   As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
        
    dArcRadius = ((0.2 * parActuatorDiameter) ^ 2 + (parActuatorDiameter / 2) ^ 2) / _
                            (2 * 0.2 * parActuatorDiameter)
                            
    oCenterPoint.Set -parActuatorLength2, parActuatorHeight, _
                        parActuatorWidth - dArcRadius
    oStartPoint.Set -parActuatorLength2, parActuatorHeight, parActuatorWidth
    oEndPoint.Set -parActuatorLength2 - parActuatorDiameter / 2, parActuatorHeight, _
                           parActuatorWidth - 0.2 * parActuatorDiameter
                    
    Set oDomeArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                        oStartPoint.x, oStartPoint.y, oStartPoint.z, _
                                                        oEndPoint.x, oEndPoint.y, oEndPoint.z)
' Side of cone
    dLinepts(0) = -parActuatorLength2 - parActuatorDiameter / 2
    dLinepts(1) = parActuatorHeight
    dLinepts(2) = parActuatorWidth - 0.2 * parActuatorDiameter
    
    dLinepts(3) = -parActuatorLength2
    dLinepts(4) = parActuatorHeight
    dLinepts(5) = parActuatorWidth1 - parActuatorWidth2
    Set oConeLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dLinepts(0), dLinepts(1), dLinepts(2), _
                                                                        dLinepts(3), dLinepts(4), dLinepts(5))
    Set oEleCollection = New Collection
    oEleCollection.Add oDomeArc
    oEleCollection.Add oConeLine
    
    Set oComplexStr = PlaceTrCString(oStartPoint, oEleCollection)
    oAxisVector.Set 0, 0, 1
    oCenterPoint.Set -parActuatorLength2, parActuatorHeight, parActuatorWidth1 - parActuatorWidth2
' Revolve it about Z-Axis
    Set ObjActuatorCone = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVector, oCenterPoint, PI * 2, True)
'Rotation
    ObjActuatorCone.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCone
    Set ObjActuatorCone = Nothing
    
' Insert your code for output 7(Positioner Body)
'Create a line string and project it

'Point 1
    dLineStrPts(0) = -parActuatorLength2 - parPositionerOffset1
    dLineStrPts(1) = (parActuatorHeight + parPositionerHeight - parPositionerOffset)
    dLineStrPts(2) = parPositionerOffset2 - parPositionerWidth
'Point 2
    dLineStrPts(3) = -parActuatorLength2 - parPositionerOffset1
    dLineStrPts(4) = (parActuatorHeight - parPositionerOffset)
    dLineStrPts(5) = parPositionerOffset2 - parPositionerWidth
'Point 3
    dLineStrPts(6) = -parActuatorLength2 - parPositionerOffset1
    dLineStrPts(7) = (parActuatorHeight - parPositionerOffset)
    dLineStrPts(8) = parPositionerOffset2
'Point 4
    dLineStrPts(9) = -parActuatorLength2 - parPositionerOffset1
    dLineStrPts(10) = (parActuatorHeight + parPositionerHeight - parPositionerOffset)
    dLineStrPts(11) = parPositionerOffset2
'Point 5
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)

    dProjLength = parPositionerOffset1 - (((parActuatorDiameter / 2) * (parPositionerOffset2 - parPositionerWidth - (parActuatorWidth1 - parActuatorWidth2))) _
                            / (parActuatorWidth - 0.2 * parActuatorDiameter - (parActuatorWidth1 - parActuatorWidth2)))

    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)
     
    oAxisVector.Set 1, 0, 0
    Dim ObjPositionerBody As Object
    Set ObjPositionerBody = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    
'Transform the output  (Rotation about stem axis)
    Set oTransMat = New AutoMath.DT4x4
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ' Rotation about Actuator Axis
    Dim oTransMat1 As DT4x4
    Dim oTransVector As AutoMath.DVector
    Set oTransMat1 = New AutoMath.DT4x4
    Set oTransVector = New AutoMath.DVector
    oTransVector.Set parActuatorLength2, -parActuatorHeight, -(parPositionerOffset2 - parPositionerWidth / 2)
    oTransMat1.Translate oTransVector
    ObjPositionerBody.Transform oTransMat1
    Set oTransVector = Nothing
    ' Axis of Rotation is Stem Axis
     oAboutVec.Set 0, 0, 1
     Dim oTransMat2 As DT4x4
     Set oTransMat2 = New AutoMath.DT4x4
     Set oTransVector = New AutoMath.DVector
     oTransMat2.Rotate -parPositionerRotAngle, oAboutVec
     ObjPositionerBody.Transform oTransMat2
     oTransVector.Set -parActuatorLength2, parActuatorHeight, (parPositionerOffset2 - parPositionerWidth / 2)
     oTransMat1.LoadIdentity
     oTransMat1.Translate oTransVector
     ObjPositionerBody.Transform oTransMat1
     ObjPositionerBody.Transform oTransMat
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositionerBody
    Set ObjPositionerBody = Nothing
    
' Insert your code for output 8(HandWheel)
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
    If parHandWheelStepRotAngle = 1 Then  '0 Deg
       oTorOrigin.Set -parHandWhOutertoCenter, _
                          (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
       oTorAxis.Set 1, 0, 0
    ElseIf parHandWheelStepRotAngle = 2 Then   ' 90 Deg
        oTorOrigin.Set parHandWheelOffset1, _
                           parHandWhOutertoCenter, -parHandWheelOffset
         
         oTorAxis.Set 0, 1, 0
         
    ElseIf parHandWheelStepRotAngle = 3 Then    '180 Deg
         oTorOrigin.Set parHandWhOutertoCenter, _
                               (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
    
        oTorAxis.Set 1, 0, 0
         
    ElseIf parHandWheelStepRotAngle = 4 Then   '270 Deg
        oTorOrigin.Set parHandWheelOffset1, _
                           parHandWhOutertoCenter, -parHandWheelOffset
         
         oTorAxis.Set 0, 1, 0
         
     End If
    Set ObjHandwheel = PlaceTorus(m_OutputColl, oTorOrigin, oTorAxis, dTorPriRad, dTorSecRad)
    
    'Transform the output  (Rotation about stem axis)
    ObjHandwheel.Transform oTransMat

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set oTorOrigin = Nothing
    Set oTorAxis = Nothing
    
 ' Insert your code for output 9(HandwheelStem)

    If parHandWheelStepRotAngle = 1 Then  '0 Deg
        oStartPoint.Set -parHandWhOutertoCenter + dHandWheelThk / 2, _
                              (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
        oEndPoint.Set -parActuatorLength1, _
                            (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
    ElseIf parHandWheelStepRotAngle = 2 Then  ' 90 deg
        oStartPoint.Set parHandWheelOffset1, parHandWhOutertoCenter - dHandWheelThk, _
                               -parHandWheelOffset
        oEndPoint.Set parHandWheelOffset1, _
                              (parActuatorHeight + parActuatorHeight1 - parActuatorHeight2), -parHandWheelOffset
         
    ElseIf parHandWheelStepRotAngle = 3 Then    ' 180 deg
         oStartPoint.Set parHandWhOutertoCenter - dHandWheelThk, _
                              (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
         oEndPoint.Set parActuatorLength2 - parActuatorLength1, _
                              (parActuatorHeight + parHandWheelOffset1), -parHandWheelOffset
         
    ElseIf parHandWheelStepRotAngle = 4 Then   '270 deg
         oStartPoint.Set parHandWheelOffset1, parHandWhOutertoCenter - dHandWheelThk, _
                               -parHandWheelOffset
         oEndPoint.Set parHandWheelOffset1, _
                             parActuatorHeight - parActuatorHeight2, -parHandWheelOffset
         
    End If
    Dim ObjHandwheelStem As Object
    Set ObjHandwheelStem = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dStemDia, True)
    
    'Transform the output  (Rotation about stem axis)
    ObjHandwheelStem.Transform oTransMat
  
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheelStem
    Set ObjHandwheelStem = Nothing
    Set oEleCollection = Nothing
    
    Set oDomeArc = Nothing
    Set oConeLine = Nothing
    Set oComplexStr = Nothing
    Set oAxisVector = Nothing
    Set oTransMat = Nothing
    Set oAboutVec = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oCenterPoint = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    Set pPos1 = Nothing
    Set pPos2 = Nothing
    Set m_oCodeListMetadata = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
