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
'   The symbol is prepared based on (SPR1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   8-Nov-07      svsmylav/KKC  DI-130618:Changed the parameter names of parPositionerOrientation and
'                               parHandWheelOrientation to parPositionerRotAngle and parHandWheelStepRotAngle respectively
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI As Double


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
    Dim ObjValveCone1 As Object
    Dim ObjValveCone2 As Object
    Dim ObjStemRod As Object
    Dim ObjValveCylinder As Object
    Dim ObjActuatorRect As Object
    Dim ObjActuatorCyl As Object

    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorWidth2 As Double
    Dim parHandWhOutertoCenter As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parHandWheelOffset1 As Double
    Dim parPositionerLength As Double
    Dim parPositionerHeight As Double
    Dim parPositionerWidth As Double
    Dim parPositionerOffset As Double
    Dim parPositionerOffset1 As Double
    Dim parRotation As Double
    Dim parInsulationThickness As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parOperationHeight As Double
    Dim parPositionerRotAngle As Double
    Dim parHandWheelStepRotAngle As Double
    
    '   Dynamic Port Input parameters variable declaration
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
    parFace1toCenter = arrayOfInputs(2) 'V1
    parFacetoFace = arrayOfInputs(3)   'V2
    parCylHeight = arrayOfInputs(4) 'V3
    parActuatorLength = arrayOfInputs(5) 'A11
    parActuatorLength1 = arrayOfInputs(6) 'A10
    parActuatorLength2 = arrayOfInputs(7) 'A9
    parActuatorDiameter = arrayOfInputs(8) 'A3
    parActuatorHeight = arrayOfInputs(9)   'A2
    parActuatorHeight1 = arrayOfInputs(10) 'A7
    parActuatorHeight2 = arrayOfInputs(11) 'A8
    parActuatorWidth = arrayOfInputs(12) 'A4
    parActuatorWidth1 = arrayOfInputs(13) 'A5
    parActuatorWidth2 = arrayOfInputs(14) 'A6
    parHandWhOutertoCenter = arrayOfInputs(15) 'H2
    parHandWheelDiameter = arrayOfInputs(16) 'H3
    parHandWheelOffset = arrayOfInputs(17) 'H4
    parHandWheelOffset1 = arrayOfInputs(18) 'H5
    parPositionerLength = arrayOfInputs(19) 'P6
    parPositionerHeight = arrayOfInputs(20) 'P4
    parPositionerWidth = arrayOfInputs(21) 'P2
    parPositionerOffset = arrayOfInputs(22) 'P5
    parPositionerOffset1 = arrayOfInputs(23) 'P3
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
    parOperationDiameter = arrayOfInputs(46) 'I2
    parOperationDiameter1 = arrayOfInputs(47) 'I3
    parOperationHeight = arrayOfInputs(48)    'I1
    parPositionerRotAngle = arrayOfInputs(49)
    parHandWheelStepRotAngle = arrayOfInputs(50)
'   Text Inputs
    parId1 = arrayOfInputs(51)
    parId2 = arrayOfInputs(52)
    parNPDUnitType = arrayOfInputs(53)
    parNPDUnitType1 = arrayOfInputs(54)
    parNPDUnitType2 = arrayOfInputs(55)
    
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parActuatorWidth2, 0) Then parActuatorWidth2 = parActuatorWidth1 / 2
    If CmpDblEqual(parActuatorLength1, 0) Then parActuatorLength1 = parActuatorLength2 / 2
    If CmpDblEqual(parActuatorHeight2, 0) Then parActuatorHeight2 = parActuatorHeight1 / 2
    
    iOutput = 0
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
        
' Place Nozzle 1 (Nozzle1)
    
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
    
 ' Insert your code for output 3(Conical Valve Body 1)
         
    Dim oStPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
         
    oStPoint.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    oEnPoint.Set 0, 0, 0
    
    Set ObjValveCone1 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone1
    Set ObjValveCone1 = Nothing
    
 ' Insert your code for output 4(Conical Valve Body 2 )
         
    oStPoint.Set parFacetoFace - parFace1toCenter - flangeThick(parPortIndex2), 0, 0
    oEnPoint.Set 0, 0, 0
        
    Set ObjValveCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex2) / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone2
    Set ObjValveCone2 = Nothing
    
 ' Insert your code for output 5(Stem Body )
   Dim dStemDia As Double
   
   dStemDia = 0.05
   oStPoint.Set 0, 0, 0
   oEnPoint.Set 0, (parActuatorWidth - parActuatorWidth2), 0
 
   Set ObjStemRod = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemRod
    Set ObjStemRod = Nothing
    
 ' Insert your code for output 6(Valve Cylindrical Body)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -parCylHeight, 0
    Set ObjValveCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCylinder
    Set ObjValveCylinder = Nothing
    
 ' Insert your code for output 7(Actuator Rectangular box )

    Dim dLineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim oAxisVector As AutoMath.DVector
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
    
' Point 1 (Bottom Left)
    dLineStrPts(0) = -parActuatorLength1
    dLineStrPts(1) = (parActuatorWidth - parActuatorWidth2)
    dLineStrPts(2) = -parActuatorHeight2
' Point 2 (Bottom Right)
    dLineStrPts(3) = parActuatorLength2 - parActuatorLength1
    dLineStrPts(4) = (parActuatorWidth - parActuatorWidth2)
    dLineStrPts(5) = -parActuatorHeight2
' Point 3 (Top Right)
    dLineStrPts(6) = parActuatorLength2 - parActuatorLength1
    dLineStrPts(7) = (parActuatorWidth - parActuatorWidth2)
    dLineStrPts(8) = parActuatorHeight1 - parActuatorHeight2
' Point 4 (Top Left)
    dLineStrPts(9) = -parActuatorLength1
    dLineStrPts(10) = (parActuatorWidth - parActuatorWidth2)
    dLineStrPts(11) = parActuatorHeight1 - parActuatorHeight2
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dProjLength As Double
    dProjLength = parActuatorWidth1
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    oAxisVector.Set 0, 1, 0
    Set ObjActuatorRect = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
' Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
' Axis of Rotation is Stem Axis
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjActuatorRect.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorRect
    Set ObjActuatorRect = Nothing
    
 ' Insert your code for output 8(Actuator Cylindrical Body)
   
   oStPoint.Set -parActuatorLength, parActuatorWidth, parActuatorHeight1 - parActuatorHeight2
   oEnPoint.Set -parActuatorLength, parActuatorWidth, parActuatorHeight
   
   Set ObjActuatorCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parActuatorDiameter, True)
   ObjActuatorCyl.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCyl
    Set ObjActuatorCyl = Nothing
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
    If parHandWheelStepRotAngle = 1 Then  '0 Deg
        oTorOrigin.Set -parHandWhOutertoCenter, _
                               (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
    
        oTorAxis.Set 1, 0, 0
    ElseIf parHandWheelStepRotAngle = 2 Then     ' 90 deg
        oTorOrigin.Set parHandWheelOffset1, _
                           parHandWhOutertoCenter, -parHandWheelOffset
         
         oTorAxis.Set 0, 1, 0
         
    ElseIf parHandWheelStepRotAngle = 3 Then      ' 180 deg
         oTorOrigin.Set parHandWhOutertoCenter, _
                               (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
    
        oTorAxis.Set 1, 0, 0
         
    ElseIf parHandWheelStepRotAngle = 4 Then       ' 270 deg
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
    
' Insert your code for output 10(Hand Wheel Stem)
   
    If parHandWheelStepRotAngle = 1 Then      ' 0 deg
        oStPoint.Set -parHandWhOutertoCenter + dHandWheelThk, _
                              (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
        oEnPoint.Set -parActuatorLength1, _
                              (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
    ElseIf parHandWheelStepRotAngle = 2 Then   ' 90 deg
        oStPoint.Set parHandWheelOffset1, parHandWhOutertoCenter - dHandWheelThk, _
                               -parHandWheelOffset
        oEnPoint.Set parHandWheelOffset1, _
                              (parActuatorWidth + parActuatorWidth1 - parActuatorWidth2), -parHandWheelOffset
         
    ElseIf parHandWheelStepRotAngle = 3 Then     ' 180 deg
         oStPoint.Set parHandWhOutertoCenter - dHandWheelThk, _
                              (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
         oEnPoint.Set parActuatorLength2 - parActuatorLength1, _
                              (parActuatorWidth + parHandWheelOffset1), -parHandWheelOffset
         
    ElseIf parHandWheelStepRotAngle = 4 Then  ' 270 deg
         oStPoint.Set parHandWheelOffset1, parHandWhOutertoCenter - dHandWheelThk, _
                               -parHandWheelOffset
         oEnPoint.Set parHandWheelOffset1, _
                             parActuatorWidth - parActuatorWidth2, -parHandWheelOffset
         
    End If
    Dim ObjHandWheelStem As Object
    Set ObjHandWheelStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
    ObjHandWheelStem.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandWheelStem
    Set ObjHandWheelStem = Nothing

' Insert your code for output 11(Positioner Body)
    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
   
' Point 1 (Bottom Left)

    dLineStrPts(0) = -parPositionerLength - parActuatorLength
    dLineStrPts(1) = (parActuatorWidth - parPositionerOffset1 + parPositionerWidth)
    dLineStrPts(2) = parPositionerOffset
    
' Point 2 (Bottom Right)
    dLineStrPts(3) = -parPositionerLength - parActuatorLength
    dLineStrPts(4) = (parActuatorWidth - parPositionerOffset1)
    dLineStrPts(5) = parPositionerOffset
' Point 3 (Top Right)
    dLineStrPts(6) = -parPositionerLength - parActuatorLength
    dLineStrPts(7) = (parActuatorWidth - parPositionerOffset1)
    dLineStrPts(8) = parPositionerOffset - parPositionerHeight
' Point 4 (Top Left)
    dLineStrPts(9) = -parPositionerLength - parActuatorLength
    dLineStrPts(10) = (parActuatorWidth - parPositionerOffset1 + parPositionerWidth)
    dLineStrPts(11) = parPositionerOffset - parPositionerHeight
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    
    dProjLength = parPositionerLength - parActuatorDiameter / 2
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    oAxisVector.Set 1, 0, 0
    Dim ObjPositioner As Object
    Set ObjPositioner = PlaceProjection(m_OutputColl, oLString, oAxisVector, dProjLength, True)
    Set oLString = Nothing
    
    ' Rotation
    Dim oTransMat1 As DT4x4
    Dim oTransVector As AutoMath.DVector
    Set oTransMat1 = New AutoMath.DT4x4
    Set oTransVector = New AutoMath.DVector
    oTransVector.Set parActuatorLength, -parActuatorWidth, -(parPositionerOffset - parPositionerHeight / 2)
    oTransMat1.Translate oTransVector
    ObjPositioner.Transform oTransMat1
    
' Axis of Rotation is Stem Axis
     oAboutVec.Set 0, 0, 1
     Dim oTransMat2 As DT4x4
     Set oTransMat2 = New AutoMath.DT4x4
     oTransMat2.Rotate -parPositionerRotAngle, oAboutVec
     ObjPositioner.Transform oTransMat2
     oTransVector.Set -parActuatorLength, parActuatorWidth, (parPositionerOffset - parPositionerHeight / 2)
     oTransMat1.LoadIdentity
     oTransMat1.Translate oTransVector
     ObjPositioner.Transform oTransMat1
     ObjPositioner.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPositioner
    Set ObjPositioner = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oAxisVector = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

