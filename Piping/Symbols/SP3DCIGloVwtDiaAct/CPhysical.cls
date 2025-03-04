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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Monday, Jun 23 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  10.Jun.2008     KKC      CR-137825  Piping instrument catalog should be updated to current Intools Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
      ''
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 3) As Double
    Dim sptOffset(1 To 3) As Double
    Dim flangeDiam(1 To 3)    As Double
    Dim depth(1 To 3) As Double
    Dim flangeThick(1 To 3) As Double
    
    Dim iOutput     As Integer
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjActuatorBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorDiameter As Double
    Dim parOperationHeight As Double
    Dim parHandWheelLength As Double
    Dim parHandWheelDiameter As Double
    Dim parHandWheelOffset As Double
    Dim parPositionerHeight As Double
    Dim parPositionerOffset As Double
    Dim parPostionerLength As Double
    Dim parPositionerOffset1 As Double
    Dim parPositionerOffset2 As Double
    Dim parOperationDiameter As Double
    Dim parOperationDiameter1 As Double
    Dim parRotation As Double
    Dim parInsulationThickness As Double
    
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

    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String

    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim dLineStringPoints(0 To 14) As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFacetoFace = arrayOfInputs(3)    'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parActuatorHeight = arrayOfInputs(5)    'D5
    parActuatorDiameter = arrayOfInputs(6)  'D6
    parOperationHeight = arrayOfInputs(7)   'D7
    parHandWheelLength = arrayOfInputs(8)  'D8
    parHandWheelDiameter = arrayOfInputs(9) 'D9
    parPositionerHeight = arrayOfInputs(10) 'D10
    parPositionerOffset = arrayOfInputs(11) ' D11
    parPositionerOffset1 = arrayOfInputs(12)  'D12
    parActuatorHeight1 = arrayOfInputs(13) 'D13
    parPositionerOffset2 = arrayOfInputs(14)  'D14
    parHandWheelOffset = arrayOfInputs(15) 'D16
    parOperationDiameter = arrayOfInputs(16)  'D17
    parOperationDiameter1 = arrayOfInputs(17)  ' D18
    parPostionerLength = arrayOfInputs(18)  ' D20
    parRotation = arrayOfInputs(19)  ' D25
    parInsulationThickness = arrayOfInputs(20)
    parNPD = arrayOfInputs(21)
    parEndPreparation = arrayOfInputs(22)
    parScheduleThickness = arrayOfInputs(23)
    parEndStandard = arrayOfInputs(24)
    parPressureRating = arrayOfInputs(25)
    parFlowDirection = arrayOfInputs(26)
    parPortIndex1 = arrayOfInputs(27)
    parNpd1 = arrayOfInputs(28)
    parEndPreparation1 = arrayOfInputs(29)
    parScheduleThickness1 = arrayOfInputs(30)
    parEndStandard1 = arrayOfInputs(31)
    parPressureRating1 = arrayOfInputs(32)
    parFlowDirection1 = arrayOfInputs(33)
    parPortIndex2 = arrayOfInputs(34)
    parNpd2 = arrayOfInputs(35)
    parEndPreparation2 = arrayOfInputs(36)
    parScheduleThickness2 = arrayOfInputs(37)
    parEndStandard2 = arrayOfInputs(38)
    parPressureRating2 = arrayOfInputs(39)
    parFlowDirection2 = arrayOfInputs(40)
    parId1 = arrayOfInputs(41)
    parId2 = arrayOfInputs(42)
    parNPDUnitType = arrayOfInputs(43)
    parNPDUnitType1 = arrayOfInputs(44)
    parNPDUnitType2 = arrayOfInputs(45)
    
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
     
' Insert your code for output 4(Body Cone1)
    oStPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex1) / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 5(Body Cone2 of Valve)
    oStPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex2) / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 6(Bottom Valve Cylinder)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -parCylHeight, 0
    Set ObjValCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parActuatorDiameter, True)
 
 '  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
' Insert your code for output 7(Actuator Body and the stem)
    oStPoint.Set 0, parActuatorHeight - parActuatorHeight1, 0
    oEnPoint.Set 0, parActuatorHeight, 0
    Set ObjActuatorBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parActuatorDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorBody
    Set ObjActuatorBody = Nothing
    
' Insert your code for the stem connecting the actuator and the valve body
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, parActuatorHeight - parActuatorHeight1, 0
    Dim objStem As Object
    Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parActuatorDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStem
    Set objStem = Nothing
    
' Place your code to form the Positioner body

    'Point 1
    dLineStringPoints(0) = parPositionerOffset
    dLineStringPoints(1) = parPositionerOffset2 - parPositionerHeight
    dLineStringPoints(2) = -parPositionerOffset1
'Point 2
    dLineStringPoints(3) = parPositionerOffset
    dLineStringPoints(4) = parPositionerOffset2
    dLineStringPoints(5) = -parPositionerOffset1
'Point 3
    dLineStringPoints(6) = -(parPostionerLength - parPositionerOffset)
    dLineStringPoints(7) = parPositionerOffset2
    dLineStringPoints(8) = -parPositionerOffset1
'Point 4
    dLineStringPoints(9) = -(parPostionerLength - parPositionerOffset)
    dLineStringPoints(10) = parPositionerOffset2 - parPositionerHeight
    dLineStringPoints(11) = -parPositionerOffset1
'Point 5
    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)
    
    Dim dProjLength As Double
    dProjLength = parPositionerOffset1 - 0.2 * parActuatorDiameter
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)
     
    oAxisVect.Set 0, 0, 1
    Dim ObjPositionerBody As Object
    Set ObjPositionerBody = PlaceProjection(m_OutputColl, oLineString, oAxisVect, dProjLength, True)
    oAxisVect.Set 0, 1, 0
    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parRotation, oAxisVect
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
    Dim dHandWheelThk As Double
    dHandWheelThk = 0.05
    TorSecRad = dHandWheelThk / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    TorOrigin.Set -(parHandWheelLength - dHandWheelThk / 2), parHandWheelOffset, 0
    TorAxis.Set -1, 0, 0
    Set ObjHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, TorPriRad, TorSecRad)
    ObjHandwheel.Transform oTransMatrix
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheel
    Set ObjHandwheel = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
    
 ' Insert your code for output 9(HandwheelStem)
    oStPoint.Set 0, parHandWheelOffset, 0
    oEnPoint.Set -(parHandWheelLength - dHandWheelThk / 2), parHandWheelOffset, 0
    Dim dStemDia As Double
    Dim ObjHandwheelStem As Object
    dStemDia = 0.05
    Set ObjHandwheelStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
    ObjHandwheelStem.Transform oTransMatrix
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHandwheelStem
    Set ObjHandwheelStem = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxisVect = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    Set oTransMatrix = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
