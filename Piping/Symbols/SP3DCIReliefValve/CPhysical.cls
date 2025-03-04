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
'   The symbol is prepared based on  (SRV1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
'''
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
    
    Dim iOutput As Integer
    Dim ObjValveBody As Object
    Dim ObjTopCylBody1 As Object
    Dim ObjTopCylBody2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorWidth As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationHeight As Double
    Dim parOperationLength As Double
    Dim parOperationWidth As Double
    Dim parOperationWidth1 As Double
    
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
    parFace1toCenter = arrayOfInputs(2)     'R1
    parFace2toCenter = arrayOfInputs(3)     'R2
    parInstrumentHeight = arrayOfInputs(4)      'R3
    parInstrumentLength = arrayOfInputs(5)      'R5
    parInstrumentWidth = arrayOfInputs(6)       'R6
    parInstrumentWidth1 = arrayOfInputs(7)      'R7
    parInstrumentHeight1 = arrayOfInputs(8)    'R4
    parActuatorHeight = arrayOfInputs(9)    'A2
    parActuatorHeight1 = arrayOfInputs(10)   'A3
    parActuatorLength = arrayOfInputs(11)    'A4
    parActuatorWidth = arrayOfInputs(12)     'A5
    parInstrumentDiameter = arrayOfInputs(13)  'R8
    parInsulationThickness = arrayOfInputs(14)
    parRotation = arrayOfInputs(15) 'A1
    parNPD = arrayOfInputs(16)                       'Input parameters for Dynamic Nozzles
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
    parOperationHeight = arrayOfInputs(36) 'I1
    parOperationLength = arrayOfInputs(37) 'I2
    parOperationWidth = arrayOfInputs(38)  'I3
    parOperationWidth1 = arrayOfInputs(39)  'I4
' Text inputs
    parId1 = arrayOfInputs(40)
    parId2 = arrayOfInputs(41)
    parNPDUnitType = arrayOfInputs(42)
    parNPDUnitType1 = arrayOfInputs(43)
    parNPDUnitType2 = arrayOfInputs(44)
    
    iOutput = 0
    If CmpDblEqual(parInstrumentWidth1, 0) Then
        parInstrumentWidth1 = parInstrumentWidth / 2
    End If
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
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
    oNozzle.Length = parFace1toCenter - parInstrumentLength

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
    oNozzle.Length = 0.1 * parFace2toCenter + flangeThick(2)

' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset(parPortIndex2) + depth(parPortIndex2), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
'The origin is considered to be at intersection of nozzle axis
 ' Insert your code for output 3(Valve Body)
    oStPoint.Set parInstrumentLength, -(0.9 * parFace2toCenter - flangeThick(2)), -(parInstrumentWidth - parInstrumentWidth1)
    oEnPoint.Set -parInstrumentLength, parInstrumentHeight, parInstrumentWidth1
    Set ObjValveBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing
    
 ' Insert your code for output 2(Top Cylinder Body1)
    oStPoint.Set 0, parInstrumentHeight, 0
    oEnPoint.Set 0, parInstrumentHeight + parInstrumentHeight1, 0
    Set ObjTopCylBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylBody1
    Set ObjTopCylBody1 = Nothing
    
    
' Insert your code for Output 4 (Positioner Body)
    Dim dLineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim oAxisVector As AutoMath.DVector
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    Set oLString = New IngrGeom3D.LineString3d
    Set oAxisVector = New AutoMath.DVector
    
' Point 1 (Bottom Left)
    dLineStrPts(0) = parActuatorLength
    dLineStrPts(1) = parActuatorHeight - parFace2toCenter
    dLineStrPts(2) = -parActuatorWidth / 2
' Point 2 (Bottom Right)
    dLineStrPts(3) = parInstrumentDiameter / 2
    dLineStrPts(4) = parActuatorHeight - parFace2toCenter
    dLineStrPts(5) = -parActuatorWidth / 2
' Point 3 (Top Right)
    dLineStrPts(6) = parInstrumentDiameter / 2
    dLineStrPts(7) = parActuatorHeight - parFace2toCenter
    dLineStrPts(8) = parActuatorWidth / 2
    
' Point 4 (Top Left)
    dLineStrPts(9) = parActuatorLength
    dLineStrPts(10) = parActuatorHeight - parFace2toCenter
    dLineStrPts(11) = parActuatorWidth / 2
' Point 5 (Bottom Left)
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(1)
    dLineStrPts(14) = dLineStrPts(2)
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dProjLength As Double
    dProjLength = parActuatorHeight1
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)
   
    oAxisVector.Set 0, 1, 0
    Dim ObjActuatorRect As Object
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

    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
    
End Sub

