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
'   The symbol is prepared based on INtools Instrument Components(SREG1)
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
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

    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjTopValCyl  As Object
    Dim ObjHoriCylinders As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    Dim parOperationLength As Double
    Dim parOperationHeight As Double
    Dim iOutput     As Integer
    
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
    parFace1toCenter = arrayOfInputs(2) 'V1
    parFacetoFace = arrayOfInputs(3)        'V2
    parCylHeight = arrayOfInputs(4)          'V3
    parInstrumentLength = arrayOfInputs(5)   'A3
    parInstrumentHeight = arrayOfInputs(6)   'A2
    parInstrumentDiameter = arrayOfInputs(7)     'A4
    parInstrumentLength1 = arrayOfInputs(8)    'A5
    parInstrumentLength2 = arrayOfInputs(9)    'A7
    parInstrumentDiameter1 = arrayOfInputs(10)        'A6
    parInsulationThickness = arrayOfInputs(11)
    parRotation = arrayOfInputs(12)
    parNPD = arrayOfInputs(13)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parPortIndex1 = arrayOfInputs(19)
    parNpd1 = arrayOfInputs(20)
    parEndPreparation1 = arrayOfInputs(21)
    parScheduleThickness1 = arrayOfInputs(22)
    parEndStandard1 = arrayOfInputs(23)
    parPressureRating1 = arrayOfInputs(24)
    parFlowDirection1 = arrayOfInputs(25)
    parPortIndex2 = arrayOfInputs(26)
    parNpd2 = arrayOfInputs(27)
    parEndPreparation2 = arrayOfInputs(28)
    parScheduleThickness2 = arrayOfInputs(29)
    parEndStandard2 = arrayOfInputs(30)
    parPressureRating2 = arrayOfInputs(31)
    parFlowDirection2 = arrayOfInputs(32)
    parOperationLength = arrayOfInputs(33) 'I2
    parOperationHeight = arrayOfInputs(34) 'I1
'Text Inputs
    parId1 = arrayOfInputs(35)
    parId2 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)
    
    iOutput = 0
    If CmpDblEqual(parFace1toCenter, 0) Then parFace1toCenter = parFacetoFace / 2
    If CmpDblEqual(parCylHeight, 0) Then parCylHeight = 0.00001
    If CmpDblEqual(parInstrumentLength2, 0) Then parInstrumentLength2 = parInstrumentLength1 / 2
' Place Nozzle 1 for output 5 (Nozzle1)
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

' Insert your code for output 1(Body Cone1 of Valve)
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
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
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
 ' Insert your code for output 2(Body Cone2 of Valve)
    oStPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    oEnPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim dCylDia As Double
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, -parCylHeight, 0
    dCylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing

' Insert your code for output 3(Top Vertical Cylinder)
    'Minimum Diameter of 0.05 is assumed for Top Vertical Cylinder as used in EDEN of INTOOLS Symbol.
    Dim dMiniCylDia As Double
    Dim dMeetRadius  As Double
    dMiniCylDia = 0.05
    dMeetRadius = Sqr((parInstrumentDiameter / 2) ^ 2 - (dMiniCylDia / 2) ^ 2)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, (parInstrumentHeight - dMeetRadius), 0
    Set ObjTopValCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dMiniCylDia, True)
' Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
' Axis of Rotation is Stem Axis
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjTopValCyl.Transform oTransMat
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopValCyl
    Set ObjTopValCyl = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
' Insert your code for output 4 (Horizontal Cylinders)
    Dim oCenterPt As AutoMath.DPosition
    Dim oRevVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 17) As Double
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oCenterPt = New AutoMath.DPosition
    Set oRevVector = New AutoMath.DVector

    'Point1
    dblLinePts(0) = parInstrumentLength2
    dblLinePts(1) = parInstrumentHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = parInstrumentLength2
    dblLinePts(4) = parInstrumentHeight - parInstrumentDiameter / 2
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = -(parInstrumentLength1 - parInstrumentLength2)
    dblLinePts(7) = parInstrumentHeight - parInstrumentDiameter / 2
    dblLinePts(8) = 0
    'Point4
    dblLinePts(9) = -(parInstrumentLength1 - parInstrumentLength2)
    dblLinePts(10) = parInstrumentHeight - parInstrumentDiameter1 / 2
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = -parInstrumentLength
    dblLinePts(13) = parInstrumentHeight - parInstrumentDiameter1 / 2
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = -parInstrumentLength
    dblLinePts(16) = parInstrumentHeight
    dblLinePts(17) = 0
    
    Set objLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dblLinePts)
    oRevVector.Set 1, 0, 0
    oCenterPt.Set 0, parInstrumentHeight, 0
    Set ObjHoriCylinders = PlaceRevolution(m_OutputColl, objLineStr, oRevVector, oCenterPt, 2 * PI, True)
    ObjHoriCylinders.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylinders
    Set ObjHoriCylinders = Nothing
    Set objLineStr = Nothing
    Set oRevVector = Nothing
    Set oCenterPt = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
 
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing

End Sub
