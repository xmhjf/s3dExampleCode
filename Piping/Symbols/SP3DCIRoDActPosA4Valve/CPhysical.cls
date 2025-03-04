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
'   Author:          SSP
'   Creation Date:  Thursday, Jan 2 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IDR4A).
'   This Symbol is same as PDS Eden Name=IDL4C
'   Physical aspect consist 1 Cylinder, 2 Cones ,1 stem ,1 box,1 dome cone,2 nozzles and 1 insulation aspect outputs.
'   The nozzles are fully Parametric.
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    
    Dim parFace1toCenter As Double             'D2
    Dim parFacetoFace As Double                'D3
    Dim parCylHeight As Double                 'D4
    Dim parActuatorLength As Double            'D5
    Dim parActuatorDiameter As Double          'D6
    Dim parActuatorHeight As Double            'D7
    Dim parActuatorHeight1 As Double           'D10
    Dim parActuatorHeight2 As Double           'D11
    Dim parActuatorWidth As Double             'D12
    Dim parActuatorWidth1 As Double            'D14
    Dim parActuatorLength1 As Double           'D25
    Dim parActuatorLength2 As Double           'D26
    Dim parActuatorWidth2 As Double            'D28
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


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorLength = arrayOfInputs(5)
    parActuatorDiameter = arrayOfInputs(6)
    parActuatorHeight = arrayOfInputs(7)
    parActuatorHeight1 = arrayOfInputs(8)
    parActuatorHeight2 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorLength1 = arrayOfInputs(12)
    parActuatorLength2 = arrayOfInputs(13)
    parActuatorWidth2 = arrayOfInputs(14)
    parInsulationThickness = arrayOfInputs(15)
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
    'Text Inputs
    parId1 = arrayOfInputs(36)
    parId2 = arrayOfInputs(37)
    parNPDUnitType = arrayOfInputs(38)
    parNPDUnitType1 = arrayOfInputs(39)
    parNPDUnitType2 = arrayOfInputs(40)
    
    iOutput = 0
    
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
    Dim StemDia As Double
    pPos1.Set 0, 0, 0
    pPos2.Set 0, parActuatorHeight - parActuatorHeight2, 0
' Assumption
' StemDia = 0.05 is taken
    StemDia = 0.05
    Set ObjStem = PlaceCylinder(m_OutputColl, pPos1, pPos2, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 7(ActuatorBox)
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Dim axisVector As New AutoMath.DVector
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
' Point 1 (Bottom Left)
    LineStrPts(0) = -parActuatorLength2
    LineStrPts(1) = parActuatorHeight - parActuatorHeight2
    LineStrPts(2) = parActuatorWidth1 + parActuatorWidth2
' Point 2 (Bottom Right)
    LineStrPts(3) = parActuatorLength1
    LineStrPts(4) = parActuatorHeight - parActuatorHeight2
    LineStrPts(5) = parActuatorWidth1 + parActuatorWidth2
' Point 3 (Top Right)
    LineStrPts(6) = parActuatorLength1
    LineStrPts(7) = parActuatorHeight + parActuatorHeight1
    LineStrPts(8) = parActuatorWidth1 + parActuatorWidth2
' Point 4 (Top Left)
    LineStrPts(9) = -parActuatorLength2
    LineStrPts(10) = parActuatorHeight + parActuatorHeight1
    LineStrPts(11) = parActuatorWidth1 + parActuatorWidth2
' Point 5 (Bottom Left)
    LineStrPts(12) = -parActuatorLength2
    LineStrPts(13) = parActuatorHeight - parActuatorHeight2
    LineStrPts(14) = parActuatorWidth1 + parActuatorWidth2
    
    
    Dim ProjLength As Double
    ProjLength = parActuatorWidth + parActuatorWidth1
    Set oLString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)

    axisVector.Set 0, 0, -1
    Set ObjActuatorBox = PlaceProjection(m_OutputColl, oLString, axisVector, ProjLength, True)
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
    Dim EleCollection   As Collection
    Dim Linepts(0 To 5) As Double
    Dim arcRadius       As Double
    Dim centerPoint   As New AutoMath.DPosition
    Dim startPoint   As New AutoMath.DPosition
    Dim endPoint   As New AutoMath.DPosition
    
        
    arcRadius = ((0.2 * parActuatorDiameter) ^ 2 + (parActuatorDiameter / 2) ^ 2) / _
                            (2 * 0.2 * parActuatorDiameter)
                            
    centerPoint.Set parActuatorLength - arcRadius, parActuatorHeight, _
                     parActuatorWidth2
    startPoint.Set parActuatorLength, parActuatorHeight, parActuatorWidth2
    endPoint.Set parActuatorLength - 0.2 * parActuatorDiameter, parActuatorHeight, _
                    parActuatorWidth2 + parActuatorDiameter / 2
                    

    Set oDomeArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                        centerPoint.x, centerPoint.y, centerPoint.z, _
                                                        startPoint.x, startPoint.y, startPoint.z, _
                                                        endPoint.x, endPoint.y, endPoint.z)
' Side of cone
    Linepts(0) = parActuatorLength - 0.2 * parActuatorDiameter
    Linepts(1) = parActuatorHeight
    Linepts(2) = parActuatorWidth2 + parActuatorDiameter / 2
    
    Linepts(3) = parActuatorLength1
    Linepts(4) = parActuatorHeight
    Linepts(5) = parActuatorWidth2
    Set oConeLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    Set EleCollection = New Collection
    EleCollection.Add oDomeArc
    EleCollection.Add oConeLine
    
    Set oComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVector.Set 1, 0, 0
    centerPoint.Set parActuatorLength1, parActuatorHeight, parActuatorWidth2
' Revolve it about X-Axis
    Set ObjActuatorCone = PlaceRevolution(m_OutputColl, oComplexStr, axisVector, centerPoint, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCone
    Set ObjActuatorCone = Nothing
    
    
    Dim Objcurves As IJDObject
    Set Objcurves = oComplexStr
    Objcurves.Remove
    Set Objcurves = Nothing
    Set Objcurves = oLString
    Objcurves.Remove
    Set Objcurves = Nothing
    
    Dim iCount As Integer
    For iCount = 1 To EleCollection.Count
        EleCollection.Remove 1
    Next iCount
    Set EleCollection = Nothing
    
    Set oLString = Nothing
    Set oDomeArc = Nothing
    Set oConeLine = Nothing
    Set oComplexStr = Nothing
    Set axisVector = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set centerPoint = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set pPos1 = Nothing
    Set pPos2 = Nothing
    Set m_oCodeListMetadata = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

