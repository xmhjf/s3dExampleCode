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
'   Creation Date:  Tuesday, Dec 19 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG1,
'   PDMS Category Name=IVRF). This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. Physical aspect consist of Two cones one converging other diverging
'   and a vertical Cylinder with varying diameter, two nozzles one facing -ive direction direction of
'   X-axis and the other in +ive direction of X-axis.
'   Insulation aspect consist of one Cylinder along X-axis and another cylinder along Y-axis
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
Private PI As Double
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
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjConvergCone As Object
    Dim ObjDivergCone As Object
    Dim ObjVertCylins As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parCylHeight As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
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
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFaceToFace = arrayOfInputs(3)        'D3
    parCylHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight = arrayOfInputs(5)      'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentDiameter = arrayOfInputs(8)     'D8
    parInstrumentDiameter1 = arrayOfInputs(9)        'D9
    parInsulationThickness = arrayOfInputs(10)          'Insulation
    parNPD = arrayOfInputs(11)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(12)
    parScheduleThickness = arrayOfInputs(13)
    parEndStandard = arrayOfInputs(14)
    parPressureRating = arrayOfInputs(15)
    parFlowDirection = arrayOfInputs(16)
    parPortIndex1 = arrayOfInputs(17)
    parNpd1 = arrayOfInputs(18)
    parEndPreparation1 = arrayOfInputs(19)
    parScheduleThickness1 = arrayOfInputs(20)
    parEndStandard1 = arrayOfInputs(21)
    parPressureRating1 = arrayOfInputs(22)
    parFlowDirection1 = arrayOfInputs(23)
    parPortIndex2 = arrayOfInputs(24)
    parNpd2 = arrayOfInputs(25)
    parEndPreparation2 = arrayOfInputs(26)
    parScheduleThickness2 = arrayOfInputs(27)
    parEndStandard2 = arrayOfInputs(28)
    parPressureRating2 = arrayOfInputs(29)
    parFlowDirection2 = arrayOfInputs(30)
    parId1 = arrayOfInputs(31)
    parId2 = arrayOfInputs(32)
    parNPDUnitType = arrayOfInputs(33)
    parNPDUnitType1 = arrayOfInputs(34)
    parNPDUnitType2 = arrayOfInputs(35)
    
    iOutput = 0
        
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
        
' Insert your code for output 1(Converging Cone)
    Dim centerTopPt As AutoMath.DPosition
    Dim centerBasePt As AutoMath.DPosition
    Set centerTopPt = New AutoMath.DPosition
    Set centerBasePt = New AutoMath.DPosition
'Origin of the symbol is assumed at the intersection of vertical cylinder axis and nozzles  axes
    centerBasePt.Set -parFace1toCenter + flangeThick(parPortIndex1), 0, 0
    centerTopPt.Set 0, 0, 0
    Set ObjConvergCone = PlaceCone(m_OutputColl, centerBasePt, centerTopPt, pipeDiam(parPortIndex1) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergCone
    Set ObjConvergCone = Nothing
    
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
    oPlacePoint.Set parFaceToFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

 ' Insert your code for output 2(Diverging Cone)
    centerBasePt.Set parFaceToFace - parFace1toCenter - flangeThick(parPortIndex2), 0, 0
    centerTopPt.Set 0, 0, 0
    Set ObjDivergCone = PlaceCone(m_OutputColl, centerBasePt, centerTopPt, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergCone
    Set ObjDivergCone = Nothing
    Set centerTopPt = Nothing
    Set centerBasePt = Nothing
    
' Insert your code for output 3(Vertical Cylinders)
    Dim CylDia As Double
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 29) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector
    'Bottom Cylinder Diameter is taken as equal to Pipe 1 Diameter as per EDEN
    CylDia = pipeDiam(parPortIndex1)
    'Point1
    dblLinePts(0) = 0
    dblLinePts(1) = -parCylHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = CylDia / 2
    dblLinePts(4) = -parCylHeight
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = CylDia / 2
    dblLinePts(7) = 0
    dblLinePts(8) = 0
    'Point4
    'Cylinder, immediately above the cone, diameter is assumed as 0.5*times the parInstrumentDiameter value
    dblLinePts(9) = parInstrumentDiameter / 4
    dblLinePts(10) = 0
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = parInstrumentDiameter / 4
    dblLinePts(13) = parInstrumentHeight - parInstrumentHeight2 - parInstrumentHeight1    '(D5-D7-D6)
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = parInstrumentDiameter1 / 2
    dblLinePts(16) = parInstrumentHeight - parInstrumentHeight2 - parInstrumentHeight1   '(D5-D7-D6)
    dblLinePts(17) = 0
    'Point7
    dblLinePts(18) = parInstrumentDiameter1 / 2
    dblLinePts(19) = parInstrumentHeight - parInstrumentHeight1                                 '(D5-D6)
    dblLinePts(20) = 0
    'Point8
    dblLinePts(21) = parInstrumentDiameter / 2
    dblLinePts(22) = parInstrumentHeight - parInstrumentHeight1                                '(D5-D6)
    dblLinePts(23) = 0
    'Point9
    dblLinePts(24) = parInstrumentDiameter / 2
    dblLinePts(25) = parInstrumentHeight
    dblLinePts(26) = 0
    'Point10
    dblLinePts(27) = 0
    dblLinePts(28) = parInstrumentHeight
    dblLinePts(29) = 0
    
    Set objLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dblLinePts)
    revVector.Set 0, 1, 0
    centerPt.Set 0, 0, 0
    Set ObjVertCylins = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylins
    Set ObjVertCylins = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    tmpLineStr.Remove
    Set objLineStr = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

