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
'   Author:          BG
'   Creation Date:  Thusday, Dec 26 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG5.
'   This symbol is Custom INtools Instrument with Nozzles are fully Parametric.
'   Physical aspect consist Two cones, one converging and other  diverging, 2-cylinders along X-axis,
'   1-Cylinder along Y-axis and two nozzles. Insulation aspect consist of same number of outputs except for the
'   Nozzles.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.May.2008     KKC  CR-135970  Provide ability to rotate actuator for on-the-fly control valves
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
    Dim ObjVertCylin As Object
    Dim ObjHoriCylin As Object
    
    Dim parFace1toCenter As Double
    Dim parFaceToFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
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
    parInstrumentHeight = arrayOfInputs(4) 'D4
    parInstrumentDiameter = arrayOfInputs(5)      'D5
    parInstrumentDiameter1 = arrayOfInputs(6)  'D6
    parInstrumentLength = arrayOfInputs(7)     'D7
    parInstrumentLength1 = arrayOfInputs(8)    'D8
    parInstrumentLength2 = arrayOfInputs(9)       'D9
    parInsulationThickness = arrayOfInputs(10)
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
    
    If UBound(arrayOfInputs) > 35 Then
       parAngle = arrayOfInputs(36)
    End If
    iOutput = 0
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' Insert your code for output 1(Vertical Cylinder)
    Dim CylDia As Double
    'Stem Diameter is assumed as 0.05  as given in EDEN of this Symbol
    CylDia = 0.05
    stPoint.Set 0, 0, 0
    enPoint.Set 0, (parInstrumentHeight - Sqr((parInstrumentDiameter / 2) ^ 2 - (CylDia / 2) ^ 2)), 0
    Set ObjVertCylin = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertCylin
    Set ObjVertCylin = Nothing
    
' Insert your code for output 2(Horizontal Cylinders along X-axis)
    Dim centerPt As AutoMath.DPosition
    Dim revVector As AutoMath.DVector
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLinePts(0 To 17) As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set centerPt = New AutoMath.DPosition
    Set revVector = New AutoMath.DVector

    'Point1
    dblLinePts(0) = -parInstrumentLength1
    dblLinePts(1) = parInstrumentHeight
    dblLinePts(2) = 0
    'Point2
    dblLinePts(3) = -parInstrumentLength1
    dblLinePts(4) = (parInstrumentHeight - parInstrumentDiameter / 2)
    dblLinePts(5) = 0
    'Point3
    dblLinePts(6) = parInstrumentLength2
    dblLinePts(7) = (parInstrumentHeight - parInstrumentDiameter / 2)
    dblLinePts(8) = 0
    'Point4
    dblLinePts(9) = parInstrumentLength2
    dblLinePts(10) = (parInstrumentHeight - parInstrumentDiameter1 / 2)
    dblLinePts(11) = 0
    'Point5
    dblLinePts(12) = parInstrumentLength
    dblLinePts(13) = (parInstrumentHeight - parInstrumentDiameter1 / 2)
    dblLinePts(14) = 0
    'Point6
    dblLinePts(15) = parInstrumentLength
    dblLinePts(16) = parInstrumentHeight
    dblLinePts(17) = 0
    
    Set objLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dblLinePts)
    revVector.Set 1, 0, 0
    centerPt.Set 0, parInstrumentHeight, 0
    Set ObjHoriCylin = PlaceRevolution(m_OutputColl, objLineStr, revVector, centerPt, 2 * PI, True)
    
    Dim oAxisVect As AutoMath.DVector
    Dim oTransMatrix As IJDT4x4
    Set oAxisVect = New AutoMath.DVector
    Set oTransMatrix = New AutoMath.DT4x4
    oAxisVect.Set 0, 1, 0
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parAngle, oAxisVect
    ObjHoriCylin.Transform oTransMatrix

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylin
    Set ObjHoriCylin = Nothing
    Dim tmpLineStr As IJDObject
    Set tmpLineStr = objLineStr
    Set objLineStr = Nothing
    tmpLineStr.Remove
    Set geomFactory = Nothing
    Set revVector = Nothing
    Set centerPt = Nothing
    
    ' Place Nozzle 1 for output 3 (Nozzle1)
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
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
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
    oPlacePoint.Set parFaceToFace - parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' Insert your code for output 2(Body Cone2 of Valve)

    stPoint.Set (parFaceToFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
