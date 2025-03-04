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
'   Creation Date:  Tuesday, Nov 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IIV,
'   PDMS Category Name=IEEF). Physical aspect consist of four nozzles and one revolved linestring.
'   Insulation aspect consist of single entity which is a cylinder.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   21.Jul.2003      ACM                             TR 44555 Put a check for nozzle length for CP3 and CP4 ports,such that it should'nt fail when Nozzle Length is less than FlangeThickness.
'                                                                   Code is modified to place Nozzles from the axis(center).
'                                                                   Removed GetCatalogDBConnection function from CPhysical.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private m_oCodeListMetadata     As IJDCodeListMetaData

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    Dim iOutput     As Double
    Dim ObjBody As Object

    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentDiameter1 As Double
    Dim parInstrumentDiameter2 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double
    
    Dim parNpd                     As Double
    Dim parEndPreparation          As Long
    Dim parScheduleThickness       As Long
    Dim parEndStandard             As Long
    Dim parPressureRating          As Long
    Dim parFlowDirection           As DistribFlow
    
    Dim parPortIndex1              As Long
    Dim parNpd1                    As Double
    Dim parEndPreparation1         As Long
    Dim parScheduleThickness1      As Long
    Dim parEndStandard1            As Long
    Dim parPressureRating1         As Long
    Dim parFlowDirection1          As DistribFlow
    Dim parPortIndex2              As Long
    Dim parNpd2                    As Double
    Dim parEndPreparation2         As Long
    Dim parScheduleThickness2      As Long
    Dim parEndStandard2            As Long
    Dim parPressureRating2         As Long
    Dim parFlowDirection2          As DistribFlow
    Dim parPortIndex3              As Long
    Dim parNpd3                    As Double
    Dim parEndPreparation3         As Long
    Dim parScheduleThickness3      As Long
    Dim parEndStandard3            As Long
    Dim parPressureRating3         As Long
    Dim parFlowDirection3          As DistribFlow
    Dim parPortIndex4              As Long
    Dim parNpd4                    As Double
    Dim parEndPreparation4         As Long
    Dim parScheduleThickness4      As Long
    Dim parEndStandard4            As Long
    Dim parPressureRating4         As Long
    Dim parFlowDirection4          As DistribFlow
    Dim parId1                     As String
    Dim parId2                     As String
    Dim parId3                     As String
    Dim parId4                     As String
    Dim parNPDUnitType             As String
    Dim parNpdUnitType1            As String
    Dim parNpdUnitType2            As String
    Dim parNpdUnitType3            As String
    Dim parNpdUnitType4            As String
    '
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 29) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)                        'D2 is equal to D3 + D4 + D5, so NOT used.
    parInstrumentLength = arrayOfInputs(3)               'D3
    parInstrumentLength1 = arrayOfInputs(4)             'D4
    parInstrumentLength2 = arrayOfInputs(5)             'D5
    parInstrumentDiameter = arrayOfInputs(6)           'D6
    parInstrumentDiameter1 = arrayOfInputs(7)         'D7
    parInstrumentDiameter2 = arrayOfInputs(8)         'D8
    parPortRotation1 = arrayOfInputs(9)                                 'D9
    parPortRotation2 = arrayOfInputs(10)                             'D10
    parNpd = arrayOfInputs(11)
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
    parPortIndex3 = arrayOfInputs(31)
    parNpd3 = arrayOfInputs(32)
    parEndPreparation3 = arrayOfInputs(33)
    parScheduleThickness3 = arrayOfInputs(34)
    parEndStandard3 = arrayOfInputs(35)
    parPressureRating3 = arrayOfInputs(36)
    parFlowDirection3 = arrayOfInputs(37)
    parPortIndex4 = arrayOfInputs(38)
    parNpd4 = arrayOfInputs(39)
    parEndPreparation4 = arrayOfInputs(40)
    parScheduleThickness4 = arrayOfInputs(41)
    parEndStandard4 = arrayOfInputs(42)
    parPressureRating4 = arrayOfInputs(43)
    parFlowDirection4 = arrayOfInputs(44)
    
    parInsulationThickness = arrayOfInputs(45)
    
    'Text Inputs
    parId1 = arrayOfInputs(46)
    parId2 = arrayOfInputs(47)
    parId3 = arrayOfInputs(48)
    parId4 = arrayOfInputs(49)
    parNPDUnitType = arrayOfInputs(50)
    parNpdUnitType1 = arrayOfInputs(51)
    parNpdUnitType2 = arrayOfInputs(52)
    parNpdUnitType3 = arrayOfInputs(53)
    parNpdUnitType4 = arrayOfInputs(54)

    iOutput = 0
    
'   Origin of the symbol is assumed at the intersection of axis joining port1 to port2,
'   and axis joining port3 to port4.

' Insert your code for output 2(Body of Venturi)
'   Point 1
    LineStringPoints(0) = -parInstrumentLength1 / 2 - parInstrumentLength
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = -parInstrumentLength1 / 2 - parInstrumentLength
    LineStringPoints(4) = 0
    LineStringPoints(5) = parInstrumentDiameter / 2
    
'   Assumption: Ratio of Coverging cone length to straight portion length on Port1 side is
'   0.7:0.3, diverging cone to straight portion length on Port1 side Port2 is the same.

'   Point 3
    LineStringPoints(6) = -parInstrumentLength1 / 2 - 0.3 * parInstrumentLength
    LineStringPoints(7) = 0
    LineStringPoints(8) = parInstrumentDiameter1 / 2
    
'   Point 4
    LineStringPoints(9) = -parInstrumentLength1 / 2
    LineStringPoints(10) = 0
    LineStringPoints(11) = parInstrumentDiameter1 / 2
    
'   Point 5
    LineStringPoints(12) = -parInstrumentLength1 / 2
    LineStringPoints(13) = 0
    LineStringPoints(14) = parInstrumentDiameter2 / 2
    
'   Point 6
    LineStringPoints(15) = parInstrumentLength1 / 2
    LineStringPoints(16) = 0
    LineStringPoints(17) = parInstrumentDiameter2 / 2
    
'   Point 7
    LineStringPoints(18) = parInstrumentLength1 / 2
    LineStringPoints(19) = 0
    LineStringPoints(20) = parInstrumentDiameter1 / 2
    
'   Point 8
    LineStringPoints(21) = parInstrumentLength1 / 2 + 0.3 * parInstrumentLength2
    LineStringPoints(22) = 0
    LineStringPoints(23) = parInstrumentDiameter1 / 2
    
'   Point 9
    LineStringPoints(24) = parInstrumentLength1 / 2 + parInstrumentLength2
    LineStringPoints(25) = 0
    LineStringPoints(26) = parInstrumentDiameter / 2
    
'   Point 10
    LineStringPoints(27) = parInstrumentLength1 / 2 + parInstrumentLength2
    LineStringPoints(28) = 0
    LineStringPoints(29) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 10, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set 0, 0, 0
    Set ObjBody = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
      
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
' Place Nozzle 1
'   Nozzle1 and Nozzle2  are considered at the base of cones
   Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition

    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNpdUnitType1, parEndPreparation1, parScheduleThickness1, _
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
    oPlacePoint.Set -parInstrumentLength1 / 2 - parInstrumentLength - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 2
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex2, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNpdUnitType2, parEndPreparation2, parScheduleThickness2, _
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
    oPlacePoint.Set parInstrumentLength1 / 2 + parInstrumentLength2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
' Place Nozzle 3
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNpdUnitType3, parEndPreparation3, parScheduleThickness3, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)

    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
''Nozzle is placed from center.i.e Nozzle length is D8/2
''This check helps to avoid failure of nozzle, if length is less than the flangeThick(parPortIndex3)

    If parInstrumentDiameter2 / 2 > flangeThick(parPortIndex3) Then
        oNozzle.Length = parInstrumentDiameter2 / 2
    Else
        oNozzle.Length = flangeThick(parPortIndex3)
    End If

' Direction of the Nozzle
    oDir.Set 0, Cos(parPortRotation1), Sin(parPortRotation1)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, _
                    (parInstrumentDiameter2 / 2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Cos(parPortRotation1), _
                    (parInstrumentDiameter2 / 2 + sptOffset(parPortIndex3) - depth(parPortIndex3)) * Sin(parPortRotation1)
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

' Place Nozzle 4
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNpd, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd4, _
                                            parNpdUnitType4, parEndPreparation4, parScheduleThickness4, _
                                            parEndStandard4, parPressureRating4, parFlowDirection4, _
                                            parId4, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
                                            
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    
''Nozzle is placed from center.i.e Nozzle length is D8/2
''This check helps to avoid failure of nozzle, if length is less than the flangeThick(parPortIndex3)

    If parInstrumentDiameter2 / 2 > flangeThick(parPortIndex4) Then
        oNozzle.Length = parInstrumentDiameter2 / 2
    Else
        oNozzle.Length = flangeThick(parPortIndex4)
    End If

' Direction of the Nozzle
    oDir.Set 0, Cos(parPortRotation2), Sin(parPortRotation2)
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, _
                    (parInstrumentDiameter2 / 2 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Cos(parPortRotation2), _
                    (parInstrumentDiameter2 / 2 + sptOffset(parPortIndex4) - depth(parPortIndex4)) * Sin(parPortRotation2)
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
    Set m_oCodeListMetadata = Nothing
 
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
 
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
