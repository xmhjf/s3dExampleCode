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
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 8 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA4A,
'   PDMS Category Name=IVCL).The symbol consists of eleven outputs, Nine physical outputs,
'   one Insulation Output and one Operation output.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.May.2008     KKC  CR-135970  Provide ability to rotate actuator for on-the-fly control valves
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
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjStem As Object
    Dim ObjLargeBox As Object
    Dim ObjSmallBox As Object
    Dim ObjCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parInsulationThickness As Double
    Dim parangle As Double
    
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

    Dim StemDia As Double
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim parOperationHeight As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)       'D4
    parActuatorHeight = arrayOfInputs(5)        'D5
    parActuatorHeight1 = arrayOfInputs(6)       'D6
    parActuatorLength = arrayOfInputs(7)         'D7
    parActuatorLength1 = arrayOfInputs(8)        'D8
    parActuatorWidth = arrayOfInputs(9)        'D9
    parActuatorWidth1 = arrayOfInputs(10)       'D10
    parActuatorHeight2 = arrayOfInputs(11)       'D11
    parActuatorWidth2 = arrayOfInputs(12)       'D12
    parActuatorWidth3 = arrayOfInputs(13)      'D13
    parActuatorLength2 = arrayOfInputs(14)       'D14
    parActuatorLength3 = arrayOfInputs(15)       'D15
    parActuatorOffset = arrayOfInputs(16)       'D16
    parActuatorCylDiameter = arrayOfInputs(17)  'D17
    parActuatorCylLength = arrayOfInputs(18)    'D18
    parActuatorHeight3 = arrayOfInputs(19)      'D19
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
    parOperationHeight = arrayOfInputs(41)     'D20
    
    parId1 = arrayOfInputs(42)
    parId2 = arrayOfInputs(43)
    parNPDUnitType = arrayOfInputs(44)
    parNPDUnitType1 = arrayOfInputs(45)
    parNPDUnitType2 = arrayOfInputs(46)
    
    If UBound(arrayOfInputs) > 46 Then
        parangle = arrayOfInputs(47)
    End If
   
    iOutput = 0
    
'   Stem diameter is taken as 0.05
    StemDia = 0.05

' Origin is assumed to be at the center of the valveBody along the flow axis

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
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + _
                        sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Insert your code for output 2(Converging Cone Body)
    stPoint.Set -(parFace1toCenter - flangeThick(parPortIndex1)), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
' Insert your code for output 3(Body Cone2 of Valve)
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 4(Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 '  Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing

' Insert your code for output 5(Stem)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight3 - parActuatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 6(LargeBox)
    Dim dLineStringPoints(0 To 14) As Double
    dLineStringPoints(0) = -(parActuatorLength1 - parActuatorOffset)
    dLineStringPoints(1) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(2) = -parActuatorWidth1

    dLineStringPoints(3) = -(parActuatorLength1 - parActuatorOffset)
    dLineStringPoints(4) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(5) = parActuatorWidth

    dLineStringPoints(6) = parActuatorLength + parActuatorOffset
    dLineStringPoints(7) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(8) = parActuatorWidth

    dLineStringPoints(9) = parActuatorLength + parActuatorOffset
    dLineStringPoints(10) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(11) = -parActuatorWidth1

    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)
    Dim oLineString1 As IngrGeom3D.LineString3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Set ObjLargeBox = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight + parActuatorHeight1, True)

    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parangle, oAxisVect
    ObjLargeBox.Transform oTransMatrix
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLargeBox
    Set ObjLargeBox = Nothing
    
' Insert your code for output 7(SmallBox)
    dLineStringPoints(0) = parActuatorLength3
    dLineStringPoints(1) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(2) = parActuatorWidth3

    dLineStringPoints(3) = parActuatorLength3
    dLineStringPoints(4) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(5) = -parActuatorWidth2

    dLineStringPoints(6) = -parActuatorLength2
    dLineStringPoints(7) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(8) = -parActuatorWidth2

    dLineStringPoints(9) = -parActuatorLength2
    dLineStringPoints(10) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(11) = parActuatorWidth3

    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)

    Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Set ObjSmallBox = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight2, True)

    ObjSmallBox.Transform oTransMatrix
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSmallBox
    Set ObjSmallBox = Nothing
    
' Insert your code for output 8(Cylinder)
    stPoint.Set parActuatorOffset, parActuatorHeight3, -parActuatorWidth1
    enPoint.Set parActuatorOffset, _
                        parActuatorHeight3, -parActuatorWidth1 - parActuatorCylLength
    Set ObjCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCylDiameter, True)
    ObjCylinder.Transform oTransMatrix
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oTransMatrix = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
