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
'   Creation Date:  Friday, Dec 3 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA2,
'   PDMS Category Name=IVCJ). Physical aspect consist of two cones and a cylinder as a part
'   valve body, Stem cylinder, Two boxes , two cylinders as part of operator and Two Nozzles.
'   Insulation aspect consist of single cylinder for valve body.
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
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjStemBody As Object
    Dim ObjLargeBoxBody As Object
    Dim ObjSmallBoxBody As Object
    Dim ObjCylinder1Body As Object
    Dim ObjCylinder2Body As Object
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2  As Object
    Dim ObjValCyl  As Object
        
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parActuatorCyl1Diameter As Double
    Dim parActuatorCyl1Length As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    
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

    Dim StemDia As Double
    Dim parOperationHeight As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parActuatorHeight = arrayOfInputs(5)       'D5
    parActuatorHeight1 = arrayOfInputs(6)      'D6
    parActuatorWidth = arrayOfInputs(7)    'D7
    parActuatorWidth1 = arrayOfInputs(8)       'D8
    parActuatorLength = arrayOfInputs(9)       'D9
    parActuatorLength1 = arrayOfInputs(10)      'D10
    parActuatorHeight2 = arrayOfInputs(11)      'D11
    parActuatorLength2 = arrayOfInputs(12)      'D12
    parActuatorLength3 = arrayOfInputs(13)     'D13
    parActuatorWidth2 = arrayOfInputs(14)      'D14
    parActuatorWidth3 = arrayOfInputs(15)      'D15
    parActuatorOffset = arrayOfInputs(16)      'D16
    parActuatorCylDiameter = arrayOfInputs(17)    'D17
    parActuatorCylLength = arrayOfInputs(18)      'D18
    parActuatorHeight3 = arrayOfInputs(19)     'D19
    parActuatorCyl1Diameter = arrayOfInputs(20)    'D20
    parActuatorCyl1Length = arrayOfInputs(21)      'D21
    parInsulationThickness = arrayOfInputs(22)
    parNPD = arrayOfInputs(23)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(24)
    parScheduleThickness = arrayOfInputs(25)
    parEndStandard = arrayOfInputs(26)
    parPressureRating = arrayOfInputs(27)
    parFlowDirection = arrayOfInputs(28)
    parPortIndex1 = arrayOfInputs(29)
    parNpd1 = arrayOfInputs(30)
    parEndPreparation1 = arrayOfInputs(31)
    parScheduleThickness1 = arrayOfInputs(32)
    parEndStandard1 = arrayOfInputs(33)
    parPressureRating1 = arrayOfInputs(34)
    parFlowDirection1 = arrayOfInputs(35)
    parPortIndex2 = arrayOfInputs(36)
    parNpd2 = arrayOfInputs(37)
    parEndPreparation2 = arrayOfInputs(38)
    parScheduleThickness2 = arrayOfInputs(39)
    parEndStandard2 = arrayOfInputs(40)
    parPressureRating2 = arrayOfInputs(41)
    parFlowDirection2 = arrayOfInputs(42)
    parOperationHeight = arrayOfInputs(43)     'D22
    
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    If UBound(arrayOfInputs) > 48 Then
        parAngle = arrayOfInputs(49)
    End If
    
    iOutput = 0
    
    ' Place Nozzle 1 for output 1 (Nozzle1)
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

' Origin is assumed to be at the center of the valveBody along the flow axis
  ' Insert your code for output 1(Body Cone1 of Valve)

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    StemDia = 0.05
    
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
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

 ' Insert your code for output 2(Body Cone2 of Valve)
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick(parPortIndex2), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
        
' Insert your code for output 4(Stem Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight3 - parActuatorHeight, 0
    Set ObjStemBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemBody
    Set ObjStemBody = Nothing
    
 ' Insert your code for output 5(Large Box Body)
    Dim dLineStringPoints(0 To 14) As Double
    dLineStringPoints(0) = -parActuatorLength
    dLineStringPoints(1) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(2) = parActuatorWidth + parActuatorOffset

    dLineStringPoints(3) = -parActuatorLength
    dLineStringPoints(4) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(5) = -(parActuatorWidth1 - parActuatorOffset)

    dLineStringPoints(6) = parActuatorLength1
    dLineStringPoints(7) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(8) = -(parActuatorWidth1 - parActuatorOffset)
    
    dLineStringPoints(9) = parActuatorLength1
    dLineStringPoints(10) = parActuatorHeight3 - parActuatorHeight
    dLineStringPoints(11) = parActuatorWidth + parActuatorOffset
    
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
    Set ObjLargeBoxBody = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight + parActuatorHeight1, True)

    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parAngle, oAxisVect
    ObjLargeBoxBody.Transform oTransMatrix
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLargeBoxBody
    Set ObjLargeBoxBody = Nothing
    
 ' Insert your code for output 6(Small Box Body)
    dLineStringPoints(0) = -parActuatorLength2
    dLineStringPoints(1) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(2) = parActuatorWidth2
    
    dLineStringPoints(3) = -parActuatorLength2
    dLineStringPoints(4) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(5) = -parActuatorWidth3

    dLineStringPoints(6) = parActuatorLength3
    dLineStringPoints(7) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(8) = -parActuatorWidth3

    dLineStringPoints(9) = parActuatorLength3
    dLineStringPoints(10) = parActuatorHeight3 + parActuatorHeight1
    dLineStringPoints(11) = parActuatorWidth2

    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)
    
    Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Set ObjSmallBoxBody = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight2, True)

    ObjSmallBoxBody.Transform oTransMatrix
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSmallBoxBody
    Set ObjSmallBoxBody = Nothing
    
 ' Insert your code for output 7(Cylinder1 Body)
    stPoint.Set -parActuatorLength, parActuatorHeight3, parActuatorOffset
    enPoint.Set -parActuatorLength - parActuatorCylLength, parActuatorHeight3, parActuatorOffset
    Set ObjCylinder1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCylDiameter, True)
    ObjCylinder1Body.Transform oTransMatrix
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1Body
    Set ObjCylinder1Body = Nothing
    
 ' Insert your code for output 8(Cylinder2 Body)
    stPoint.Set parActuatorLength1, parActuatorHeight3, parActuatorOffset
    enPoint.Set parActuatorLength1 + parActuatorCyl1Length, parActuatorHeight3, parActuatorOffset
    Set ObjCylinder2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCyl1Diameter, True)
    ObjCylinder2Body.Transform oTransMatrix
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2Body
    Set ObjCylinder2Body = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oTransMatrix = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
