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
'   Author:          SSP
'   Creation Date:  Friday, Jan 3 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'  The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRXPA).
'  Physical aspect consist 1 Cylinder, 2 Cones ,1 stem ,3 boxes,2 nozzles and 1 insulation aspect outputs.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  20 June 2003     MS      TR : Added OperationHeight as input parameter and also operation aspect to the symbol.
'                                                                  Changed the Input parameters order.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.May.2008     KKC  CR-135970  Provide ability to rotate actuator for on-the-fly control valves
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData

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
    Dim ObjValCylinder As Object
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjStem As Object
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    
    Dim parFace1toCenter As Double              'D2
    Dim parFacetoFace As Double                 'D3
    Dim parCylHeight As Double               'D4
    Dim parActuatorHeight As Double             'D5
    Dim parActuatorHeight1 As Double            'D6
    Dim parActuatorHeight2 As Double            'D7
    Dim parActuatorLength As Double             'D8
    Dim parActuatorLength1 As Double            'D9
    Dim parActuatorWidth As Double              'D10
    Dim parActuatorWidth1 As Double             'D11
    Dim parActuatorLength2 As Double            'D12
    Dim parActuatorLength3 As Double            'D13
    Dim parActuatorWidth2 As Double             'D14
    Dim parActuatorWidth3 As Double             'D15
    Dim parActuatorWidth4 As Double             'D16
    Dim parActuatorLength4 As Double            'D17
    Dim parActuatorLength5 As Double            'D18
    Dim parActuatorHeight3 As Double            'D19
    Dim parActuatorHeight4 As Double            'D20
    Dim parActuatorHeight5 As Double            'D21
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
    
    Dim parOperationHeight As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorHeight1 = arrayOfInputs(6)
    parActuatorHeight2 = arrayOfInputs(7)
    parActuatorLength = arrayOfInputs(8)
    parActuatorLength1 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorLength2 = arrayOfInputs(12)
    parActuatorLength3 = arrayOfInputs(13)
    parActuatorWidth2 = arrayOfInputs(14)
    parActuatorWidth3 = arrayOfInputs(15)
    parActuatorWidth4 = arrayOfInputs(16)
    parActuatorLength4 = arrayOfInputs(17)
    parActuatorLength5 = arrayOfInputs(18)
    parActuatorHeight3 = arrayOfInputs(19)
    parActuatorHeight4 = arrayOfInputs(20)
    parActuatorHeight5 = arrayOfInputs(21)
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
    parOperationHeight = arrayOfInputs(43)
    'Text Inputs
    parId1 = arrayOfInputs(44)
    parId2 = arrayOfInputs(45)
    parNPDUnitType = arrayOfInputs(46)
    parNPDUnitType1 = arrayOfInputs(47)
    parNPDUnitType2 = arrayOfInputs(48)
    
    If UBound(arrayOfInputs) > 48 Then
        parAngle = arrayOfInputs(49)
    End If
  
    iOutput = 0
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
' Place Nozzle 1
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
    'Assumption
'    ValCylDiam = parFace1toCenter is taken
    ValCylDiam = parFace1toCenter
    
    Set ObjValCylinder = PlaceCylinder(m_OutputColl, pPos1, pPos2, ValCylDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCylinder
    Set ObjValCylinder = Nothing
    
' Insert your code for output 6(Stem)
    Dim StemDia As Double
    pPos1.Set 0, 0, 0
    pPos2.Set 0, parActuatorHeight, 0
    ' Assumption
    ' StemDia = 0.05 are taken
    StemDia = 0.05
    
    Set ObjStem = PlaceCylinder(m_OutputColl, pPos1, pPos2, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 7(Box1)
    Dim dLineStringPoints(0 To 14) As Double
    dLineStringPoints(0) = -parActuatorLength
    dLineStringPoints(1) = parActuatorHeight
    dLineStringPoints(2) = parActuatorWidth

    dLineStringPoints(3) = -parActuatorLength
    dLineStringPoints(4) = parActuatorHeight
    dLineStringPoints(5) = -parActuatorWidth1

    dLineStringPoints(6) = parActuatorLength1
    dLineStringPoints(7) = parActuatorHeight
    dLineStringPoints(8) = -parActuatorWidth1
    
    dLineStringPoints(9) = parActuatorLength1
    dLineStringPoints(10) = parActuatorHeight
    dLineStringPoints(11) = parActuatorWidth
    
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
    Set ObjBox1 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight1, True)

    Dim oTransMatrix As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    oTransMatrix.LoadIdentity
    oTransMatrix.Rotate parAngle, oAxisVect
    ObjBox1.Transform oTransMatrix

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 8(Box2)
    dLineStringPoints(0) = -parActuatorLength2
    dLineStringPoints(1) = parActuatorHeight + parActuatorHeight1
    dLineStringPoints(2) = parActuatorWidth2
    
    dLineStringPoints(3) = -parActuatorLength2
    dLineStringPoints(4) = parActuatorHeight + parActuatorHeight1
    dLineStringPoints(5) = -parActuatorWidth3

    dLineStringPoints(6) = parActuatorLength3
    dLineStringPoints(7) = parActuatorHeight + parActuatorHeight1
    dLineStringPoints(8) = -parActuatorWidth3

    dLineStringPoints(9) = parActuatorLength3
    dLineStringPoints(10) = parActuatorHeight + parActuatorHeight1
    dLineStringPoints(11) = parActuatorWidth2

    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)
    
    Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Set ObjBox2 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight2, True)

    ObjBox2.Transform oTransMatrix

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing
    
' Insert your code for output 9(Box3)
    dLineStringPoints(0) = -parActuatorLength4
    dLineStringPoints(1) = parActuatorHeight + parActuatorHeight5 - parActuatorHeight3
    dLineStringPoints(2) = parActuatorWidth + parActuatorWidth4
    
    dLineStringPoints(3) = -parActuatorLength4
    dLineStringPoints(4) = parActuatorHeight + parActuatorHeight5 - parActuatorHeight3
    dLineStringPoints(5) = parActuatorWidth

    dLineStringPoints(6) = parActuatorLength5
    dLineStringPoints(7) = parActuatorHeight + parActuatorHeight5 - parActuatorHeight3
    dLineStringPoints(8) = parActuatorWidth

    dLineStringPoints(9) = parActuatorLength5
    dLineStringPoints(10) = parActuatorHeight + parActuatorHeight5 - parActuatorHeight3
    dLineStringPoints(11) = parActuatorWidth + parActuatorWidth4

    dLineStringPoints(12) = dLineStringPoints(0)
    dLineStringPoints(13) = dLineStringPoints(1)
    dLineStringPoints(14) = dLineStringPoints(2)
    
    Set oLineString1 = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStringPoints)

    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 0, 1, 0
    Set ObjBox3 = PlaceProjection(m_OutputColl, oLineString1, oAxisVect, parActuatorHeight3 + parActuatorHeight3, True)

    ObjBox3.Transform oTransMatrix

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing
       
    Set m_oCodeListMetadata = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set pPos1 = Nothing
    Set pPos2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
