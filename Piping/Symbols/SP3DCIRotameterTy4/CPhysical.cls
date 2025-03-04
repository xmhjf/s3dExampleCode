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
'   Author:          MS
'   Creation Date:  Friday, 3rd Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR2, PDMS Category Name=IFNH)
'   The symbol consists of four outputs, three physical and one insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   16.JAN.2008     MA   CR-113430 Implemented Part Data Basis for the values 732 and 733
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI As Double

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
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double
    
    Dim iOutput     As Double
    Dim ObjRotameterBody As Object
    Dim ObjTopCylBody As Object
    Dim ObjBottCylBody As Object
        
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parFacetoFace As Double
    Dim parPortRotation1 As Double
    Dim parFace1toCenter As Double
    Dim parPortRotation2 As Double
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
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)   'D2
    parInstrumentHeight1 = arrayOfInputs(3)  'D3
    parInstrumentDiameter = arrayOfInputs(4)      'D4
    parInstrumentLength = arrayOfInputs(5)        'D5
    parInstrumentLength1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInstrumentWidth1 = arrayOfInputs(8)       'D8
    parFacetoFace = arrayOfInputs(9)  'D9
    parPortRotation1 = arrayOfInputs(10)      'D10
    parFace1toCenter = arrayOfInputs(11)     'D11
    parPortRotation2 = arrayOfInputs(12)      'D12
    parInsulationThickness = arrayOfInputs(13)
    parNPD = arrayOfInputs(14)
    parEndPreparation = arrayOfInputs(15)
    parScheduleThickness = arrayOfInputs(16)
    parEndStandard = arrayOfInputs(17)
    parPressureRating = arrayOfInputs(18)
    parFlowDirection = arrayOfInputs(19)
    parPortIndex1 = arrayOfInputs(20)
    parNpd1 = arrayOfInputs(21)
    parEndPreparation1 = arrayOfInputs(22)
    parScheduleThickness1 = arrayOfInputs(23)
    parEndStandard1 = arrayOfInputs(24)
    parPressureRating1 = arrayOfInputs(25)
    parFlowDirection1 = arrayOfInputs(26)
    parPortIndex2 = arrayOfInputs(27)
    parNpd2 = arrayOfInputs(28)
    parEndPreparation2 = arrayOfInputs(29)
    parScheduleThickness2 = arrayOfInputs(30)
    parEndStandard2 = arrayOfInputs(31)
    parPressureRating2 = arrayOfInputs(32)
    parFlowDirection2 = arrayOfInputs(33)
    parId1 = arrayOfInputs(34)
    parId2 = arrayOfInputs(35)
    parNPDUnitType = arrayOfInputs(36)
    parNPDUnitType1 = arrayOfInputs(37)
    parNPDUnitType2 = arrayOfInputs(38)
'    parPortIndex3 = arrayOfInputs(39)
'    parNpd3 = arrayOfInputs(40)
'    parEndPreparation3 = arrayOfInputs(41)
'    parScheduleThickness3 = arrayOfInputs(42)
'    parEndStandard3 = arrayOfInputs(43)
'    parPressureRating3 = arrayOfInputs(44)
'    parFlowDirection3 = arrayOfInputs(45)
'    parPortIndex4 = arrayOfInputs(46)
'    parNpd4 = arrayOfInputs(47)
'    parEndPreparation4 = arrayOfInputs(48)
'    parScheduleThickness4 = arrayOfInputs(49)
'    parEndStandard4 = arrayOfInputs(50)
'    parPressureRating4 = arrayOfInputs(51)
'    parFlowDirection4 = arrayOfInputs(52)
'    parId3 = arrayOfInputs(53)
'    parId4 = arrayOfInputs(54)
'    parNPDUnitType3 = arrayOfInputs(55)
'    parNPDUnitType4 = arrayOfInputs(56)
'    parFace3toCenter = arrayOfInputs(57)
'    parFace4toCenter = arrayOfInputs(58)

    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    'The origin is considered to be mid of parInstrumentHeight1 and on the vertical axis
    'Insert your code for output 1(Rotameter Body)
    stPoint.Set -parInstrumentLength, -parInstrumentHeight1 / 2, parInstrumentWidth
    enPoint.Set parInstrumentLength1, parInstrumentHeight1 / 2, -parInstrumentWidth1
    Set ObjRotameterBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotameterBody
    Set ObjRotameterBody = Nothing
    
    'Insert your code for output 3(Top Cylinder Body)
    stPoint.Set 0, parInstrumentHeight1 / 2, 0
    enPoint.Set 0, parInstrumentHeight / 2, 0
    Set ObjTopCylBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylBody
    Set ObjTopCylBody = Nothing
    
    'Insert your code for output 4(Bottom Cylinder Body)
    stPoint.Set 0, -parInstrumentHeight1 / 2, 0
    enPoint.Set 0, -parInstrumentHeight / 2, 0
    Set ObjBottCylBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottCylBody
    Set ObjBottCylBody = Nothing
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    'Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
    'CreateRetrieveDynamicNozzle
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
    oNozzle.Length = parFace1toCenter

    'Direction of the Nozzle
    oDir.Set Cos(parPortRotation2), 0, -Sin(parPortRotation2)
    oDistribPort.SetDirectionVector oDir

    'Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parFace1toCenter + sptOffset(parPortIndex1) - depth(parPortIndex1)) * Cos(parPortRotation2), _
                                -parFacetoFace / 2, _
                                -(parFace1toCenter + sptOffset(parPortIndex1) - depth(parPortIndex1)) * Sin(parPortRotation2)
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing

    'Place Nozzle 2
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
    oNozzle.Length = parFace1toCenter

    'Direction of the Nozzle
    oDir.Set Cos(parPortRotation1), 0, -Sin(parPortRotation1)
    oDistribPort.SetDirectionVector oDir

    'Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set (parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2)) * Cos(parPortRotation1), _
                                parFacetoFace / 2, _
                                -(parFace1toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2)) * Sin(parPortRotation1)
    oLogicalDistPort.SetCenterLocation oPlacePoint

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -731): Level indicator, with inlet/outlet ports
    '   (Part data Basis value -732): Level indicator, with inlet/outlet ports and vent/drain
    '                                               ports, with straight outlet
    '   (Part data Basis value -733): Level indicator, with inlet/outlet ports and vent/drain
    '                                               ports, with 90o elbow outlet

    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 731

        Case 732
            parPortIndex3 = arrayOfInputs(39)
            parNpd3 = arrayOfInputs(40)
            parEndPreparation3 = arrayOfInputs(41)
            parScheduleThickness3 = arrayOfInputs(42)
            parEndStandard3 = arrayOfInputs(43)
            parPressureRating3 = arrayOfInputs(44)
            parFlowDirection3 = arrayOfInputs(45)
            parPortIndex4 = arrayOfInputs(46)
            parNpd4 = arrayOfInputs(47)
            parEndPreparation4 = arrayOfInputs(48)
            parScheduleThickness4 = arrayOfInputs(49)
            parEndStandard4 = arrayOfInputs(50)
            parPressureRating4 = arrayOfInputs(51)
            parFlowDirection4 = arrayOfInputs(52)
            parId3 = arrayOfInputs(53)
            parId4 = arrayOfInputs(54)
            parNPDUnitType3 = arrayOfInputs(55)
            parNPDUnitType4 = arrayOfInputs(56)
            parFace3toCenter = arrayOfInputs(57)
            parFace4toCenter = arrayOfInputs(58)
            
            'Place Nozzle 3
            Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNPD, parNPDUnitType, _
                                parEndPreparation, parScheduleThickness, parEndStandard, _
                                parPressureRating, parFlowDirection, parNpd3, _
                                parNPDUnitType3, parEndPreparation3, parScheduleThickness3, _
                                parEndStandard3, parPressureRating3, parFlowDirection3, _
                                parId3, m_OutputColl, m_oCodeListMetadata, _
                                True, pipeDiam, flangeThick, _
                                flangeDiam, sptOffset, depth)
                                                    
            Set oLogicalDistPort = oNozzle
            Set oDistribPort = oNozzle
            oNozzle.Length = parFace3toCenter - (parInstrumentHeight / 2)
            
            'Direction of the Nozzle
            oDir.Set 0, 1, 0
            oDistribPort.SetDirectionVector oDir
        
            'Position of the nozzle should be the connect point of the nozzle
            oPlacePoint.Set 0, parFace3toCenter + sptOffset(parPortIndex3) _
                                                - depth(parPortIndex3), 0
            
            oLogicalDistPort.SetCenterLocation oPlacePoint
        
            'Set the output
            m_OutputColl.AddOutput "VNoz3", oNozzle
            Set oNozzle = Nothing
     
            'Place Nozzle 4
            Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNPD, parNPDUnitType, _
                                parEndPreparation, parScheduleThickness, parEndStandard, _
                                parPressureRating, parFlowDirection, parNpd4, _
                                parNPDUnitType4, parEndPreparation4, parScheduleThickness4, _
                                parEndStandard4, parPressureRating4, parFlowDirection4, _
                                parId4, m_OutputColl, m_oCodeListMetadata, _
                                True, pipeDiam, flangeThick, _
                                flangeDiam, sptOffset, depth)
                                                    
            Set oLogicalDistPort = oNozzle
            Set oDistribPort = oNozzle
            oNozzle.Length = parFace4toCenter - (parInstrumentHeight / 2)
            
            'Direction of the Nozzle
            oDir.Set 0, -1, 0
            oDistribPort.SetDirectionVector oDir
        
            'Position of the nozzle should be the connect point of the nozzle
            oPlacePoint.Set 0, -parFace4toCenter - sptOffset(parPortIndex4) _
                                                + depth(parPortIndex4), 0
            
            oLogicalDistPort.SetCenterLocation oPlacePoint
        
            'Set the output
            m_OutputColl.AddOutput "VNoz4", oNozzle
            Set oNozzle = Nothing
            
        Case 733
            parPortIndex3 = arrayOfInputs(39)
            parNpd3 = arrayOfInputs(40)
            parEndPreparation3 = arrayOfInputs(41)
            parScheduleThickness3 = arrayOfInputs(42)
            parEndStandard3 = arrayOfInputs(43)
            parPressureRating3 = arrayOfInputs(44)
            parFlowDirection3 = arrayOfInputs(45)
            parPortIndex4 = arrayOfInputs(46)
            parNpd4 = arrayOfInputs(47)
            parEndPreparation4 = arrayOfInputs(48)
            parScheduleThickness4 = arrayOfInputs(49)
            parEndStandard4 = arrayOfInputs(50)
            parPressureRating4 = arrayOfInputs(51)
            parFlowDirection4 = arrayOfInputs(52)
            parId3 = arrayOfInputs(53)
            parId4 = arrayOfInputs(54)
            parNPDUnitType3 = arrayOfInputs(55)
            parNPDUnitType4 = arrayOfInputs(56)
            parFace3toCenter = arrayOfInputs(57)
            parFace4toCenter = arrayOfInputs(58)
                        
            'Place Nozzle 3
            Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex3, parNPD, parNPDUnitType, _
                                parEndPreparation, parScheduleThickness, parEndStandard, _
                                parPressureRating, parFlowDirection, parNpd3, _
                                parNPDUnitType3, parEndPreparation3, parScheduleThickness3, _
                                parEndStandard3, parPressureRating3, parFlowDirection3, _
                                parId3, m_OutputColl, m_oCodeListMetadata, _
                                True, pipeDiam, flangeThick, _
                                flangeDiam, sptOffset, depth)
                                                    
            Set oLogicalDistPort = oNozzle
            Set oDistribPort = oNozzle
            oNozzle.Length = parFace3toCenter - (parInstrumentHeight / 2)
            
            'Direction of the Nozzle
            oDir.Set 0, 1, 0
            oDistribPort.SetDirectionVector oDir
        
            'Position of the nozzle should be the connect point of the nozzle
            oPlacePoint.Set 0, parFace3toCenter + sptOffset(parPortIndex3) _
                                                - depth(parPortIndex3), 0
            
            oLogicalDistPort.SetCenterLocation oPlacePoint
        
            'Set the output
            m_OutputColl.AddOutput "VNoz3", oNozzle
            Set oNozzle = Nothing
            
            'Place Nozzle 4
            Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex4, parNPD, parNPDUnitType, _
                                parEndPreparation, parScheduleThickness, parEndStandard, _
                                parPressureRating, parFlowDirection, parNpd4, _
                                parNPDUnitType4, parEndPreparation4, parScheduleThickness4, _
                                parEndStandard4, parPressureRating4, parFlowDirection4, _
                                parId4, m_OutputColl, m_oCodeListMetadata, _
                                True, pipeDiam, flangeThick, _
                                flangeDiam, sptOffset, depth)
                                                    
            Set oLogicalDistPort = oNozzle
            Set oDistribPort = oNozzle
            oNozzle.Length = flangeThick(parPortIndex4)
            
            'Direction of the Nozzle
            oDir.Set -1, 0, 0
            oDistribPort.SetDirectionVector oDir
        
            'Position of the nozzle should be the connect point of the nozzle
            oPlacePoint.Set -(parFace4toCenter - parInstrumentHeight / 2) _
                    - sptOffset(parPortIndex4) + depth(parPortIndex4), -parFace4toCenter, 0
            
            oLogicalDistPort.SetCenterLocation oPlacePoint
        
            'Set the output
            m_OutputColl.AddOutput "VNoz4", oNozzle
            Set oNozzle = Nothing
            
            'Place elbow
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New GeometryFactory
            Dim objCircle   As IngrGeom3D.Circle3d
            Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         0, -parInstrumentHeight / 2, 0, _
                         0, -1, 0, _
                         pipeDiam(parPortIndex4) / 2)
            
            Dim CP As New AutoMath.DPosition 'revolution center point
            Dim CV As New AutoMath.DVector 'rotation vector for rotation
            CP.Set -(parFace4toCenter - parInstrumentHeight / 2), -parInstrumentHeight / 2, 0
            CV.Set 0, 0, -1
            
            Dim ObjElbow As Object
            Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                        CV, CP, PI / 2, True)
            
            'Set the output
            m_OutputColl.AddOutput "Arc", ObjElbow
            Set ObjElbow = Nothing
            Set CP = Nothing
            Set CV = Nothing
            Set objCircle = Nothing
            Set oGeomFactory = Nothing
            
        Case Else
            GoTo ErrorLabel:

    End Select
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
