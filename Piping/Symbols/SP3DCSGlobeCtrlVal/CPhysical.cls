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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 23 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol is prepared as per the PDS Piping Component Data Reference Guide, MC = I15AT or I15AS,
'   SN = I15AZ at page no 362 in pcd.dpf file. This Symbol has got both Physical and Insulation aspects.
'   This symbol is Custom Speciality Instrument with Nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double
    Dim flangeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim depth(1 To 2) As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjEllipsoid As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

    Dim axis            As New AutoMath.DVector
    Dim centPoint       As New AutoMath.DPosition
    Dim stPoint         As New AutoMath.DPosition
    Dim enPoint         As New AutoMath.DPosition
    Dim ellipseMajorX   As Double
    Dim ellipseMajorY   As Double
    Dim ellipseMajorZ   As Double
    Dim ellipseMMRatio  As Double
        
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
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    parNpd = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
    parId1 = arrayOfInputs(26)
    parId2 = arrayOfInputs(27)
    parNpdUnitType = arrayOfInputs(28)
    parNpdUnitType1 = arrayOfInputs(29)
    parNpdUnitType2 = arrayOfInputs(30)


    iOutput = 0

' Insert your code for output 1(Ellipsoid)
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
'   Consider Ellipse Majoj Axis is 0.6 times the larger face to center distance /2 and
'   Minor axis as 0.4 times the same center distance.
    ellipseMMRatio = 0.4 / 0.6
    If parFace1toCenter < parFace2toCenter Then
        ellipseMajorX = 0.6 * (parFace2toCenter * 0.5)
    Else
        ellipseMajorX = 0.6 * (parFace1toCenter * 0.5)
    End If
    ellipseMajorY = CenterPos.y
    ellipseMajorZ = CenterPos.z
    
    Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
    
    Set ObjEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            m_OutputColl.ResourceManager, CenterPos.x, CenterPos.y, CenterPos.z, _
                                            0, 0, 1, ellipseMajorX, ellipseMajorY, ellipseMajorZ, _
                                            ellipseMMRatio, 0, PI)

    centPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
    axis.Set 1, 0, 0
    Set ObjEllipsoid = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axis, _
                        centPoint, 2 * PI, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipsoid
    Set ObjEllipsoid = Nothing
    
'   remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = ObjEllipticalArc
    Set ObjEllipticalArc = Nothing
    sp3dElem1.Remove
    
' Place Nozzle 1
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
    oPlacePoint.Set CenterPos.x - parFace1toCenter - sptOffset(parPortIndex1) + depth(parPortIndex1), _
                    CenterPos.y, _
                    CenterPos.z
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
' Insert your code for output 2(Body Cone1)
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    stPoint.Set CenterPos.x - parFace1toCenter + flangeThick(parPortIndex1), _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x, _
                CenterPos.y, _
                CenterPos.z
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1) / 2, 0, True)
'   Set the output
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
    oPlacePoint.Set CenterPos.x + parFace2toCenter + sptOffset(parPortIndex2) - depth(parPortIndex2), _
                    CenterPos.y, _
                    CenterPos.z
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
' Insert your code for output 3(Body Cone2)
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    stPoint.Set CenterPos.x + parFace2toCenter - flangeThick(parPortIndex2), _
                CenterPos.y, _
                CenterPos.z
    enPoint.Set CenterPos.x, _
                CenterPos.y, _
                CenterPos.z
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2) / 2, 0, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
' Insert your code for output 7 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    
    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
  
        End If
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub
