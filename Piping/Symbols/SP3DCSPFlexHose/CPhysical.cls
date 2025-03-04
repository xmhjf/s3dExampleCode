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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S36A Flex Hose Symbol.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const MIN_LENGTH = 0.0001
Private InchToMtr       As Double
Private PI       As Double

Private Sub Class_Initialize()
    InchToMtr = 25.4 / 1000
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
    Dim ObjStrainer As Object
    Dim ObjCollar As Object
    Dim ObjCollarExtn As Object
    Dim ObjCollarIns As Object
    
    Dim parFacetoFace As Double
    
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
    parFacetoFace = arrayOfInputs(2)    'A
    parNPD = arrayOfInputs(3)
    parEndPreparation = arrayOfInputs(4)
    parScheduleThickness = arrayOfInputs(5)
    parEndStandard = arrayOfInputs(6)
    parPressureRating = arrayOfInputs(7)
    parFlowDirection = arrayOfInputs(8)
    parPortIndex1 = arrayOfInputs(9)
    parNpd1 = arrayOfInputs(10)
    parEndPreparation1 = arrayOfInputs(11)
    parScheduleThickness1 = arrayOfInputs(12)
    parEndStandard1 = arrayOfInputs(13)
    parPressureRating1 = arrayOfInputs(14)
    parFlowDirection1 = arrayOfInputs(15)
    parPortIndex2 = arrayOfInputs(16)
    parNpd2 = arrayOfInputs(17)
    parEndPreparation2 = arrayOfInputs(18)
    parScheduleThickness2 = arrayOfInputs(19)
    parEndStandard2 = arrayOfInputs(20)
    parPressureRating2 = arrayOfInputs(21)
    parFlowDirection2 = arrayOfInputs(22)
    parId1 = arrayOfInputs(23)
    parId2 = arrayOfInputs(24)
    parNPDUnitType = arrayOfInputs(25)
    parNPDUnitType1 = arrayOfInputs(26)
    parNPDUnitType2 = arrayOfInputs(27)
    
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
    
    'Following logic for nozzle length calculation is taken from EDEN
    Dim one As Double
    Dim threenhalf As Double
    Dim dia As Double
    Dim sect As Double

    
    dia = pipeDiam(parPortIndex1)
    
    one = InchToMtr
    threenhalf = 3.5 * InchToMtr
        If (dia <= one) Then
           sect = 2# * dia
        Else
           If (dia > one And dia <= threenhalf) Then
              sect = 1# * dia
           Else
              sect = 0.5 * dia
           End If
        End If

    Dim dNozzleLength As Double

    dNozzleLength = sect + flangeThick(parPortIndex1)
    If dNozzleLength < flangeThick(parPortIndex1) Then dNozzleLength = flangeThick(parPortIndex1)
    oNozzle.Length = dNozzleLength
     
' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
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
    dNozzleLength = sect + flangeThick(parPortIndex2)
    If dNozzleLength < flangeThick(parPortIndex2) Then dNozzleLength = flangeThick(parPortIndex2)
    oNozzle.Length = dNozzleLength

' Direction of the Nozzle
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set parFacetoFace / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    
 ' Insert your code for output 3(Torus1 for representing flexible hose)
    Dim sect_1 As Double
    Dim bangle As Double
    Dim bradius As Double
    sect_1 = parFacetoFace - flangeThick(parPortIndex1) _
            - flangeThick(parPortIndex2) - sect * 2#
    bangle = PI / 2 '90 Degrees
    bradius = sect_1 * 0.25
    
    Dim MinCylDia As Double
    MinCylDia = pipeDiam(parPortIndex1) * 0.1
   
'   Check that the stem diameter is within 15mm and 50mm
   If MinCylDia < 0.015 Then MinCylDia = 0.015
   If MinCylDia > 0.05 Then MinCylDia = 0.05

'   Place circle
    Dim cirNormalX As Double
    Dim cirNormalY As Double
    Dim cirNormalZ As Double

    Dim cir1CenterX As Double
    Dim cir1CenterY As Double
    Dim cir1CenterZ As Double
    
    cirNormalX = 0
    cirNormalY = 1
    cirNormalZ = 0

    cir1CenterX = -bradius
    cir1CenterY = -bradius
    cir1CenterZ = 0
    
    Dim objCircle1 As IngrGeom3D.Circle3d
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Set objCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, MinCylDia)

    Dim oAxisVec As AutoMath.DVector
    Dim oRevCentPt As AutoMath.DPosition

    Set oAxisVec = New AutoMath.DVector
    Set oRevCentPt = New AutoMath.DPosition
    
    oAxisVec.Set 0, 0, 1
    oRevCentPt.Set -bradius * 2, -bradius, 0
    Dim ObjTorus As Object
    Set ObjTorus = PlaceRevolution(m_OutputColl, objCircle1, oAxisVec, oRevCentPt, PI / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTorus
    Set ObjTorus = Nothing
 
 ' Insert your code for output 4(Torus2)
    oRevCentPt.Set 0, -bradius, 0
    Set ObjTorus = PlaceRevolution(m_OutputColl, objCircle1, oAxisVec, oRevCentPt, PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTorus
    Set ObjTorus = Nothing
    
 ' Insert your code for output 5(Torus3)
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oTransVec As AutoMath.DVector
    
    Set oTransVec = New AutoMath.DVector
    oTransVec.Set bradius * 2, 0, 0
    oTransMat.LoadIdentity
    oTransMat.Translate oTransVec
    objCircle1.Transform oTransMat
    
    oAxisVec.Set 0, 0, -1
    oRevCentPt.Set bradius * 2, -bradius, 0
    Set ObjTorus = PlaceRevolution(m_OutputColl, objCircle1, oAxisVec, oRevCentPt, PI / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTorus
    Set ObjTorus = Nothing
    Set oPlacePoint = Nothing

    Set oDir = Nothing
    Set oGeomFactory = Nothing
    Set oAxisVec = Nothing
    Set oRevCentPt = Nothing
    Set oTransMat = Nothing
    Set m_oCodeListMetadata = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    Set objCircle1 = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
