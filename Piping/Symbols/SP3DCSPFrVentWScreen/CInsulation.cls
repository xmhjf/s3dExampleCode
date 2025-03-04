VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Oct 21 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is PDS on-the-fly S1A6 S11B Custom Speciality Free Vent with Screen symbol.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   11.Apr.2005     svsmylav        TR 74682: Cylinder length was zero if both flange thickness and insulation thickness were zero.
'                                           Fix adds a a small value(0.0001) to ensure non-zero cylinder length.
'  08.SEP.2006      KKC             DI-95670  Replace names with initials in all revision history sheets and symbols
'   7.Nov.2007      RRK             TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                   in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI           As Double
Private Const MODULE = "CInsulation:" 'Used for error messages
Const MIN_LENGTH = 0.0001

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
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjFreeeVentBody As Object
    
    Dim parFacetoCenter As Double
    Dim parVentRadius As Double
    Dim VentAngle As Double
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'A
    parVentRadius = arrayOfInputs(3)    'B
    parInsulationThickness = arrayOfInputs(4)
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parId1 = arrayOfInputs(18)
    parNPDUnitType = arrayOfInputs(19)
    parNPDUnitType1 = arrayOfInputs(20)
    
    iOutput = 0

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
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
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
' Insert your code for output 3(Nozzle 1 Insulation)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    Dim dPipeLength As Double
    Dim dInsulationDiameter As Double

    dInsulationDiameter = flangeDiam(parPortIndex1) + 2 * parInsulationThickness
    If CmpDblLessThan(flangeDiam(parPortIndex1), pipeDiam(parPortIndex1)) Then dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness

    dPipeLength = parFacetoCenter - parVentRadius
    If CmpDblLessThan(dPipeLength, MIN_LENGTH) Then dPipeLength = MIN_LENGTH

    stPoint.Set 0, 0, 0
    enPoint.Set flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick(parPortIndex1), 0) Then
        enPoint.Set 0.0001, 0, 0
        dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    End If
    
    Dim ObjNozzle1Ins As Object
    Set ObjNozzle1Ins = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1Ins
    Set ObjNozzle1Ins = Nothing

 ' Insert your code for output 4(Nozzle1BodyIns)
    stPoint.Set flangeThick(parPortIndex1) + parInsulationThickness, 0, 0
    enPoint.Set dPipeLength, 0, 0
    
    'Providing a check on start point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick(parPortIndex1), 0) Then
        stPoint.Set 0.0001, 0, 0
        dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    End If

    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness

    Dim ObjNozzle1BodyIns As Object
    Set ObjNozzle1BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1BodyIns
    Set ObjNozzle1BodyIns = Nothing
    
 ' Insert your code for output 4(FreeVentBodyIns)
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = parFacetoCenter - parVentRadius
    cir1CenterY = 0
    cir1CenterZ = 0
    
    'Check if dInsulationDiameter/2  exceeds parVentRadius and adjust it
    dInsulationDiameter = pipeDiam(parPortIndex1) + 2 * parInsulationThickness
    If (dInsulationDiameter / 2) > parVentRadius Then dInsulationDiameter = parVentRadius - MIN_LENGTH
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, dInsulationDiameter / 2)
 
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
 
    centPoint.Set parFacetoCenter - parVentRadius, -parVentRadius, 0
    axis.Set 0, 0, -1
    
    VentAngle = PI
    Set ObjFreeeVentBody = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, VentAngle, _
                                    True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFreeeVentBody
    Set ObjFreeeVentBody = Nothing
    
    Set geomFactory = Nothing
    Set centPoint = Nothing
    Set axis = Nothing

    ' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
