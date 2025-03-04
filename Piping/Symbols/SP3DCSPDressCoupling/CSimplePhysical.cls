VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author: svsmylav
'   Creation Date:  Friday, Oct 8 2004
'
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'    This Symbol details are as per PDS EDEN S1A1.
'
'  Change History:
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const MIN_LENGTH = 0.0001


Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double

    Dim oPart As IJDPart
    Dim oPipeComponent As IJDPipeComponent
    
    Set oPart = arrayOfInputs(1)
'    Set oPipeComponent = oPart
    
'    Dim oNozzleData() As NozzleData
'    NozzleInitialize oPart, oNozzleData
    
'    Dim oNozzleCol As Collection
'    Set oNozzleCol = New Collection
    
    Dim oOutputCol As Collection
    Set oOutputCol = New Collection
    
    Dim oT4x4Current As IJDT4x4
    Set oT4x4Current = New DT4x4
    
    Dim oT4x4Temp As IJDT4x4
    Set oT4x4Temp = New DT4x4
    
    Dim oGeometryFactory As GeometryFactory
    Set oGeometryFactory = New GeometryFactory

    Dim oOriginPos As DPosition
    Set oOriginPos = New DPosition
    oOriginPos.Set 0, 0, 0
    
'**********************************************************************************
'**********************************************************************************
'   BEGIN
'**********************************************************************************
'**********************************************************************************
    Dim parFacetoFace As Double
    Dim parCylinderDiameter As Double

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
    parFacetoFace = arrayOfInputs(2) 'A
    parCylinderDiameter = arrayOfInputs(3) 'B
    parNPD = arrayOfInputs(4)
    parEndPreparation = arrayOfInputs(5)
    parScheduleThickness = arrayOfInputs(6)
    parEndStandard = arrayOfInputs(7)
    parPressureRating = arrayOfInputs(8)
    parFlowDirection = arrayOfInputs(9)
    parPortIndex1 = arrayOfInputs(10)
    parNpd1 = arrayOfInputs(11)
    parEndPreparation1 = arrayOfInputs(12)
    parScheduleThickness1 = arrayOfInputs(13)
    parEndStandard1 = arrayOfInputs(14)
    parPressureRating1 = arrayOfInputs(15)
    parFlowDirection1 = arrayOfInputs(16)
    parPortIndex2 = arrayOfInputs(17)
    parNpd2 = arrayOfInputs(18)
    parEndPreparation2 = arrayOfInputs(19)
    parScheduleThickness2 = arrayOfInputs(20)
    parEndStandard2 = arrayOfInputs(21)
    parPressureRating2 = arrayOfInputs(22)
    parFlowDirection2 = arrayOfInputs(23)
    parId1 = arrayOfInputs(24)
    parId2 = arrayOfInputs(25)
    parNPDUnitType = arrayOfInputs(26)
    parNPDUnitType1 = arrayOfInputs(27)
    parNPDUnitType2 = arrayOfInputs(28)

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
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = flangeThick(parPortIndex1)

' Direction of the Nozzle
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset(parPortIndex1) + depth(parPortIndex1), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    oOutputCol.Add oNozzle
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
    oPlacePoint.Set parFacetoFace / 2 + sptOffset(parPortIndex2) - depth(parPortIndex2), 0, 0
    oLogicalDistPort.SetCenterLocation oPlacePoint

' Set the output
    oOutputCol.Add oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
'        NozzleDefineDirectionVector oNozzleData, NozzleGeometryType_Linear
        Dim stPoint As AutoMath.DPosition
        Dim enPoint As AutoMath.DPosition
        
        Set stPoint = New AutoMath.DPosition
        Set enPoint = New AutoMath.DPosition
        
        Dim f_to_f_dim As Variant
        f_to_f_dim = parFacetoFace

        Dim cyl_dia As Variant
        cyl_dia = parCylinderDiameter

        Dim f_to_c_dim_1 As Variant
        f_to_c_dim_1 = f_to_f_dim * 0.5

        Dim f_to_c_dim_2 As Variant
        f_to_c_dim_2 = f_to_c_dim_1

        Dim sec1 As Variant
        'Ensure that cylinder length (sec1) is greater than zero
        If flangeThick(parPortIndex1) > MIN_LENGTH Then
            sec1 = flangeThick(parPortIndex1) * 2#
        Else
            sec1 = 2 * MIN_LENGTH
        End If

        Dim sec2 As Variant
        'Ensure that cylinder length (sec1) is greater than zero
        If flangeThick(parPortIndex2) > MIN_LENGTH Then
        sec2 = flangeThick(parPortIndex2) * 2#
        Else
            sec2 = 2 * MIN_LENGTH
        End If

        Dim sec As Variant
'        sec = f_to_f_dim - oNozzleData(1).dFlangeThickness - oNozzleData(2).dFlangeThickness - sec1 - sec2
        sec = f_to_f_dim - flangeThick(parPortIndex1) - flangeThick(parPortIndex2) - sec1 - sec2

'        Dim dist As Variant
'        dist = -(f_to_c_dim_2 + dimension_9)

'        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP1, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(1).dCptOffset
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(1).dDepth
'
'        Dim oCylinderCapped1 As Object
'        Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dFlangeDiameter), True)
'
'        oCylinderCapped1.Transform oT4x4Current
'
'        oOutputCol.Add oCylinderCapped1
        
'        oT4x4Current.MultMatrix oT4x4Temp
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = sec1
        
        stPoint.Set -parFacetoFace / 2 + flangeThick(parPortIndex1), 0, 0
        enPoint.Set stPoint.x + sec1, stPoint.y, stPoint.z
        
        Dim oCylinderCapped2 As Object
        Set oCylinderCapped2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex1), True)

'        oCylinderCapped2.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped2
        
'        oT4x4Current.MultMatrix oT4x4Temp
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = sec
        stPoint.Set -parFacetoFace / 2 + flangeThick(parPortIndex1) + sec1, 0, 0
        enPoint.Set stPoint.x + sec, stPoint.y, stPoint.z
        
        Dim oCylinderCapped3 As Object
        Set oCylinderCapped3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parCylinderDiameter, True)
'        oCylinderCapped3.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped3
        
'        oT4x4Current.MultMatrix oT4x4Temp
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = sec2
        stPoint.Set parFacetoFace / 2 - flangeThick(parPortIndex2) - sec2, 0, 0
        enPoint.Set stPoint.x + sec2, stPoint.y, stPoint.z
    
        Dim oCylinderCapped4 As Object
        Set oCylinderCapped4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam(parPortIndex2), True)
        
'        oCylinderCapped4.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped4
        
'        oT4x4Current.MultMatrix oT4x4Temp
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(2).dDepth
'
'        Dim oCylinderCapped5 As Object
'        Set oCylinderCapped5 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dFlangeDiameter), True)
'
'        oCylinderCapped5.Transform oT4x4Current
'
''        oOutputCol.Add oCylinderCapped5
'
'        oT4x4Current.MultMatrix oT4x4Temp
'
'
'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(2).dCptOffset
'
'
'        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP2, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
'
'
'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, dist
'
'
'        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP0, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        





    
'**********************************************************************************
'**********************************************************************************
'   END
'**********************************************************************************
'**********************************************************************************
    
'    Dim oNozzle As IJDNozzle
'    Dim lNozzleId As Long
'    Dim lNozzleIndex As Long
'
'    Dim oNozzlePos As DPosition
'
'    Dim oNozzleVec As DVector
'    Set oNozzleVec = New DVector
'
'    'place the nozzles & add to the output collection
'    For lNozzleIndex = 1 To oNozzleCol.Count
'        lNozzleId = oNozzleCol.Item(lNozzleIndex)
'
'        'move the nozzle in the opposite direction by the flange thickness
'        oNozzleVec.Set -oNozzleData(lNozzleId).oDirectionVector.x, -oNozzleData(lNozzleId).oDirectionVector.y, -oNozzleData(lNozzleId).oDirectionVector.z
'        oNozzleVec.Length = oNozzleData(lNozzleId).dFlangeThickness
'
'        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone.Offset(oNozzleVec)
'
'        Set oNozzle = CreateNozzle(CInt(lNozzleId), oPart, m_OutputColl, oNozzleData(lNozzleId).oDirectionVector, oNozzlePos)
'
'        oOutputCol.Add oNozzle
'    Next lNozzleIndex
'
    Dim lOutputIndex As Long
'
    'add all outputs
    For lOutputIndex = 1 To oOutputCol.Count
        m_OutputColl.AddOutput arrayOfOutputs(lOutputIndex), oOutputCol.Item(lOutputIndex)
    Next lOutputIndex
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

