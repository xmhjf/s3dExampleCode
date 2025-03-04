Attribute VB_Name = "EDEN2SP3D"
'**********************************************************************************
'   Intergraph Corporation. All Rights Reserved.
'
'   File:   EDEN2SP3D.bas
'   Author: SS
'   Date:   Jul/23/2004
'
'   Description:
'
'   Change History:
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Public Const E_FAIL = &H80004005

Public Const PI As Double = 3.141592654

'nozzle data
Type NozzleData
    dNpd As Double
    strNpdUnitType As String
    
    dPipeDiameter As Double
    dFlangeDiameter As Double
    dFlangeThickness As Double
    dCptOffset As Double
    dDepth As Double
    
    lScheduleThickness As Long
    lEndPreparation As Long
    lPressureRating As Long
    lEndStandard As Long
    
    lFlowDirection As Long
    
    oPlacementPoint As DPosition
    oDirectionVector As DVector
    oT4x4 As IJDT4x4
    
    AxisDirection_CPPrimary As Long
    AxisDirection_CPSecondary As Long
    AxisDirection_CPNormal As Long
End Type

'EDEN geometry type
Public Const NozzleGeometryType_Linear As Long = 1
Public Const NozzleGeometryType_Elbolet As Long = 2
Public Const NozzleGeometryType_Bend As Long = 3
Public Const NozzleGeometryType_AngleValve As Long = 4
Public Const NozzleGeometryType_EccReducer As Long = 5
Public Const NozzleGeometryType_BranchTee As Long = 6
Public Const NozzleGeometryType_BranchLat As Long = 7
Public Const NozzleGeometryType_BranchWye As Long = 8
Public Const NozzleGeometryType_Branch2Wye As Long = 9
Public Const NozzleGeometryType_Cross As Long = 10
Public Const NozzleGeometryType_Operator As Long = 11
Public Const NozzleGeometryType_Latrolet As Long = 12
Public Const NozzleGeometryType_Olet As Long = 13
Public Const NozzleGeometryType_SingleCP As Long = 14
Public Const NozzleGeometryType_RET180 As Long = 15
Public Const NozzleGeometryType_OrificeFlange As Long = 16
Public Const NozzleGeometryType_GenericComponent As Long = 17

'nozzles (EDEN connect points)
Public Const NozzleId_CP0 As Long = 0
Public Const NozzleId_CP1 As Long = 1
Public Const NozzleId_CP2 As Long = 2
Public Const NozzleId_CP3 As Long = 3
Public Const NozzleId_CP4 As Long = 4
Public Const NozzleId_CP5 As Long = 5

'the 3 axis of direction for drawing
Public Const AxisDirection_Primary As Long = 1
Public Const AxisDirection_Secondary As Long = 2
Public Const AxisDirection_Normal As Long = 3

Public Const AxisDirection_Separator As Long = 10

Public Const DELTA_TOLERANCE As Double = 0.00001

'initializes the nozzle data
Public Function NozzleInitialize(ByRef oPart As IJDPart, ByRef oNozzleData() As NozzleData)
'''    Dim oPipePort As IJDPipePort
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    
    Set oCollection = oPart.GetNozzles
    
    Dim ii As Long
    Dim lCount As Long
    
    lCount = oCollection.Size
    
    ReDim oNozzleData(0 To lCount) As NozzleData
    
    For ii = 1 To lCount
        Set oPipePort = oCollection.Item(ii)
    
        oNozzleData(ii).dNpd = oPipePort.NPD
        oNozzleData(ii).strNpdUnitType = oPipePort.NPDUnitType
        
        oNozzleData(ii).dPipeDiameter = oPipePort.PipingOutsideDiameter
        oNozzleData(ii).dFlangeDiameter = oPipePort.FlangeOrHubOutsideDiameter
        oNozzleData(ii).dFlangeThickness = oPipePort.FlangeOrHubThickness
        oNozzleData(ii).dCptOffset = oPipePort.FlangeProjectionOrSocketOffset
        oNozzleData(ii).dDepth = oPipePort.SeatingOrGrooveOrSocketDepth
                    
        oNozzleData(ii).lScheduleThickness = oPipePort.ScheduleThickness
        oNozzleData(ii).lEndPreparation = oPipePort.EndPreparation
        oNozzleData(ii).lPressureRating = oPipePort.PressureRating
        oNozzleData(ii).lEndStandard = oPipePort.EndStandard
        
        oNozzleData(ii).lFlowDirection = oPipePort.FlowDirection
    
        Set oNozzleData(ii).oPlacementPoint = New DPosition
        Set oNozzleData(ii).oDirectionVector = New DVector
        Set oNozzleData(ii).oT4x4 = New DT4x4
        
        oNozzleData(ii).AxisDirection_CPPrimary = AxisDirection_Separator * ii + AxisDirection_Primary
        oNozzleData(ii).AxisDirection_CPSecondary = AxisDirection_Separator * ii + AxisDirection_Secondary
        oNozzleData(ii).AxisDirection_CPNormal = AxisDirection_Separator * ii + AxisDirection_Normal
    Next
    
    Set oPipePort = Nothing
    Set oCollection = Nothing
End Function

'initializes the nozzle direction vector(s) depending on the EDEN geometry type
Public Function NozzleDefineDirectionVector(ByRef oNozzleData() As NozzleData, lGeometryType As Long)
    Select Case lGeometryType
        Case NozzleGeometryType_Linear
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_Elbolet
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_AngleValve
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_EccReducer
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_BranchTee
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_BranchLat
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_BranchWye
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_Branch2Wye
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
            oNozzleData(4).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_Cross
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
            oNozzleData(4).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_Operator
        
        Case NozzleGeometryType_Latrolet
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_Olet
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_SingleCP
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
        
        Case NozzleGeometryType_RET180
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_OrificeFlange
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
        
        Case NozzleGeometryType_GenericComponent
            oNozzleData(1).oDirectionVector.Set -1, 0, 0
            oNozzleData(2).oDirectionVector.Set 1, 0, 0
            oNozzleData(3).oDirectionVector.Set 1, 0, 0
            oNozzleData(4).oDirectionVector.Set 1, 0, 0
            
        Case Else
            'error
    End Select
End Function

'places a nozzle (including the origin with nozzleid 0)
Public Sub NozzlePlace(ByRef m_OutputColl As Object, ByRef oPart As IJDPart, ByRef oNozzleData() As NozzleData, ByVal lNozzleId As Long, oPlacementPoint As IJDPosition, ByRef oT4x4Current As IJDT4x4, ByRef oNozzleCol As Collection, ByRef oOutputCol As Collection)
    Dim iNozzleID As Integer
    iNozzleID = lNozzleId
    
    Set oNozzleData(lNozzleId).oT4x4 = oT4x4Current.Clone
    Set oNozzleData(lNozzleId).oPlacementPoint = oNozzleData(lNozzleId).oT4x4.TransformPosition(oPlacementPoint.Clone)
        
    If iNozzleID <> 0 Then
        Set oNozzleData(lNozzleId).oDirectionVector = oNozzleData(lNozzleId).oT4x4.TransformVector(oNozzleData(lNozzleId).oDirectionVector.Clone)
        oNozzleCol.Add iNozzleID
    Else
        'get the invert for the origin matrix
        Dim oT4x4Invert As IJDT4x4
        Set oT4x4Invert = oNozzleData(0).oT4x4.Clone
        oT4x4Invert.Invert
        
        Dim ii As Long
        Dim lTempNozzleId As Long
        
        'transform all the outputs already placed w.r.t. the origin
        For ii = 1 To oOutputCol.Count
            oOutputCol.Item(ii).Transform oT4x4Invert
        Next ii
   
        'transform all nozzles already placed w.r.t. the origin
        For ii = 1 To oNozzleCol.Count
            lTempNozzleId = oNozzleCol.Item(ii)
        
            Set oNozzleData(lTempNozzleId).oPlacementPoint = oT4x4Invert.TransformPosition(oNozzleData(lTempNozzleId).oPlacementPoint.Clone)
            Set oNozzleData(lTempNozzleId).oDirectionVector = oT4x4Invert.TransformVector(oNozzleData(lTempNozzleId).oDirectionVector.Clone)
            oNozzleData(lTempNozzleId).oT4x4.MultMatrix oT4x4Invert
        Next ii
        
        'set the origin on nozzle 0
        oNozzleData(lNozzleId).oPlacementPoint.Set 0, 0, 0
        
        'load zero values for 12, 13, 14 indices of nozzle zero
        oNozzleData(0).oT4x4.IndexValue(12) = 0
        oNozzleData(0).oT4x4.IndexValue(13) = 0
        oNozzleData(0).oT4x4.IndexValue(14) = 0
        
        'load zero values for 12, 13, 14 indices of current transformation vector
        oT4x4Current.IndexValue(12) = 0
        oT4x4Current.IndexValue(13) = 0
        oT4x4Current.IndexValue(14) = 0
    End If
End Sub

'this function moves the current drawing position along the given axis by the given distance
Public Function MoveAlongAxis(ByRef oT4x4Current As IJDT4x4, ByVal lAxisDirection As Long, ByVal lDistance As Double)
    Dim oT4x4Temp As IJDT4x4
    Set oT4x4Temp = New DT4x4
    
    oT4x4Temp.LoadIdentity
    
    Select Case lAxisDirection
        Case AxisDirection_Primary
            oT4x4Temp.IndexValue(12) = lDistance
        Case AxisDirection_Secondary
            oT4x4Temp.IndexValue(13) = lDistance
        Case AxisDirection_Normal
            oT4x4Temp.IndexValue(14) = lDistance
        Case Else
            'error
    End Select
    
    oT4x4Current.MultMatrix oT4x4Temp
End Function

'this function rotates the drawing position along the given axis by the given angle
Public Function RotateOrientation(ByRef oT4x4Current As IJDT4x4, ByVal lAxis As Long, ByVal dAngle As Double)
    Dim oVector As DVector
    Set oVector = New DVector
    
    Select Case lAxis
        Case AxisDirection_Primary
            oVector.Set 1, 0, 0
        Case AxisDirection_Secondary
            oVector.Set 0, 1, 0
        Case AxisDirection_Normal
            oVector.Set 0, 0, 1
        Case Else
            'error
    End Select
    
    oT4x4Current.Rotate dAngle, oVector
End Function

'sets the active orientation w.r.t. to a nozzle
Public Function DefineActiveOrientation(ByRef oNozzleData() As NozzleData, ByRef oT4x4Current As IJDT4x4, ByVal lPrimary As Long, ByVal lSecondary As Long)
    Dim dAngle1 As Double
    Dim dAngle2 As Double
    
    Dim oVectorX As DVector
    Dim oVectorY As DVector
    Dim oVectorZ As DVector
        
    Set oVectorX = New DVector
    Set oVectorY = New DVector
    Set oVectorZ = New DVector
    
    oVectorX.Set 1, 0, 0
    oVectorY.Set 0, 1, 0
    oVectorZ.Set 0, 0, 1
    
    Dim ii As Long
    
    For ii = 1 To UBound(oNozzleData)
        If lPrimary = oNozzleData(ii).AxisDirection_CPPrimary Or lPrimary = oNozzleData(ii).AxisDirection_CPSecondary Or lPrimary = oNozzleData(ii).AxisDirection_CPNormal And _
            lSecondary = oNozzleData(ii).AxisDirection_CPPrimary Or lSecondary = oNozzleData(ii).AxisDirection_CPSecondary Or lSecondary = oNozzleData(ii).AxisDirection_CPNormal Then
            
            'load the nozzle matrix. since we have to keep the active point, we don't include indices 12, 13, 14.
            oT4x4Current.IndexValue(0) = oNozzleData(ii).oT4x4.IndexValue(0)
            oT4x4Current.IndexValue(1) = oNozzleData(ii).oT4x4.IndexValue(1)
            oT4x4Current.IndexValue(2) = oNozzleData(ii).oT4x4.IndexValue(2)
            oT4x4Current.IndexValue(3) = oNozzleData(ii).oT4x4.IndexValue(3)
            oT4x4Current.IndexValue(4) = oNozzleData(ii).oT4x4.IndexValue(4)
            oT4x4Current.IndexValue(5) = oNozzleData(ii).oT4x4.IndexValue(5)
            oT4x4Current.IndexValue(6) = oNozzleData(ii).oT4x4.IndexValue(6)
            oT4x4Current.IndexValue(7) = oNozzleData(ii).oT4x4.IndexValue(7)
            oT4x4Current.IndexValue(8) = oNozzleData(ii).oT4x4.IndexValue(8)
            oT4x4Current.IndexValue(9) = oNozzleData(ii).oT4x4.IndexValue(9)
            oT4x4Current.IndexValue(10) = oNozzleData(ii).oT4x4.IndexValue(10)
            oT4x4Current.IndexValue(11) = oNozzleData(ii).oT4x4.IndexValue(11)
            oT4x4Current.IndexValue(15) = oNozzleData(ii).oT4x4.IndexValue(15)
            
            If lPrimary = oNozzleData(ii).AxisDirection_CPPrimary And lSecondary = oNozzleData(ii).AxisDirection_CPSecondary Then
                'do nothing
            ElseIf lPrimary = oNozzleData(ii).AxisDirection_CPPrimary And lSecondary = oNozzleData(ii).AxisDirection_CPNormal Then
                oT4x4Current.Rotate 1.5 * PI, oVectorX
            ElseIf lPrimary = oNozzleData(ii).AxisDirection_CPSecondary And lSecondary = oNozzleData(ii).AxisDirection_CPPrimary Then
                oT4x4Current.Rotate 1.5 * PI, oVectorZ
                oT4x4Current.Rotate PI, oVectorX
            ElseIf lPrimary = oNozzleData(ii).AxisDirection_CPSecondary And lSecondary = oNozzleData(ii).AxisDirection_CPNormal Then
                oT4x4Current.Rotate 1.5 * PI, oVectorZ
                oT4x4Current.Rotate 1.5 * PI, oVectorX
            ElseIf lPrimary = oNozzleData(ii).AxisDirection_CPNormal And lSecondary = oNozzleData(ii).AxisDirection_CPPrimary Then
                oT4x4Current.Rotate 0.5 * PI, oVectorY
                oT4x4Current.Rotate 0.5 * PI, oVectorX
            ElseIf lPrimary = oNozzleData(ii).AxisDirection_CPNormal And lSecondary = oNozzleData(ii).AxisDirection_CPSecondary Then
                oT4x4Current.Rotate 0.5 * PI, oVectorY
            Else
                'error
            End If
            
            Exit For
        End If
    Next
End Function
