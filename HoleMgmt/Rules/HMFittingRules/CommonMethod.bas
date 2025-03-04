Attribute VB_Name = "CommonMethod"
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: CommonSelection.bas
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: Common function
'
' Description:
'
' History:
'       08/12/04 Suresh  Modified GetRequestApp() and added method
'                        GetNominalDiaOfConduit()for TR 34661
'********************************************************************

Option Explicit

Private m_oErrors As New IMSErrorLog.JServerErrors
Private Const Module = "HMFittingSelectionRules.CommonMethod"

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: GetNameOfSmartItem
'
'Interface: Public function
'
'Abstract: Get the name of smart item to select the proper cable way coaming
'
'Attention : Some part should be define in Catalog.
'
'********************************************************************
Public Function GetNameOfSmartItem(oHoleTrace As IJHoleTraceAE) As String
    Const METHOD = "GetNameOfSmartItem"
    On Error GoTo ErrorHandler
   
    Dim oHoleSmartOcc As IJHoleSmartOcc
    Set oHoleSmartOcc = oHoleTrace.GetSmartOccurrence
    Dim oSmartOccurrence As IJSmartOccurrence
    Set oSmartOccurrence = oHoleSmartOcc
    
    'Get the name of Part
    GetNameOfSmartItem = oSmartOccurrence.Item
        
Cleanup:
    Set oHoleSmartOcc = Nothing
    Set oSmartOccurrence = Nothing
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetRequestApp
'
' Abstract: Get the Requesting Application
'
' Description: Returns the value indicating which application requests hole.
'******************************************************************************
Public Function GetRequestApp(oOutfitting As IJDObjectCollection) As HMHoleRequestApp
    Const METHOD = "GetRequestApp"
    On Error GoTo ErrorHandler
    
    Dim PreviousApp As HMHoleRequestApp
    Dim CurrentApp As HMHoleRequestApp
    Dim oObj As Object
    Dim bFirst As Boolean
    
    bFirst = True
    GetRequestApp = HM_UnknownApp
        
    If oOutfitting.Count = 0 Then
        CurrentApp = HM_StandAloneApp
    Else
        For Each oObj In oOutfitting
            If TypeOf oObj Is IJRtePipePathFeat Then
                CurrentApp = HM_PipingApp
            ElseIf TypeOf oObj Is IJRteDuctPathFeat Then
                CurrentApp = HM_HVACApp
            ElseIf TypeOf oObj Is IJRteCablewayPathFeat Then
                CurrentApp = HM_CableApp
            ElseIf TypeOf oObj Is IJRteConduitPathFeat Then
                CurrentApp = HM_ConduitApp
            ElseIf TypeOf oObj Is IJEquipment Then
                CurrentApp = HM_EquipmentApp
            Else
                CurrentApp = HM_OtherApp
            End If
            If bFirst = False Then
                If PreviousApp <> CurrentApp Then
                    CurrentApp = HM_OtherApp
                    Exit For
                End If
            End If
            PreviousApp = CurrentApp
            bFirst = False
            Set oObj = Nothing
        Next oObj
    End If
    
    GetRequestApp = CurrentApp

Cleanup:
    Set oObj = Nothing
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, Module & " - " & METHOD, Err.Description
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetOrientationAngle
'
' Abstract: Get the cross section's orientation angle
'
' Description: Orientation angle means the angle how much the CableWay itself
'              rotates at the same position. oUVector and oVVector are
'              the vectors of cross section.
'******************************************************************************
Public Function GetOrientationAngle(oUVector As IJDVector, oVVector As IJDVector) As Double
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oNormal As New DVector
    Dim oNewUVector As New DVector
    Dim oPlane As IJPlane
    
    Dim Ux As Double, Uy As Double, Uz As Double
    Dim Nx As Double, Ny As Double, Nz As Double
        
    oUVector.Get Ux, Uy, Uz
    
    'get normal vector of uvector and vvector
    Set oNormal = oUVector.Cross(oVVector)
    oNormal.Get Nx, Ny, Nz
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
    'If new plane is created, uvector and vvector are changed to be mapped to the origin axis.
    Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, Ux, Uy, Uz, Nx, Ny, Nz)
    oPlane.GetUDirection Ux, Uy, Uz
    
    'create a comparing uvector by using the previous uvector
    oNewUVector.Set Ux, Uy, Uz
    
    'get angle
    GetOrientationAngle = oNewUVector.Angle(oUVector, oNormal)
    
    Set oGeomFactory = Nothing
    Set oNormal = Nothing
    Set oNewUVector = Nothing
    Set oPlane = Nothing
End Function

'******************************************************************************
' Routine: GetWidthDepthOfCableway
'
' Abstract: Get width and depth of Cable way
'******************************************************************************
Public Sub GetWidthDepthOfCableway(ByVal oHoleTrace As IJHoleTraceAE, _
                                   eSelectedType As CrossSectionShapeTypes, _
                                   dWidth As Double, _
                                   dDepth As Double)
    'Get the Outfitting, Only Cable Way should be taken.
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    
    Dim oObj As Object
    Dim oRtePathFeat As IJRtePathFeat
    
    Dim dCableRad As Double
    Dim bOuterDia As Boolean
    
    Const PI = 3.14159265
    
    'For Loop input object ( 1 .. icount .. n ),
    'In case of CW, only one CW should be taken and hook up the proper GSCAD symbol file.
    For Each oObj In oOutfitting
        If TypeOf oObj Is IJRteCablewayPathFeat Then
            Set oRtePathFeat = oObj
            
            Dim oRteFeatUtility As IJRtePathCrossSectUtility
            Set oRteFeatUtility = oRtePathFeat
            
            If Not oRteFeatUtility Is Nothing Then
                oRteFeatUtility.GetCrossSectionData Nothing, _
                                                    False, _
                                                    eSelectedType, _
                                                    dWidth, _
                                                    dDepth, _
                                                    dCableRad, _
                                                    bOuterDia
                               
            End If
            Set oRteFeatUtility = Nothing
            Set oRtePathFeat = Nothing
            Exit For
        End If
    Next oObj
    
    Set oObj = Nothing
    Set oOutfitting = Nothing
End Sub

'******************************************************************************
' Routine: GetStructInfo
'
' Abstract: Get the information of struct like plate type and tightness
'
' Description: the plate type and tightness can be gotten via IJPort.
'              If there is some problem to get the tightness and type, set the
'              default value as Nontight and DeckPlate.
'
'******************************************************************************
Public Sub GetStructInfo(ByVal oHoleTrace As IJHoleTraceAE, _
                         ePlateType As StructPlateType, _
                         ePlateTightness As StructPlateTightness)
    ' Get the structure
    Dim oStructure As Object
    Set oStructure = oHoleTrace.GetParentStructure

    If (TypeOf oStructure Is IJPlate) Then
         Dim oPlate As IJPlate
         Set oPlate = oStructure
         ePlateType = oPlate.plateType
         ePlateTightness = oPlate.Tightness
         Set oPlate = Nothing
    Else
         ePlateTightness = NonTight
         ePlateType = DeckPlate
    End If
    
    Set oStructure = Nothing
End Sub

'******************************************************************************
' Routine: GetOuterDiaOfConduit
'
' Abstract: Get Outer diameter of the Conduit feature
'******************************************************************************
Public Function GetNominalDiaofConduit(oHoleTrace As IJHoleTraceAE) As Double
    Dim oObject As Object
    Dim oOutfitting As IJDObjectCollection
    Dim oConduitPathFeat As IJRteConduitPathFeat
    Dim oUomVBInterface As UomVBInterface
        
    'Get the Collection of Outfittings from Holetrace
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    
    'Loop through the Collection
    For Each oObject In oOutfitting
        'If the type of object is IJRteConduitPathFeat
        If TypeOf oObject Is IJRteConduitPathFeat Then
            'Set the Object
            Set oConduitPathFeat = oObject
            Set oUomVBInterface = New UomVBInterface
            
            GetNominalDiaofConduit = oUomVBInterface.ConvertUnitToUnit(UNIT_DISTANCE, _
                                                                       oConduitPathFeat.Ncd, _
                                                                       oUomVBInterface.GetUnitId(UNIT_DISTANCE, _
                                                                                                 oConduitPathFeat.NcdUnitType), _
                                                                       DISTANCE_MILLIMETER)
            
            Set oUomVBInterface = Nothing
            Set oConduitPathFeat = Nothing
                  
            Exit For
        End If
    Next
    
    'Clean up
    Set oObject = Nothing
    Set oOutfitting = Nothing
    Set oConduitPathFeat = Nothing
End Function

'******************************************************************************
' Routine: GetResourceMgrFromObject
'
' Abstract: return the resource manager from an object
'******************************************************************************
Public Function GetResourceMgrFromObject(ByVal pObject As Object) As Object
    On Error Resume Next
    
    Dim oJDObject As IJDObject
    Dim oResourceMgr As IUnknown
    
    Set oJDObject = pObject
    If Not oJDObject Is Nothing Then
        Set oResourceMgr = oJDObject.ResourceManager
    
        If Not oResourceMgr Is Nothing Then
            Set GetResourceMgrFromObject = oResourceMgr
            Set oResourceMgr = Nothing
        End If
        
        Set oJDObject = Nothing
    End If
        
    Exit Function
End Function

'******************************************************************************
' Routine: AdjustFittingToAlongPipeSleeve
'
' Abstract: Adjust Fitting transformation matrix such that the Pipe Sleeve is
'           orientated to be along the Pipe
'******************************************************************************
Public Sub AdjustFittingToAlongPipeSleeve(oHoleTraceAE As IJHoleTraceAE, _
                                          oPartOcc As IJPartOcc, _
                                          oMatrix As IJDT4x4)
    On Error Resume Next
    
    Dim dX1 As Double
    Dim dY1 As Double
    Dim dZ1 As Double
    Dim dX2 As Double
    Dim dY2 As Double
    Dim dZ2 As Double
    Dim dDot As Double
    Dim dDist As Double
    Dim dDist_Min As Double

    Dim oPnt_Mid As AutoMath.DPosition
    Dim oPnt_Root As AutoMath.DPosition
    
    Dim oVec_M As AutoMath.DVector
    Dim oVec_N As AutoMath.DVector
    Dim oVec_W As AutoMath.DVector
    Dim oVec_Pipe As AutoMath.DVector

    Dim oObj1 As Object
    Dim oRtePathFeat As IJRtePathFeat
    Dim oOutfitting As IJDObjectCollection
    
    ' Get collection of Outfitting items
    Set oOutfitting = oHoleTraceAE.GetParentOutfitting
    If oOutfitting Is Nothing Then
        Exit Sub
    ElseIf oOutfitting.Count < 1 Then
        Exit Sub
    End If
    
    dDist_Min = -1#
    Set oPnt_Mid = New AutoMath.DPosition
    Set oPnt_Root = New AutoMath.DPosition
    oPnt_Root.Set oMatrix.IndexValue(12), oMatrix.IndexValue(13), oMatrix.IndexValue(14)
    
    ' Loop thru each of the Pipe's Path Features
    ' Find closest Feature to the Fitting origin point
    Dim bSet As Boolean
    
    For Each oObj1 In oOutfitting
        If TypeOf oObj1 Is IJRtePathFeat Then
            Set oRtePathFeat = oObj1
            oRtePathFeat.GetStartLocation dX1, dY1, dZ1
            oRtePathFeat.GetEndLocation dX2, dY2, dZ2

            bSet = False
            oPnt_Mid.Set (dX2 + dX1) / 2#, (dY2 + dY1) / 2#, (dZ2 + dZ1) / 2#
            dDist = oPnt_Root.DistPt(oPnt_Mid)
            If dDist_Min < 0# Then
                dDist_Min = dDist
                Set oVec_Pipe = New AutoMath.DVector
                bSet = True
            ElseIf dDist < dDist_Min Then
                dDist_Min = dDist
                bSet = True
            End If
            
            If Abs(dX1 - dX2) < 0.00001 And Abs(dY1 - dY2) < 0.00001 And Abs(dZ1 - dZ2) < 0.00001 Then
               If TypeOf oObj1 Is IJRteEndPathFeat Then
                  ' For end path feature, start/end location may be same,use leg vector
                  Dim oLeg1 As IJRtePathLeg
                  Dim oLeg2 As IJRtePathLeg
                  Dim dX As Double, dY As Double, dZ As Double
                  
                  oRtePathFeat.GetLegs oLeg1, oLeg2
                  If Not oLeg1 Is Nothing Then
                     oLeg1.GetDirectionVector oRtePathFeat, dX, dY, dZ
                     dX1 = dX1 + dX
                     dY1 = dY1 + dY
                     dZ1 = dZ1 + dZ
                     dX2 = dX2 - dX
                     dY2 = dY2 - dY
                     dZ2 = dZ2 - dZ
                  End If
               End If
            End If
            
            If bSet = True Then
                oVec_Pipe.Set (dX2 - dX1), (dY2 - dY1), (dZ2 - dZ1)
            End If
        End If
    Next oObj1
            
    If dDist_Min < 0# Then
        Exit Sub
    End If
            
    oVec_Pipe.Length = 1#
    
    ' Verify Pipe Direction and Working Plane Normal are NOT Parallel
    Set oVec_W = New AutoMath.DVector
    oVec_W.Set oMatrix.IndexValue(8), oMatrix.IndexValue(9), oMatrix.IndexValue(10)
    dDot = oVec_Pipe.Dot(oVec_W)
    
    If Abs(dDot) > 0.99991 Then
        ' Pipe Direction IS Parallel to Working Plane
        Exit Sub
        
    ElseIf dDot >= 0.000001 Then
        ' Pipe Direction IS IN SAME direction as Working Plane
    
    ElseIf dDot <= -0.000001 Then
        ' Reverse Pipe Direction to be in same direction as Working Plane
        oVec_Pipe.Length = -1#
    
    Else
        ' Pipe Direction is Perpendicular to Working Plane
    End If
    
    ' Calculate Normal vector to both the Pipe Direction and Working Plane Normal
    Set oVec_N = oVec_Pipe.Cross(oVec_W)
    oVec_N.Length = 1#
            
    ' Calculate Perpendicular vector to the Normal above and the Pipe Direction
    Set oVec_M = oVec_N.Cross(oVec_Pipe)
    oVec_M.Length = 1#
            
    ' Set the Fitting Transformation Matrix so that it is Along the Pipe using:
    '   oVec_M
    '   oVec_N
    '   oVec_Pipe
    '   oPnt_Root
    oMatrix.LoadIdentity
    oMatrix.IndexValue(0) = oVec_M.x
    oMatrix.IndexValue(1) = oVec_M.y
    oMatrix.IndexValue(2) = oVec_M.z
    oMatrix.IndexValue(4) = oVec_N.x
    oMatrix.IndexValue(5) = oVec_N.y
    oMatrix.IndexValue(6) = oVec_N.z
    oMatrix.IndexValue(8) = oVec_Pipe.x
    oMatrix.IndexValue(9) = oVec_Pipe.y
    oMatrix.IndexValue(10) = oVec_Pipe.z
    oMatrix.IndexValue(12) = oPnt_Root.x
    oMatrix.IndexValue(13) = oPnt_Root.y
    oMatrix.IndexValue(14) = oPnt_Root.z
    
    Exit Sub
End Sub

 
