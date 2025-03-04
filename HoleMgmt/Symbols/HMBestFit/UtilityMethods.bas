Attribute VB_Name = "UtilityMethods"
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: Hole Management - HoleMgmt\Middle\Symbols\HMBestFit
'
'File: UtilityMethods.bas
'
'Abstract: ComputeTrace of BestFit symbol
'
'Author: sypark@ship.samsung.co.kr
'
'Revision:
'     05/22/01 - sypark@ship.samsung.co.kr - initial release
'
'Note: ReportAndRaiseUnanticipatedError method is located in
'      CommonApp/Client/Bin/GSCADUtilities.dll and may have to
'      be replaced
'
' Changes:
'   Date    By    Description
' 06/09/03  CRS  Created method CreatePipeTurnProjection.
' 09/04/03  CRS  Modified GetPipeFlangeDimensions() to use PathLeg rather
'                than PathRun when looking for flange diameter.  Also
'                moved some of the code into a separate function.
' 09/09/03  CRS  Added function GetPipeSleeveOD to obtain outside
'                diameter attribute of catalog part.
'  01/20/04 asri TR 53579. Search for TR 53579 in the file. Redundant function GetFeaturesCol on IJRtePathFeat
'                was used.Because of this unable to locate the flange.
'                Now using GetFeatures.
'
'*******************************************************************

Option Explicit

Private Const MODULE = "HMBestFit.UtilityMethods(UtilityMethods.bas)"
Private m_oErrors As New IMSErrorLog.JServerErrors

'For the searched joint type
Public Enum JointType
    Flange = 1
    SquareFlange = 2
    Sleeve = 3
    None = 4
End Enum

'*******************************************************************************************
'*******************************************************************************************
' this first set of routines are dealing with pipes
'*******************************************************************************************
'*******************************************************************************************

'********************************************************************
'Method: CreatePipeProjection
'
'Interface: public function
'
'Abstract: Get the start point of pipe and pass them to CreateProjection
'********************************************************************
Public Function CreatePipeProjection(oRtePathFeat As IJRtePathFeat, dDiameter As Double) As IngrGeom3D.IJProjection
    Const Method = "CreatePipeProjection"
    On Error GoTo ErrorHandler

    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double

    'Get the start and end point of pipe
    oRtePathFeat.GetStartLocation dStartX, dStartY, dStartZ
    oRtePathFeat.GetEndLocation dEndX, dEndY, dEndZ
           
    If Abs(dStartX - dEndX) < 0.00001 And _
       Abs(dStartY - dEndY) < 0.00001 And _
       Abs(dStartZ - dEndZ) < 0.00001 Then
       If TypeOf oRtePathFeat Is IJRteEndPathFeat Then
          Dim oLeg1 As IJRtePathLeg
          Dim oLeg2 As IJRtePathLeg
          Dim dX As Double, dY As Double, dZ As Double
          
          oRtePathFeat.GetLegs oLeg1, oLeg2
          If Not oLeg1 Is Nothing Then
             oLeg1.GetDirectionVector oRtePathFeat, dX, dY, dZ
             dStartX = dStartX + dX
             dStartY = dStartY + dY
             dStartZ = dStartZ + dZ
             dEndX = dEndX - dX
             dEndY = dEndY - dY
             dEndZ = dEndZ - dZ
          End If
       End If
    End If
    Set CreatePipeProjection = CreateProjection(dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ, dDiameter)
   
Cleanup:

    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: CreatePipeTurnProjection
'
'Interface: public function
'
'Abstract: Get the start point of pipe and pass them to CreateProjection
'
'Algorithm:
'  * Determine the diameter of the projection using the existing code in ComputePipeTrace.
'    This takes into consideration the insulation, hole clearance, watertight integrity and stuff like that.
'  * Compute the arc that defines the centerline of the turn feature.
'  * Find the point of intersection of the centerline of step 2 and the working plane.  Call this point A.
'  * Create vector (V0) defined by the line from the center of rotation of the arc and point A.
'  * Determine the vector representing the Working Plane normal (V1).
'  * The vector (V2) that defines the direction of the projection will be 90 degrees (pi / 2) minus the angle between V0 and V1.
'  * The length of the projection is arbitrary.  So lets make it 2x the arc length.  Determine the arc length.
'    Get the radius, r
'    Get the SweepAngle, a
'    Arc length = r x a
'  * Compute the start point of the projection.
'  * Compute the end point of the projection.
'  * Create the projection based on end points and circle diameter.
'
'********************************************************************
Public Function CreatePipeTurnProjection _
        (oRtePathFeat As IJRtePathFeat, _
        oWorkingPlane As IJPlane, _
        dDiameter As Double) _
        As IngrGeom3D.IJProjection
        
    Const Method = "CreatePipeTurnProjection"
    On Error GoTo ErrorHandler

    'Get the start and end point of pipe
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    oRtePathFeat.GetStartLocation dStartX, dStartY, dStartZ
    
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    oRtePathFeat.GetEndLocation dEndX, dEndY, dEndZ
           
    ' Compute the arc that defines the centerline of the turn feature.
    ' Get the center of rotation of the curve.
    Dim ptCenter As IJDPosition
    Set ptCenter = New AutoMath.DPosition
    Dim dRadius As Double

    Dim oTurnPathFeat As IJRteTurnPathFeat
    Set oTurnPathFeat = oRtePathFeat

    dRadius = oTurnPathFeat.GetCenterOfRotationAndRadius(ptCenter)

    Dim dCenterX As Double
    Dim dCenterY As Double
    Dim dCenterZ As Double
    dCenterX = ptCenter.x
    dCenterY = ptCenter.y
    dCenterZ = ptCenter.z

    Dim oArc3d As IJArc
    Dim oGeomFactory As GeometryFactory
    Set oGeomFactory = New GeometryFactory
    Set oArc3d = oGeomFactory.Arcs3d.CreateByCenterStartEnd _
            (Nothing, dCenterX, dCenterY, dCenterZ, _
            dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ)
    
'   The length of the projection is arbitrary.  So lets make it 2x the arc length.  Determine the arc length.
    Dim dSweepAngle As Double
    dSweepAngle = oArc3d.SweepAngle
    Dim dArcLength As Double
    dArcLength = dSweepAngle * dRadius

'   3) Find the point of intersection of the centerline of step 2 and the working plane.  Call this point A.
    Dim oIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Set oIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    Dim oIntersectionUnknown As IUnknown
    Dim oAgtorUnk As IUnknown
    Dim oNullObject As Object
    oIntersector.PlaceIntersectionObject _
            oNullObject, oWorkingPlane, oArc3d, _
            oAgtorUnk, oIntersectionUnknown
            
    If TypeOf oIntersectionUnknown Is IJPointsGraphBody Then
    
        Dim oPointsGraphBody As IJPointsGraphBody
        Set oPointsGraphBody = oIntersectionUnknown
        
        Dim oPointUtils As IJSGOPointsGraphUtilities
        Set oPointUtils = New SGOPointsGraphUtilities
 
        Dim oPointsColl As Collection
        Set oPointsColl = oPointUtils.GetPositionsFromPointsGraph(oPointsGraphBody)
        
        Dim oPoint As IJDPosition
        Set oPoint = oPointsColl.Item(1)
                    
    End If
    
'   4) Create vector (V0) defined by the line from the center of rotation of the arc and point A.
    Dim oVector0 As IJDVector
    Set oVector0 = ptCenter.Subtract(oPoint)

'   5) Determine the vector representing the Working Plane normal (V1).
    Dim oVector1 As IJDVector
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oWorkingPlane.GetNormal dNormalX, dNormalY, dNormalZ
    Set oVector1 = New DVector
    oVector1.x = dNormalX
    oVector1.y = dNormalY
    oVector1.z = dNormalZ

'   6) The vector (V2) that defines the direction of the projection will be 90 degrees (pi / 2) minus the angle between V0 and V1.
'      But this is probably better.
    Dim oVector2 As IJDVector
    Set oVector2 = oVector0.Cross(oVector1)
    Set oVector2 = oVector0.Cross(oVector2)
    oVector2.Length = 1

'   7) Compute the start point of the projection.
    dStartX = oPoint.x + oVector2.x * dArcLength
    dStartY = oPoint.y + oVector2.y * dArcLength
    dStartZ = oPoint.z + oVector2.z * dArcLength

'   8) Compute the end point of the projection.
    dEndX = oPoint.x - oVector2.x * dArcLength
    dEndY = oPoint.y - oVector2.y * dArcLength
    dEndZ = oPoint.z - oVector2.z * dArcLength

'   9) Create the projection based on end points and circle diameter.
    Set CreatePipeTurnProjection = CreateProjection _
            (dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ, dDiameter)

    GoTo Cleanup

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    
Cleanup:
    Set ptCenter = Nothing
    Set oTurnPathFeat = Nothing
    Set oGeomFactory = Nothing
    Set oArc3d = Nothing
    Set oIntersector = Nothing
    Set oIntersectionUnknown = Nothing
    Set oAgtorUnk = Nothing
    Set oNullObject = Nothing
    Set oPointsGraphBody = Nothing
    Set oPointUtils = Nothing
    Set oPointsColl = Nothing
    Set oPoint = Nothing
    Set oVector0 = Nothing
    Set oVector1 = Nothing
    Set oVector2 = Nothing
    
End Function  ' CreatePipeTurnProjection

'********************************************************************
'Method: GetPipeFlangeDimensions
'
'Interface: public function
'
'Abstract: Get the diameter of Pipe Component ( Flange & Sleeve )
'********************************************************************
Public Function GetPipeFlangeDimensions _
        (oRtePathFeat As IJRtePathFeat, ByRef eJointType As JointType) As Double
    
    Const Method = "GetPipeFlangeDimensions"
    On Error GoTo ErrorHandler
     
    ' Initialize the return variables.
    GetPipeFlangeDimensions = 0#
    eJointType = None
    
    ' Most likely there is one leg associated with the feature.  There
    ' could be two in the case of a turn feature.
    Dim oLeg1 As IJRtePathLeg
    Dim oLeg2 As IJRtePathLeg
    Dim nLegCount As Long
    nLegCount = oRtePathFeat.GetLegs(oLeg1, oLeg2)

    Dim iLeg As Integer
    For iLeg = 1 To nLegCount
        Dim oPathCollection As IJDTargetObjectCol
        Dim nFeatureCount As Long
        Dim iFeatureIndex As Long
        
        ' Get the fetures on the leg.
'        TR 53579.Using GetFeatures instead of the old GetFeaturesCol
        If iLeg = 1 Then
            oLeg1.GetFeatures oPathCollection, nFeatureCount
        Else
            oLeg2.GetFeatures oPathCollection, nFeatureCount
        End If
        
        Dim iCount As Integer   'The count of components which the PathLeg has.
        For iCount = 1 To oPathCollection.Count
        
            Dim oObject As Object
            Set oObject = oPathCollection.Item(iCount)
            If oObject Is Nothing Then GoTo NextItem
            
            Dim oRteFeat As IJRtePathFeat
            Set oRteFeat = oObject
            If oRteFeat Is Nothing Then GoTo NextItem
            
            'Get the flange or Hub Outside dia.
            Dim dFlangeDimension As Double
            Dim eType As JointType
            dFlangeDimension = GetPipeFeatFlangeDimensions _
                    (oRteFeat, eType)
        
            If dFlangeDimension > 0 And _
                    GetPipeFlangeDimensions < dFlangeDimension Then
                GetPipeFlangeDimensions = dFlangeDimension
                eJointType = eType
            End If
    
NextItem:
            Set oObject = Nothing
            Set oRteFeat = Nothing
        Next iCount
        
        Set oPathCollection = Nothing
    Next iLeg
    
    GoTo Cleanup
  
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description

Cleanup:
    Set oLeg1 = Nothing
    Set oLeg2 = Nothing
    Set oPathCollection = Nothing
    Set oObject = Nothing
    Set oRteFeat = Nothing
    
End Function ' GetPipeFlangeDimensions

'********************************************************************
'Method: GetPipeFeatFlangeDimensions
'
'Interface: private function
'
'Abstract: Get the diameter of Pipe Component ( Flange & Sleeve )
'********************************************************************
Private Function GetPipeFeatFlangeDimensions _
        (oRteFeat As IJRtePathFeat, _
        ByRef eJointType As JointType) As Double
        
    Const Method = "GetPipeFeatFlangeDimensions"
    On Error GoTo ErrorHandler
    
    GetPipeFeatFlangeDimensions = 0#
     
    'Set the joint type to none.
    eJointType = None
    
        Dim eFunction As PathFeatureFunctions
        Dim eObjectType As PathFeatureObjectTypes
        oRteFeat.GetTypeAndFunction eFunction, eObjectType
        
        'Check if the component is end feature or component
        If eObjectType = PathFeatureType_ALONG Or _
            eFunction = PathFeatureFunction_COMPONENT Or _
            eObjectType = PathFeatureType_END Then
          
            'Get the base part from Feature.
            Dim oPartElement As IJElements
            Dim lPartCount As Long
            Set oPartElement = oRteFeat.GetBaseParts(lPartCount)
            If oPartElement Is Nothing Then GoTo Cleanup
            
            Dim oPartObject As Variant
            Dim oPartOcc As IJPartOcc
            Dim oPart As IJDPart
            Dim oPartCollection As IJDCollection
            Dim oNozzle As IJCatalogPipePort
            
            Dim iCount As Integer
            
            'Get the nozzle part from BasePart.
            For Each oPartObject In oPartElement
                Set oPartOcc = oPartObject
                If oPartOcc Is Nothing Then GoTo Cleanup
                
                oPartOcc.GetPart oPart
                If oPart Is Nothing Then GoTo Cleanup

                Set oPartCollection = oPart.GetNozzles
                
                If Not oPartCollection Is Nothing Then
                    For iCount = 1 To oPartCollection.Size
                        Set oNozzle = oPartCollection.Item(iCount)
    
                        'Get the flange or Hub Outside dia.
                        Dim dFlangeDimension As Double
                        dFlangeDimension = oNozzle.FlangeOrHubOutsideDiameter
        
                        If dFlangeDimension > 0 And _
                                GetPipeFeatFlangeDimensions < dFlangeDimension Then
                            GetPipeFeatFlangeDimensions = dFlangeDimension
                            eJointType = GetPipeJointType(oPart)
                        End If
                        
                    Next iCount
                End If
            Next oPartObject
        End If
    
    GoTo Cleanup
  
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    
Cleanup:
    Set oPartElement = Nothing
    Set oPartOcc = Nothing
    Set oPart = Nothing
    Set oPartCollection = Nothing
    Set oPartObject = Nothing
    Set oNozzle = Nothing
    
End Function ' GetPipeFeatFlangeDimensions

'********************************************************************
'Method: GetPipeDimensions
'
'Interface: public function
'
'Abstract: Get the diameter of Pipe
'********************************************************************
Public Function GetPipeDimensions(oRtePathFeat As IJRtePathFeat) As Double
    Const Method = "GetPipeDimensions"
    On Error GoTo ErrorHandler
    
    'If one pipe was taken, There is two ports on the start and end of pipe.
    'If the size of two ports of pipe is different, The maximun size should be taken
    Dim dMaxPortDiaSize  As Double
    Dim dCurretPortDiaSize As Double
                          
    dMaxPortDiaSize = 0#
    dCurretPortDiaSize = 0#
                          
    Dim oPartElement As IJElements
    Dim lPartCount As Long
    Set oPartElement = oRtePathFeat.GetBaseParts(lPartCount)
    
    If Not oPartElement Is Nothing Then
        Dim oPartObject As Object
        Dim oGenPart As IJRtePathGenPart
        Dim oPartOcc As IJDistribPartOccur
        Dim oPipePort As IJDPipePort
        
        For Each oPartObject In oPartElement
            Set oGenPart = oPartObject
            Set oPartOcc = oGenPart
            Dim oPortElement As IJElements
            Set oPortElement = oPartOcc.GetPorts(DistribPortStatus_BASE, DistribPortType_PIPE)
            
            Dim iCount As Integer
            For iCount = 1 To oPortElement.Count
                Set oPipePort = oPortElement.Item(iCount)
                dCurretPortDiaSize = oPipePort.PipingOutsideDiameter
                
                If dCurretPortDiaSize > dMaxPortDiaSize Then
                    dMaxPortDiaSize = dCurretPortDiaSize
                End If
            Next iCount
        Next oPartObject
    Else
        'try getting it from the run
        Dim oPathRun As IJRtePathRun
        Dim oPipeRun As IJRtePipeRun
        
        Set oPathRun = oRtePathFeat.GetPathRun
        Set oPipeRun = oPathRun
        If Not oPipeRun Is Nothing Then
            Dim strType As String
            dMaxPortDiaSize = oPipeRun.GetNominalDiameter(strType)
            dMaxPortDiaSize = dMaxPortDiaSize / 1000#
        End If
        Set oPathRun = Nothing
        Set oPipeRun = Nothing
    End If
    
    GetPipeDimensions = dMaxPortDiaSize

Cleanup:
    Set oPipePort = Nothing
    Set oPortElement = Nothing
    Set oPartOcc = Nothing
    Set oGenPart = Nothing
    Set oPartObject = Nothing
    Set oPartElement = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetPipeJointType
'
'Interface: public function
'
'Abstract: GetPipeJointType
'********************************************************************
Public Function GetPipeJointType(oPart As IJDPart) As JointType
    Const Method = "GetPipeJointType"
    On Error GoTo ErrorHandler
                         
    Dim oPipeComp As IJDPipeComponent
    Set oPipeComp = oPart
    
    If oPipeComp.CommodityType = 1000 Then
        GetPipeJointType = Sleeve
    
    ''Following code is up to each ship company rule
    ''Check the commoditytype of SquareFlange
    ''    ElseIf oPipeComp.CommodityType = 1000 Then
    ''        GetJointType = SquareFlange
    Else
       GetPipeJointType = Flange
    End If
    
Cleanup:
    Set oPipeComp = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: PipeSimplification
'
'Interface: public function
'
'Abstract: PipeSimplification
'          If the pipes are less then 3, We need to check the simplification rule.
'          1. Get the flange dia of main pipe and compare the radius from another pipe
'          2. If the difference between main radius and compared raius is less then value
'             which is defind by user.
'          3. Returns the compared diameter.
'********************************************************************
Public Function PipeSimplification(dFlangeDia As Double, dPipeHoleMinDist As Double, _
                                   oMainRteObject As IJRtePathFeat, _
                                   oOutFittingCollection As IMSCoreCollections.IJDObjectCollection) As Double
    Const Method = "PipeSimplification "
    On Error GoTo ErrorHandler
    
    Dim oObj As Object
    Dim oRtePathFeat As IJRtePathFeat
    
    Dim dComparedFlangeDia As Double
    Dim eJointType As JointType
    
    PipeSimplification = 0
    
    For Each oObj In oOutFittingCollection
        If TypeOf oObj Is IJRtePipePathFeat Then
            Set oRtePathFeat = oObj
      
            If Not oRtePathFeat Is oMainRteObject Then
                dComparedFlangeDia = GetPipeFlangeDimensions(oRtePathFeat, eJointType)
    
                If dComparedFlangeDia <= 0 Then
                    dComparedFlangeDia = GetPipeDimensions(oRtePathFeat)
                End If
            
                If dComparedFlangeDia > dFlangeDia Then
                    If (dComparedFlangeDia - dFlangeDia) <= dPipeHoleMinDist Then
                        PipeSimplification = dComparedFlangeDia
                        Exit For
                    End If
                End If
            End If
        End If
    Next oObj
    
    'This is for not returnning zero when there is no checked hole.
    If PipeSimplification <= 0 Then
        PipeSimplification = dFlangeDia
    End If
        
Cleanup:
    Set oObj = Nothing
    Set oRtePathFeat = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'*******************************************************************************************
'*******************************************************************************************
' this next set of routines are dealing with ducts
'*******************************************************************************************
'*******************************************************************************************

'********************************************************************
'Method: CreateDuctProjection
'
'Interface: public function
'
'Abstract: Get the four point from start and end point of duct and pass them to CreateProjection
'********************************************************************
Public Function CreateDuctProjection(oRteFeature As IJRtePathFeat, dWidth As Double, dDepth As Double, _
                                     dDuctCorRadius As Double) As IMSCoreCollections.IJDObjectCollection
    Const Method = "CreateDuctProjection "
    On Error GoTo ErrorHandler
    
    Dim oTempProjection As IJProjection
    
    Dim dDiameter As Double
    Dim dDiaWithClearance As Double
    
    Dim dRteStartX As Double, dRteStartY As Double, dRteStartZ As Double
    Dim dRteEndX As Double, dRteEndY As Double, dRteEndZ As Double
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    
    Dim oTempDuctCollection As New IMSCoreCollections.JObjectCollection
    
    oRteFeature.GetStartLocation dRteStartX, dRteStartY, dRteStartZ
    oRteFeature.GetEndLocation dRteEndX, dRteEndY, dRteEndZ
    
    
    Dim dUVectorX As Double, dUVectorY As Double, dUVectorZ As Double
    Dim dVVectorX As Double, dVVectorY As Double, dVVectorZ As Double

    Dim oRteFeatUtility As IJRtePathCrossSectUtility
    
    Set oRteFeatUtility = oRteFeature
    oRteFeatUtility.GetWidthAndDepthAxis Nothing, dUVectorX, dUVectorY, dUVectorZ, _
                                        dVVectorX, dVVectorY, dVVectorZ
    
    
      
    dDiameter = dDuctCorRadius * 2
    
    Dim dHalfWidth As Double
    Dim dHalfDepth As Double
    dHalfWidth = dWidth / 2 - dDuctCorRadius
    dHalfDepth = dDepth / 2 - dDuctCorRadius
    
    ' Coordinates of the center of the projection that defines each
    ' corner after rotation is taken into consideration.
    Dim x0 As Double, y0 As Double
    
    ' Coordinates of the center of the projection that defines each
    ' corner before any rotation is taken into consideration.
    Dim x1 As Double, y1 As Double
    
     'We need to consider each of the four corners.
    Dim i As Integer
    For i = 0 To 3
        ' Each corner is located by its position from the center.
        ' Therefore define the quadrent that we are in (+/+, +/-, -/+, -/-).
        If i < 2 Then
            x1 = dHalfWidth
        Else
            x1 = -dHalfWidth
        End If
        
        If i Mod 2 = 0 Then
            y1 = dHalfDepth
        Else
            y1 = -dHalfDepth
        End If
        
      ' Determine the end-points of the projection to be created.
        dStartX = dRteStartX + x1 * dUVectorX + y1 * dVVectorX
        dStartY = dRteStartY + x1 * dUVectorY + y1 * dVVectorY
        dStartZ = dRteStartZ + x1 * dUVectorZ + y1 * dVVectorZ
        
        dEndX = dRteEndX + x1 * dUVectorX + y1 * dVVectorX
        dEndY = dRteEndY + x1 * dUVectorY + y1 * dVVectorY
        dEndZ = dRteEndZ + x1 * dUVectorZ + y1 * dVVectorZ
        
        ' Create the projection and add it to a collection of projections.
        Set oTempProjection = CreateProjection(dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ, dDiameter)
        oTempDuctCollection.Add oTempProjection
        Set oTempProjection = Nothing
    Next i  ' next corner.
    
    Set CreateDuctProjection = oTempDuctCollection
       
Cleanup:
    Set oTempDuctCollection = Nothing
    Set oTempProjection = Nothing
    Set oRteFeatUtility = Nothing
    
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: CreateFlatOvalProjection
'
'Interface: public function
'
'Abstract: Get the four point from start and end point of duct and pass them to CreateProjection
'********************************************************************
Public Function CreateFlatOvalProjection(oRteFeature As IJRtePathFeat, dWidth As Double, dDepth As Double) As IMSCoreCollections.IJDObjectCollection
    Const Method = "CreateFlatOvalProjection "
    On Error GoTo ErrorHandler
    
    Dim oProjection As IJProjection
    Dim oDuctCollection As New IMSCoreCollections.JObjectCollection
    
    Dim dRteStartX As Double, dRteStartY As Double, dRteStartZ As Double
    Dim dRteEndX As Double, dRteEndY As Double, dRteEndZ As Double
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    
    oRteFeature.GetStartLocation dRteStartX, dRteStartY, dRteStartZ
    oRteFeature.GetEndLocation dRteEndX, dRteEndY, dRteEndZ
    
    Dim oRteFeatUtility As IJRtePathCrossSectUtility
     
    Dim dUVectorX As Double, dUVectorY As Double, dUVectorZ As Double
    Dim dVVectorX As Double, dVVectorY As Double, dVVectorZ As Double

    Set oRteFeatUtility = oRteFeature
    If Not oRteFeatUtility Is Nothing Then
        oRteFeatUtility.GetWidthAndDepthAxis Nothing, dUVectorX, dUVectorY, dUVectorZ, _
                                         dVVectorX, dVVectorY, dVVectorZ
    End If

    ' Coordinates of the center of the projection that defines each
    ' corner before any rotation is taken into consideration.
    Dim x1 As Double, y1 As Double
    
      
    x1 = (dWidth - dDepth) / 2
    y1 = 0
    
    dStartX = dRteStartX + x1 * dUVectorX + y1 * dVVectorX
    dStartY = dRteStartY + x1 * dUVectorY + y1 * dVVectorY
    dStartZ = dRteStartZ + x1 * dUVectorZ + y1 * dVVectorZ
    
    dEndX = dRteEndX + x1 * dUVectorX + y1 * dVVectorX
    dEndY = dRteEndY + x1 * dUVectorY + y1 * dVVectorY
    dEndZ = dRteEndZ + x1 * dUVectorZ + y1 * dVVectorZ
        
    Set oProjection = CreateProjection(dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ, dDepth)
    
    oDuctCollection.Add oProjection
    Set oProjection = Nothing
    
    dStartX = dRteStartX - x1 * dUVectorX - y1 * dVVectorX
    dStartY = dRteStartY - x1 * dUVectorY - y1 * dVVectorY
    dStartZ = dRteStartZ - x1 * dUVectorZ - y1 * dVVectorZ
    
    dEndX = dRteEndX - x1 * dUVectorX - y1 * dVVectorX
    dEndY = dRteEndY - x1 * dUVectorY - y1 * dVVectorY
    dEndZ = dRteEndZ - x1 * dUVectorZ - y1 * dVVectorZ
    Set oProjection = CreateProjection(dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ, dDepth)
    
    oDuctCollection.Add oProjection
    Set oProjection = Nothing
       
    Set CreateFlatOvalProjection = oDuctCollection
       
Cleanup:
    Set oDuctCollection = Nothing
    Set oProjection = Nothing
    Set oRteFeatUtility = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetDuctDimensions
'
'Interface: public function
'
'Abstract: Get the Width and Depth of Duct
'********************************************************************
Public Function GetDuctDimensions(oRtePathFeat As IJRtePathFeat, _
                                  ByRef dWidth As Double, ByRef dDepth As Double, _
                                  ByRef dRadius As Double) As CrossSectionShapeTypes
    Const Method = "GetDuctDimensions"
    On Error GoTo ErrorHandler
    
    Dim oRteFeatUtility As IJRtePathCrossSectUtility
    Dim bOuterDia As Boolean
    
    Set oRteFeatUtility = oRtePathFeat
    If oRteFeatUtility Is Nothing Then
        MsgBox "COULD NOT GET oRteFeatUtility FROM DUCT FEATURE"
        GoTo Cleanup
    End If
    
    oRteFeatUtility.GetCrossSectionData Nothing, False, GetDuctDimensions, dWidth, dDepth, dRadius, bOuterDia
    
    
Cleanup:
    Set oRteFeatUtility = Nothing

    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetDuctFlangeDimensions
'
'Interface: public function
'
'Abstract: Get the dimensions of Duct Component ( Flange & Sleeve )
'********************************************************************
Public Function GetDuctFlangeDimensions(oRtePathFeat As IJRtePathFeat, ByRef dWidth As Double, _
                                            ByRef dDepth As Double, eJointType As JointType) As CrossSectionShapeTypes
    
    Const Method = "GetDuctFlangeDimensions"
    On Error GoTo ErrorHandler
     
    ' Initialize the return variables.
    eJointType = None
    
    ' Most likely there is one leg associated with the feature.  There
    ' could be two in the case of a turn feature.
    Dim oLeg1 As IJRtePathLeg
    Dim oLeg2 As IJRtePathLeg
    Dim nLegCount As Long
    nLegCount = oRtePathFeat.GetLegs(oLeg1, oLeg2)

    Dim iLeg As Integer
    For iLeg = 1 To nLegCount
        Dim oPathCollection As IJDTargetObjectCol
        Dim nFeatureCount As Long
        Dim iFeatureIndex As Long
        
        ' Get the fetures on the leg.
'        TR 53579.Using GetFeatures instead of the old GetFeaturesCol
        If iLeg = 1 Then
            oLeg1.GetFeatures oPathCollection, nFeatureCount
        Else
            oLeg2.GetFeatures oPathCollection, nFeatureCount
        End If
        
        'The count of components which the PathLeg has.
        Dim eCSShapeTypes As CrossSectionShapeTypes
        Dim iCount As Integer
        For iCount = 1 To oPathCollection.Count
        
            Dim oObject As Object
            Set oObject = oPathCollection.Item(iCount)
            If oObject Is Nothing Then GoTo NextItem
            
            Dim oRteFeat As IJRtePathFeat
            Set oRteFeat = oObject
            If oRteFeat Is Nothing Then GoTo NextItem
            
            'Get the flange or Hub Outside dia.
            Dim dFlangeWidth As Double
            Dim dFlangeDepth As Double
            Dim eType As JointType
            Dim eShapeTypes As CrossSectionShapeTypes
            eShapeTypes = GetDuctFeatFlangeDimensions _
                                        (oRteFeat, dFlangeWidth, dFlangeDepth, eType)
        
            If (dFlangeWidth > 0 And dFlangeDepth > 0) And _
                    (dWidth < dFlangeWidth Or dDepth < dFlangeDepth) Then
                dWidth = dFlangeWidth
                dDepth = dFlangeDepth
                eJointType = eType
                eCSShapeTypes = eShapeTypes
            End If
    
NextItem:
            Set oObject = Nothing
            Set oRteFeat = Nothing
        Next iCount
        
        Set oPathCollection = Nothing
    Next iLeg
    GetDuctFlangeDimensions = eCSShapeTypes
    GoTo Cleanup
  
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description

Cleanup:
    Set oLeg1 = Nothing
    Set oLeg2 = Nothing
    Set oPathCollection = Nothing
    Set oObject = Nothing
    Set oRteFeat = Nothing
    
End Function ' GetDuctFlangeDimensions

'********************************************************************
'Method: GetDuctFeatFlangeDimensions
'
'Interface: private function
'
'Abstract: Get the Dimensions of Duct Component ( Flange )
'********************************************************************
Private Function GetDuctFeatFlangeDimensions(oRtePathFeat As IJRtePathFeat, _
                                        ByRef dWidth As Double, ByRef dDepth As Double, _
                                        eJointType As JointType) As CrossSectionShapeTypes
    Const Method = "GetDuctFeatFlangeDimensions"
    On Error GoTo ErrorHandler

    Dim oPathCollection As IJDTargetObjectCol
    Dim iCount As Integer
    Dim eFunction As PathFeatureFunctions
    Dim eObjectType As PathFeatureObjectTypes
    
    'Initialize joint type as None
    eJointType = None
    oRtePathFeat.GetTypeAndFunction eFunction, eObjectType
    
    If eObjectType = PathFeatureType_ALONG Or eFunction = PathFeatureFunction_SPLIT Or eFunction = PathFeatureFunction_COMPONENT Then
        Dim oPartElement As IJElements
        Dim oGenPart As IJRtePathGenPart
        Dim lPartCount As Long
        Dim oDistribPartOcc As IJDistribPartOccur
        Dim oPartOcc As IJPartOcc
        Dim oPart As IJDPart
        Dim oPartCollection As IJDCollection
        Dim oHvacPort As IJDHvacPort
        Dim oPartObject As Object
        Dim iiCount As Integer
        
        Set oPartElement = oRtePathFeat.GetBaseParts(lPartCount)
        If oPartElement Is Nothing Then GoTo Cleanup
        
        For Each oPartObject In oPartElement
            Dim oPortElement As IJElements
           
            Set oGenPart = oPartObject
            Set oDistribPartOcc = oGenPart
            Set oPortElement = oDistribPartOcc.GetPorts(DistribPortStatus_BASE, DistribPortType_DUCT)
    
            For iiCount = 1 To oPortElement.Count
                Set oHvacPort = oPortElement.Item(iiCount)
                
                If oHvacPort.EndPrep < 300 Then
                    Dim oPortCrossSection As IJDOutfittingCrossSection
                    Set oPortCrossSection = oHvacPort.GetCrossSection
    
                    GetDuctFeatFlangeDimensions = oPortCrossSection.GetShape
                        
                    dWidth = oPortCrossSection.Width + 2 * oHvacPort.FlangeWidth 'Flange width = Width of CS  + 2 *  Flange Width
                    dDepth = oPortCrossSection.Depth + 2 * oHvacPort.FlangeWidth  'Flange Depth = Depth of CS  + 2 *  Flange Width
                        
                    Set oPartOcc = oPartObject
                    oPartOcc.GetPart oPart
                    eJointType = GetDuctJointType(oPart)
                    Set oPortCrossSection = Nothing
                End If
            Next iiCount
            
            Set oGenPart = Nothing
            Set oPortElement = Nothing
        Next oPartObject
    End If
    
Cleanup:
    Set oPartElement = Nothing
    Set oDistribPartOcc = Nothing
    Set oPart = Nothing
    Set oPartCollection = Nothing
    Set oPartObject = Nothing
    
    Exit Function
  
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function 'GetDuctFeatFlangeDimensions

'********************************************************************
'Method: GetDuctJointType
'
'Interface: public function
'
'Abstract: GetDuctJointType
'********************************************************************
Public Function GetDuctJointType(oPart As IJDPart) As JointType
    Const Method = "GetDuctJointType"
    On Error GoTo ErrorHandler
    
    Dim strCatalogName As String
    strCatalogName = oPart.GetRelatedPartClassName

    If strCatalogName = "Rect_FlatFlange" Or strCatalogName = "Rect_Sleeve" Then
        GetDuctJointType = SquareFlange
    ElseIf strCatalogName = "Round_Sleeve" Then
       GetDuctJointType = Sleeve
    Else
       GetDuctJointType = Flange
    End If
    
Cleanup:
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'*******************************************************************************************
'*******************************************************************************************
' this next set of routines are dealing with cableways
'*******************************************************************************************
'*******************************************************************************************

'********************************************************************
'Method: GetCablewayDimensions
'
'Interface: public function
'
'Abstract: Get the Width and Depth of cableway
'********************************************************************
Public Function GetCablewayDimensions(oRtePathFeat As IJRtePathFeat, _
                                      ByRef dWidth As Double, ByRef dDepth As Double) As CrossSectionShapeTypes
    Const Method = "GetCablewayDimensions"
    On Error GoTo ErrorHandler
        
    Dim oRteFeatUtility As IJRtePathCrossSectUtility
    Dim bOuterDia As Boolean
    Dim dRadius As Double
    
    Set oRteFeatUtility = oRtePathFeat
    If oRteFeatUtility Is Nothing Then
        MsgBox "COULD NOT GET oRteFeatUtility FROM CABLEWAY FEATURE"
        GoTo Cleanup
    End If
    
    oRteFeatUtility.GetCrossSectionData Nothing, False, GetCablewayDimensions, dWidth, dDepth, dRadius, bOuterDia
    
Cleanup:
    Set oRteFeatUtility = Nothing

    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetCablewayFlangeDimensions
'
'Interface: public function
'
'Abstract: Get the Width of cableway Component ( Flange )
'********************************************************************
Public Function GetCablewayFlangeDimensions(oRtePathFeat As IJRtePathFeat, _
                                            ByRef dWidth As Double, ByRef dDepth As Double, _
                                            eJointType As JointType) As CrossSectionShapeTypes
    Const Method = "GetCablewayFlangeDimensions"
    On Error GoTo ErrorHandler
     
    Dim oPathRun As IJRtePathRun
    Dim oPathCollection As IJDTargetObjectCol
    Dim oObject As Object
    
    Dim iCount As Integer
    
    Set oPathRun = oRtePathFeat.GetPathRun
    Set oPathCollection = oPathRun.GetPathFeatures
     
    'Initialize joint type as None
    eJointType = None
    
    For iCount = 1 To oPathCollection.Count
        Set oObject = oPathCollection.Item(iCount)
        
        If oObject Is Nothing Then GoTo JumpToNextStep
        
        Dim eFunction As PathFeatureFunctions
        Dim eObjectType As PathFeatureObjectTypes
        Dim oRteFeat As IJRtePathFeat
        
        Set oRteFeat = oObject
        If oRteFeat Is Nothing Then GoTo JumpToNextStep
        
        oRteFeat.GetTypeAndFunction eFunction, eObjectType
    
        If eObjectType = PathFeatureType_ALONG Or eFunction = PathFeatureFunction_COMPONENT Then
            Dim oPartElement As IJElements
            Dim oGenPart As IJRtePathGenPart
            Dim lPartCount As Long
            
            Set oPartElement = oRteFeat.GetBaseParts(lPartCount)
            If oPartElement Is Nothing Then GoTo JumpToNextStep
                                
            Dim oDistribPartOcc As IJDistribPartOccur
            Dim oPartOcc As IJPartOcc
            Dim oPart As IJDPart
            Dim oPartCollection As IJDCollection
            Dim oHvacPort As IJDHvacPort
            Dim oPartObject As Object
            
            Dim iiCount As Integer
            
            For Each oPartObject In oPartElement
                Dim oPortElement As IJElements
               
                Set oGenPart = oPartObject
                Set oDistribPartOcc = oGenPart
                Set oPortElement = oDistribPartOcc.GetPorts(DistribPortStatus_BASE, DistribPortType_CABLETRAY)
 
                For iiCount = 1 To oPortElement.Count
                    Set oHvacPort = oPortElement.Item(iiCount)
                    
                    If oHvacPort.EndPrep < 300 Then
                        Dim oPortCrossSection As IJDOutfittingCrossSection
                        Set oPortCrossSection = oHvacPort.GetCrossSection
    
                        GetCablewayFlangeDimensions = oPortCrossSection.GetShape
                            
                        dWidth = oPortCrossSection.Width
                        dDepth = oPortCrossSection.Depth
                            
                        Set oPartOcc = oPartObject
                        oPartOcc.GetPart oPart
                        eJointType = GetCablewayJointType(oPart)
                        Set oPortCrossSection = Nothing
                    End If
                Next iiCount
                
                Set oGenPart = Nothing
                Set oPortElement = Nothing
            Next oPartObject
        End If
JumpToNextStep:
    Next iCount
  
Cleanup:
    Set oPathRun = Nothing
    Set oPathCollection = Nothing
    Set oObject = Nothing
    Set oRteFeat = Nothing
    Set oPartElement = Nothing
    Set oDistribPartOcc = Nothing
    Set oPart = Nothing
    Set oPartCollection = Nothing
    Set oPartObject = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetCablewayJointType
'
'Interface: public function
'
'Abstract: GetCablewayJointType
'********************************************************************
Public Function GetCablewayJointType(oPart As IJDPart) As JointType
    Const Method = "GetCablewayJointType"
    On Error GoTo ErrorHandler
    
    Dim strCatalogName As String
    strCatalogName = oPart.GetRelatedPartClassName

    If strCatalogName = "Rect_FlatFlange" Or strCatalogName = "Rect_Sleeve" Then
        GetCablewayJointType = SquareFlange
    ElseIf strCatalogName = "Round_Sleeve" Then
       GetCablewayJointType = Sleeve
    Else
       GetCablewayJointType = Flange
    End If
    
Cleanup:
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'*******************************************************************************************
'*******************************************************************************************
' this next set of routines are dealing with conduits
'*******************************************************************************************
'*******************************************************************************************

'********************************************************************
'Method: GetConduitFlangeDimensions
'
'Interface: public function
'
'Abstract: Get the diameter of Pipe Component ( Flange & Sleeve )
'********************************************************************
Public Function GetConduitFlangeDimensions(oRtePathFeat As IJRtePathFeat, ByRef eJointType As JointType) As Double
    Const Method = "GetConduitFlangeDimensions"
    On Error GoTo ErrorHandler
     
    Dim oPathRun As IJRtePathRun
    Dim oPathCollection As IJDTargetObjectCol
    Dim oObject As Object
    
    Dim dTempFlangeDimension As Double
    
    Dim iCount As Integer   'The count of components which the PathRun has.
    
    Set oPathRun = oRtePathFeat.GetPathRun
    Set oPathCollection = oPathRun.GetPathFeatures
     
    GetConduitFlangeDimensions = 0#
    dTempFlangeDimension = 0#
    
    'Set the none
    eJointType = None
    
    For iCount = 1 To oPathCollection.Count
        Set oObject = oPathCollection.Item(iCount)
        
        If oObject Is Nothing Then GoTo JumpToNextStep
        
        Dim eFunction As PathFeatureFunctions
        Dim eObjectType As PathFeatureObjectTypes
        Dim oRteFeat As IJRtePathFeat
        
        Set oRteFeat = oObject
        If oRteFeat Is Nothing Then GoTo JumpToNextStep
        
        oRteFeat.GetTypeAndFunction eFunction, eObjectType
        
        'Check if the component is end feature or component
        If eObjectType = PathFeatureType_ALONG Or _
            eFunction = PathFeatureFunction_COMPONENT Or _
            eObjectType = PathFeatureType_END Then
          
            'Get the base part from Feature.
            Dim oPartElement As IJElements
            Dim lPartCount As Long
            Set oPartElement = oRteFeat.GetBaseParts(lPartCount)
            If oPartElement Is Nothing Then GoTo JumpToNextStep
            
            Dim oPartObject As Variant
            Dim oPartOcc As IJPartOcc
            Dim oPart As IJDPart
            Dim oPartCollection As IJDCollection
            Dim oNozzle As IJCatalogPipePort
            
            Dim iiCount As Integer
            
            'Get the nozzle part from BasePart.
            For Each oPartObject In oPartElement
                Set oPartOcc = oPartObject
                If oPartOcc Is Nothing Then GoTo JumpToNextStep
                
                oPartOcc.GetPart oPart
                If oPart Is Nothing Then GoTo JumpToNextStep

                Set oPartCollection = oPart.GetNozzles
                
                For iiCount = 1 To oPartCollection.Size
                    Set oNozzle = oPartCollection.Item(iiCount)

                    'Get the flange or Hub Outside dia.
                    If oNozzle.FlangeOrHubOutsideDiameter > 0 Then
                        dTempFlangeDimension = oNozzle.FlangeOrHubOutsideDiameter
                       
                        'Return the flange diameter and joint type
                        If GetConduitFlangeDimensions < dTempFlangeDimension Then
                            GetConduitFlangeDimensions = dTempFlangeDimension
                            eJointType = GetConduitJointType(oPart)
                        End If
                    End If
                Next iiCount
            Next oPartObject
        End If
JumpToNextStep:
    Next iCount
    
Cleanup:
    Set oPathRun = Nothing
    Set oPathCollection = Nothing
    Set oObject = Nothing
    Set oRteFeat = Nothing
    Set oPartElement = Nothing
    Set oPartOcc = Nothing
    Set oPart = Nothing
    Set oPartCollection = Nothing
    Set oPartObject = Nothing
    Set oNozzle = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetConduitDimensions
'
'Interface: public function
'
'Abstract: Get the diameter of conduit
'********************************************************************
Public Function GetConduitDimensions(oRtePathFeat As IJRtePathFeat) As Double
    Const Method = "GetConduitDimensions"
    On Error GoTo ErrorHandler
    
    Dim oConduitPathFeat As IJRteConduitPathFeat
    Set oConduitPathFeat = oRtePathFeat
    
    GetConduitDimensions = oConduitPathFeat.OuterDiameter
    
Cleanup:
    Set oConduitPathFeat = Nothing
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetConduitJointType
'
'Interface: public function
'
'Abstract: GetConduitJointType
'********************************************************************
Public Function GetConduitJointType(oPart As IJDPart) As JointType
    Const Method = "GetConduitJointType"
    On Error GoTo ErrorHandler
                         
    Dim oPipeComp As IJDPipeComponent
    Set oPipeComp = oPart
    
    If oPipeComp.CommodityType = 1000 Then
        GetConduitJointType = Sleeve
    
    ''Following code is up to each ship company rule
    ''Check the commoditytype of SquareFlange
    ''    ElseIf oPipeComp.CommodityType = 1000 Then
    ''        GetJointType = SquareFlange
    Else
       GetConduitJointType = Flange
    End If
    
Cleanup:
    Set oPipeComp = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'*******************************************************************************************
'*******************************************************************************************
' this next set of routines are generic methods used by all others
'*******************************************************************************************
'*******************************************************************************************

'********************************************************************
'Method: CreateProjection
'
'Interface: public function
'
'Abstract: Create IJProjection object (cylinders)
'********************************************************************
Public Function CreateProjection(oStartX As Double, oStartY As Double, oStartZ As Double, _
                                 oEndX As Double, oEndY As Double, oEndZ As Double, _
                                 oDiameter As Double) As IngrGeom3D.IJProjection
    Const Method = "CreateProjection"
    On Error GoTo ErrorHandler

    Dim oCircle As IngrGeom3D.IJCircle
    Dim oProjection As IngrGeom3D.IJProjection
    Dim oNormal As IJDVector
    
    Dim dProjLength As Double

    Set oNormal = New AutoMath.DVector
    oNormal.x = oEndX - oStartX
    oNormal.y = oEndY - oStartY
    oNormal.z = oEndZ - oStartZ
    oNormal.Length = 1

    dProjLength = Sqr((oEndX - oStartX) ^ 2 + (oEndY - oStartY) ^ 2 + (oEndZ - oStartZ) ^ 2)

    Set oCircle = New Circle3d
    Set oProjection = New Projection3d

    oCircle.DefineByCenterNormalRadius oStartX, oStartY, oStartZ, _
                                       oNormal.x, oNormal.y, oNormal.z, _
                                       oDiameter / 2

    oProjection.DefineByCurve oCircle, oNormal.x, oNormal.y, oNormal.z, dProjLength, False

    Set CreateProjection = oProjection

Cleanup:
    Set oNormal = Nothing
    Set oCircle = Nothing
    Set oProjection = Nothing

    Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetPlateThickness
'
'Interface: public function
'
'Abstract: GetPlateThickness
'********************************************************************
Public Function GetPlateThickness(oStructure As Object) As Double
    Const Method = "GetPlateThickness "
    On Error GoTo ErrorHandler
    
    If TypeOf oStructure Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oStructure
        GetPlateThickness = oPlate.Thickness
        Set oPlate = Nothing
    Else
        GetPlateThickness = 0#
    End If

Cleanup:
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: PlateTightness
'
'Interface: public function
'
'Abstract: PlateTightness
'
'Note: this is being used to place the cuts in a plate fitting (double ring or
'      center flange) in which there is not a port yet for the plate. so if
'      the oStructure is nothing then there is no PlateThickness
'********************************************************************
Public Function PlateTightness(oStructure As Object) As Boolean
    Const Method = "PlateTightness "
    On Error GoTo ErrorHandler
    
    'Default should be false
    PlateTightness = False
    
    If oStructure Is Nothing Then Exit Function
    
    Dim eTypeOfTightness As StructPlateTightness
    
    If TypeOf oStructure Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oStructure
        eTypeOfTightness = oPlate.Tightness
        
        If eTypeOfTightness = NonTight Then
            PlateTightness = False
        Else
            PlateTightness = True
        End If
        
        Set oPlate = Nothing
    End If

Cleanup:
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetInsulationThickness
'
'Interface: public function
'
'Abstract: GetInsulationThickness
'********************************************************************
Public Function GetInsulationThickness(oRtePathFeat As IJRtePathFeat) As Double
    Const Method = "GetInsulationThickness"
    On Error GoTo ErrorHandler
                         
    GetInsulationThickness = 0#
                          
    Dim oPathRun As IJRtePathRun
    Dim oInsulation As IJRteInsulation
    
    Set oPathRun = oRtePathFeat.GetPathRun
    Set oInsulation = oPathRun
    GetInsulationThickness = oInsulation.Thickness
    
Cleanup:
    Set oPathRun = Nothing
    Set oInsulation = Nothing
    
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function

'********************************************************************
'Method: GetHoleSizeByRoundValue
'
'Interface: public function
'
'Abstract: GetHoleSizeByRoundValue
'********************************************************************
Public Function GetHoleSizeByRoundValue(dHoleSize As Double, bRoundValue As Boolean) As Double
    Const Method = "GetHoleSizeByRoundValue"
    On Error GoTo ErrorHandler
    
    If bRoundValue Then
        GetHoleSizeByRoundValue = Round(dHoleSize, 2)
    Else
        GetHoleSizeByRoundValue = dHoleSize
    End If
    
Cleanup:
    Exit Function
  
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function


Public Function GetPipeSleeveOD(oPart As IJDPart) As Double

    Const Method = "GetPipeSleeveOD"
    On Error GoTo ErrorHandler
    
    ' Set initial value.  This value will be returned if the part is not
    ' a pipe sleeve.
    GetPipeSleeveOD = 0#
    
    ' Do not continue if the part does not support IJDAttributes.
    Dim oAttributes As IJDAttributes
    Set oAttributes = oPart
    If oAttributes Is Nothing Then GoTo Cleanup
    
    ' Find the proper interface.
    Dim bInterfaceFound As Boolean
    bInterfaceFound = False
    
    ' Do not continue if there are no attributes.
    If oAttributes.Count <= 0 Then GoTo Cleanup
    
    ' Loop thru each Interface in the Attributes Collection
    Dim oAttributeColl As CollectionProxy
    Dim InterfaceID As Variant
    For Each InterfaceID In oAttributes
        Set oAttributeColl = oAttributes.CollectionOfAttributes(InterfaceID)

        ' verify the current interface Collection is valid
        If oAttributeColl Is Nothing Then GoTo NextInterface
            
        If oAttributeColl.InterfaceInfo.IsHardCoded Then GoTo NextInterface

        ' verify that the current Attribute Interface collection
        ' represents a "User" Attribute interface not System Attribute(??)
        If oAttributeColl.InterfaceInfo.UserName = "IJUAHoleFittingProps" Then
            bInterfaceFound = True
            Exit For
            
        End If
            
NextInterface:
        Set oAttributeColl = Nothing
        Set InterfaceID = Nothing
        
    Next

    If Not bInterfaceFound Then GoTo Cleanup
    
    ' Now that the correct interface has been found.  Locate the
    ' attribute for the OD of the pipe sleeve.
    Dim oAttribute As IJDAttribute
    Dim oAttributeInfo As IJDAttributeInfo
    For Each oAttribute In oAttributeColl
        Set oAttributeInfo = oAttribute.AttributeInfo
        
        ' Return the value of the pipe sleeve OD.
        If oAttributeInfo.UserName = "Outside Diameter" Then
            GetPipeSleeveOD = oAttribute.Value
            GoTo Cleanup
        End If
            
        ' Prepare for the next iteration.
        Set oAttributeInfo = Nothing
        Set oAttribute = Nothing
        
    Next oAttribute
    
    ' Failed to find the OD attribute so return the default value.
    GoTo Cleanup
    
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description

Cleanup:
    Set oAttributes = Nothing
    Set oAttributeColl = Nothing
    Set InterfaceID = Nothing
    Set oAttribute = Nothing
    Set oAttributeInfo = Nothing

End Function ' GetPipeSleeveOD
 
