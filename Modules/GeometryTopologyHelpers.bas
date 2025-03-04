Attribute VB_Name = "GeometryTopologyHelpers"
Option Explicit

' depends on ModelGeometryFacelets (IJWireBody, ...)
' depends on ModelGeometry (DWireBody2, ...)
' depends on GeomOperations (DGeomOpsIntersect, ...)

Private Const NEAR As Integer = 1
Private Const FAR As Integer = 2
'
' acccess Geometry&Topology helpers
'
Private Function GetGeomOpsIntersect() As DGeomOpsIntersect
    Set GetGeomOpsIntersect = New DGeomOpsIntersect
End Function
Private Function GetTopologyIntersect() As IJDTopologyIntersect
    Set GetTopologyIntersect = New DGeomOpsIntersect
End Function
Public Function GetGeomOpsSurfDef() As DGeomOpsSurfDef
    Set GetGeomOpsSurfDef = New DGeomOpsSurfDef
End Function
Public Function GetGeomOpsSolidBody() As DGeomOpsSolidBody
    Set GetGeomOpsSolidBody = New DGeomOpsSolidBody
End Function
Public Function GetTopologyToolBox() As IJDTopologyToolBox
    Set GetTopologyToolBox = New DGeomOpsToolBox
End Function
Public Function GetGeomSweptVolume() As IJGeomSweptVolume
    Set GetGeomSweptVolume = New DGeomOpsSolidBody
End Function
Private Function GetProject() As IProject
    Set GetProject = New Project
End Function
Public Function GetGeometryMisc() As IJGeometryMisc
    Set GetGeometryMisc = New DGeomOpsMisc
End Function
Public Function GetGeometryBoolean() As IJDGeometryBoolean
    Set GetGeometryBoolean = New DGeomOpsMisc
End Function
Public Function GetGeometryBoolean2() As IJGeometryBoolean
    Set GetGeometryBoolean2 = New DGeomOpsMisc
End Function
Public Function GetGeometryCutOut() As IJDGeometryCutOut
    Set GetGeometryCutOut = New DGeomOpsCutOut
End Function
Public Function GetGeomSeamHelper() As IJGeomSeamHelper
    Set GetGeomSeamHelper = New GeomSeamHelper
End Function
'
' processing Geometry
'
Sub Geometry_GetMetrics(pGeometryOfCurve As IJDGeometry, ByRef bIsCurve As Boolean, ByRef bIsLine As Boolean, ByRef dLength As Double, ByRef pVectorOfNormal As IJDVector, ByRef pPositionOfStartPoint As IJDPosition)
    Let bIsCurve = True
    Let bIsLine = False
    If TypeOf pGeometryOfCurve Is IJCurve Then
        Dim pCurve As IJCurve
        Set pCurve = pGeometryOfCurve
        Let dLength = pCurve.Length()
        If TypeOf pCurve Is IJLine Then
            Let bIsLine = True
             Set pVectorOfNormal = Vector_FromLine(pCurve)
             Set pPositionOfStartPoint = Nothing
        Else
            Set pVectorOfNormal = Vector_NormalFromCurve(pCurve)
            Set pPositionOfStartPoint = Position_FromCurve(pCurve, 0)
        End If
    ElseIf TypeOf pGeometryOfCurve Is IJWireBody Then
        Dim pWireBody As IJWireBody
        Set pWireBody = pGeometryOfCurve
        Let dLength = Length_FromWireBody(pWireBody)
        Set pVectorOfNormal = Vector_NormalFromWireBody(pWireBody)
        Set pPositionOfStartPoint = Position_FromWireBody(pWireBody, 0)
    Else
        Let bIsCurve = False
    End If
End Sub

'
' processing ModelBody
'
Public Function ModelBody_FromGeometry(oGeometry As Object) As IJDModelBody
    ' initialize result
    Set ModelBody_FromGeometry = Nothing
        
    ' compute model body
    Dim pModelBody As IJDModelBody
    If TypeOf oGeometry Is IJDModelBody Then
        Set pModelBody = oGeometry
    Else
        ' compute temp geometry
        Dim oGeometryTemp As Object: Set oGeometryTemp = oGeometry
        ' special case for unbounded planes
        If TypeOf oGeometry Is IJPlane Then
            Dim pPlane As IJPlane: Set pPlane = oGeometry
            ' special case for unbounded planes
            If pPlane.BoundaryCount = 0 Then
                Dim pPositionOfRootPoint As IJDPosition: Set pPositionOfRootPoint = Position_FromPlane(pPlane)
                Dim pVectorOfNormal As IJDVector: Set pVectorOfNormal = Vector_FromPlane(pPlane)
                Dim pPlaneByPointNormal As IJPlaneByPointNormal: Set pPlaneByPointNormal = GetGCGeomFactory().PlaneByPointNormal
                Set oGeometryTemp = pPlaneByPointNormal.PlaceGeometry(Nothing, _
                                                                      Point_FromPosition(pPositionOfRootPoint), _
                                                                      Line_FromPositions(Nothing, pPositionOfRootPoint, pPositionOfRootPoint.Offset(pVectorOfNormal)), _
                                                                      1000#)
            End If
        End If

        Call GetGeometryMisc().CreateModelGeometryFromGType(Nothing, oGeometryTemp, Nothing, pModelBody)
    End If
    
    ' return result
    Set ModelBody_FromGeometry = pModelBody
End Function
Function ModelBodies_FromModelBodyLumps(pModelBody As IJModelBody) As IJElements
    Dim pEnumUnknown As IEnumUnknown
    Call GetGeometryMisc.ExplodeModelBodyByLumps(Nothing, pModelBody, pEnumUnknown)
    
    Dim pElements As IJElements
    Call GetAcisHelper().GetCollectionFromEnum(pEnumUnknown, pElements)
    Set ModelBodies_FromModelBodyLumps = pElements
End Function
Sub ModelBody_UpdateFromModelBody(pModelBody1 As IJDModelBody, pModelBody0 As IJDModelBody)
    Call pModelBody1.UpdateBodyData(pModelBody0, True)
End Sub
'
' processing PointsGraphBody
'
Function Positions_FromPointsGraphBody(pPointsGraphBody As IJPointsGraphBody) As IJElements
    ' initialize result
    Set Positions_FromPointsGraphBody = Nothing
    
    ' create enumerator
    Dim pEnumVariantOfPositions As IEnumVARIANT
    Dim iCount As Integer
    Call pPointsGraphBody.EnumPositions(pEnumVariantOfPositions, iCount)
    If iCount = 0 Then Exit Function
    
    ' create collection
    Dim pElementsOfPositions As IJElements
    Call GetStructSymbolTools().TransformIEnumVariantToIJElements(pEnumVariantOfPositions, pElementsOfPositions)
    
    ' return result
    Set Positions_FromPointsGraphBody = pElementsOfPositions
End Function
'
' processing WireBody
'
Function IsObjectCurve(pObject As Object) As Boolean
    If TypeOf pObject Is IJCurve _
    Or TypeOf pObject Is IJWireBody Then
        Let IsObjectCurve = True
    Else
        IsObjectCurve = False
    End If
End Function
Function Position_FromWireBody(pWireBody As IJWireBody, bFlag As Integer) As IJDPosition
    ' retrieve coordinates
    Dim pPositionStart As IJDPosition, pPositionEnd As IJDPosition
    Call pWireBody.GetEndPoints(pPositionStart, pPositionEnd)
       
    ' set position
    Dim pPosition As IJDPosition
    If bFlag = 0 Then
        Set pPosition = pPositionStart
    Else
        Set pPosition = pPositionEnd
    End If
    
    ' return result
    Set Position_FromWireBody = pPosition
End Function
Public Function Position_GetCOGFromSurfaceBody(pSurfaceBody As IJSurfaceBody) As IJDPosition
    ' get center point
    Dim pPositionOfCOG As IJDPosition
    Call pSurfaceBody.GetCenterOfGravity(pPositionOfCOG)

    ' return it
    Set Position_GetCOGFromSurfaceBody = pPositionOfCOG
End Function
Public Function Vector_NormalFromWireBody(pWireBody As IJWireBody) As IJDVector
    Set Vector_NormalFromWireBody = Nothing
    
    Dim pPositionOfRootPoint As IJDPosition
    Dim pVectorOfNormal As IJDVector
    On Error Resume Next
    Call pWireBody.IsPlanar(pPositionOfRootPoint, pVectorOfNormal)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If Not pVectorOfNormal Is Nothing Then
        Set Vector_NormalFromWireBody = pVectorOfNormal
    End If
End Function
Public Function Length_FromWireBody(pWireBody As IJWireBody) As Double
    Dim pModelBody As IJDModelBody
    Set pModelBody = pWireBody
                
    Dim dLength As Double, dArea As Double, dVolume As Double
    Dim dAccuracy As Double
    Call pModelBody.GetDimMetrics(EPSILON, dAccuracy, dLength, dArea, dVolume)
    Let Length_FromWireBody = dLength
End Function
Public Function Position_ProjectOnCurve(pPosition As IJDPosition, oCurve As Object) As IJDPosition
    Dim pWireBody As IJWireBody: Set pWireBody = ModelBody_FromGeometry(oCurve)
    
    ' compute projected point
    Dim pPositionOfResult As IJDPosition
    If True Then
        Dim pVectorOfResult As IJDVector
        On Error Resume Next
        Call GetTopologyToolBox().ProjectPointOnWireBody(pWireBody, pPosition, pPositionOfResult, pVectorOfResult)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    Set Position_ProjectOnCurve = pPositionOfResult
End Function
Public Function Position_AtDistanceAlongOnCurve(pPositionOfReferencePoint As IJDPosition, oCurve As Object, dDistance As Double, pPositionOfTrackPoint As IJDPosition, eTrackFlag As eTrackFlag) As IJDPosition
    Dim pWireBody As IJWireBody: Set pWireBody = ModelBody_FromGeometry(oCurve)
    
    ' compute pseudo track point, which should project on the curve, because we position it close to the reference point along the tangent
    Dim pPositionOfPseudoTrackPoint As IJDPosition
    If True Then
        ' compute the vector from the reference point to the track point
        Dim pVectorOfTrackPoint As IJDVector: Set pVectorOfTrackPoint = pPositionOfTrackPoint.Subtract(pPositionOfReferencePoint)
        
        ' compute the tangent vector to the curve
        Dim pVectorOfTangency As IJDVector: Set pVectorOfTangency = Vector_FromTangentToWireBodyAtPosition(oCurve, pPositionOfReferencePoint)
        
        ' reposition a pseudo track point along the tangent close to the reference point
        Dim pVectorOfOffset As IJDVector: Set pVectorOfOffset = Vector_Scale(pVectorOfTangency, dDistance / 10)
        
        ' if the real track point is on the other side, then re-position the pseudo track point on the other side
        If pVectorOfTangency.Dot(pVectorOfTrackPoint) < 0 Then Set pVectorOfOffset = Vector_Scale(pVectorOfOffset, -1)
        
        ' if the track flag is GCFar,  then re-position the pseudo track point on the other side
        If eTrackFlag = GCFar Then Set pVectorOfOffset = Vector_Scale(pVectorOfOffset, -1)
      
        ' position the pseudo track point
        Set pPositionOfPseudoTrackPoint = pPositionOfReferencePoint.Offset(pVectorOfOffset)
    End If
    
    ' compute point along
    Dim pPositionOfResult As IJDPosition
    If True Then
        On Error Resume Next
        Call GetTopologyToolBox().GetPointAlongCurveAtDistance(pWireBody, pPositionOfReferencePoint, dDistance, pPositionOfPseudoTrackPoint, pPositionOfResult)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    Set Position_AtDistanceAlongOnCurve = pPositionOfResult
End Function
Public Function WireBody_AtSurfaceBodiesIntersection(pSurfaceBody1 As IJSurfaceBody, pSurfaceBody2 As IJSurfaceBody, _
                                                     pPointOfTrackPoint As IJPoint, eTrackFlag As eTrackFlag, lIntersectionOption As Long) As IJWireBody
    ' initialize result
    Set WireBody_AtSurfaceBodiesIntersection = Nothing
    
    ' compute all possible wirebodies
    Dim pElementsOfWireBodies As IJElements:
    If True Then
        ' compute intersection
        Dim pWireBody As IJWireBody
        If True Then
            On Error Resume Next
            Call GetGeomOpsIntersect().PlaceIntersectionObject(Nothing, pSurfaceBody1, pSurfaceBody2, Nothing, pWireBody, False, lIntersectionOption)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        ' build a collection of wirebodies
        Set pElementsOfWireBodies = ModelBodies_FromModelBodyLumps(pWireBody)
    End If
    
    ' choose the nearer/further wirebody
    Dim pWireBodyOfResult As IJWireBody
    Set pWireBodyOfResult = WireBody_FromWireBodies(pElementsOfWireBodies, pPointOfTrackPoint, eTrackFlag)
    
    ' return result
    Set WireBody_AtSurfaceBodiesIntersection = pWireBodyOfResult
End Function
Public Function WireBody_ByProjection(pWireBodyOfCurve As IJWireBody, _
                                      pSurfaceBodyOfSurface As IJSurfaceBody, _
                                      pVectorOfProjection As IJDVector, _
                                      pPointOfTrackPoint As IJPoint, eTrackFlag As eTrackFlag, lIntersectionOption As Long) As IJWireBody
     ' initialize result
    Set WireBody_ByProjection = Nothing
    
    ' compute possible wirebodies
    Dim pElementsOfWireBodies As IJElements:
    If True Then
        ' compute projection
        Dim pWireBody As IJWireBody
        If True Then
            Dim lCountOfIntersections As Long
            On Error Resume Next
            Call GetProject().CurveAlongVectorOnToSurface(Nothing, pSurfaceBodyOfSurface, pWireBodyOfCurve, pVectorOfProjection, _
                                                          Nothing, pWireBody, lIntersectionOption)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        ' build a collection of wirebodies
        Set pElementsOfWireBodies = ModelBodies_FromModelBodyLumps(pWireBody)
    End If
    
    ' choose the nearer/further wirebody
    Dim pWireBodyOfResult As IJWireBody
    Set pWireBodyOfResult = WireBody_FromWireBodies(pElementsOfWireBodies, pPointOfTrackPoint, eTrackFlag)
    
    ' return result
    Set WireBody_ByProjection = pWireBodyOfResult
End Function
Public Function Position_AtWireBodiesIntersection(pWireBody1 As IJWireBody, _
                                                  pWireBody2 As IJWireBody, _
                                                  pPointOfTrackPoint As IJPoint, _
                                                  eTrackFlag As eTrackFlag) As IJDPosition
    ' initialize result
    Set Position_AtWireBodiesIntersection = Nothing
        
    ' compute all possible positions of intersection points
    Dim pElementsOfPositions As IJElements
    If True Then
        ' compute intersection
        Dim pPointsGraphBody As IJPointsGraphBody
        If True Then
            On Error Resume Next
            Call GetGeomOpsIntersect().PlaceIntersectionObject(Nothing, pWireBody1, pWireBody2, Nothing, pPointsGraphBody)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        
        ' build a collection of positions
        Set pElementsOfPositions = Positions_FromPointsGraphBody(pPointsGraphBody)
        
        If pElementsOfPositions Is Nothing Then Exit Function
        If pElementsOfPositions.Count = 0 Then Exit Function
    End If
    
    ' choose the nearer/further intersection position
    Dim pPositionOfResult As IJDPosition:
    Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Nothing
    If Not pPointOfTrackPoint Is Nothing Then Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
    Set pPositionOfResult = Position_FromPositionsV1(pElementsOfPositions, pPositionOfTrackPoint, eTrackFlag)
    
    ' return result
    Set Position_AtWireBodiesIntersection = pPositionOfResult
End Function
Public Function Position_AtWireBodiesMinimumDistance(pWireBody1 As IJWireBody, _
                                                     pWireBody2 As IJWireBody) As IJDPosition
    Dim pPosition As IJDPosition
    
    Dim pModelBody1 As IJDModelBody: Set pModelBody1 = pWireBody1
    If True Then
        On Error Resume Next
        Dim pPosition1 As IJDPosition, pPosition2 As IJDPosition, dDistance As Double
        Call pModelBody1.GetMinimumDistance(pWireBody2, pPosition1, pPosition2, dDistance)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    Set pPosition = pPosition1.Offset(Vector_Scale(pPosition2.Subtract(pPosition1), 0.5))
    
    ' return result
    Set Position_AtWireBodiesMinimumDistance = pPosition
End Function
Public Function Position_AtSurfaceBodyAndWireBodyIntersection(pSurfaceBody As IJSurfaceBody, _
                                                              pWireBody As IJWireBody, _
                                                              pPointOfTrackPoint As IJPoint, _
                                                              eTrackFlag As eTrackFlag) As IJDPosition
    ' initialize result
    Set Position_AtSurfaceBodyAndWireBodyIntersection = Nothing
    
    ' compute possible positions
    Dim pElementsOfPositions As IJElements
    If True Then
        ' compute intersection
        Dim pPointsGraphBody As IJPointsGraphBody
        If True Then
            Set pPointsGraphBody = New DPointsGraphBody2
            On Error Resume Next
            Call GetGeomOpsIntersect().ModifyIntersectionObject(Nothing, pSurfaceBody, pWireBody, pPointsGraphBody, False)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
        
        ' build a collection of positions
        Set pElementsOfPositions = Positions_FromPointsGraphBody(pPointsGraphBody)
        If pElementsOfPositions Is Nothing Then Exit Function
        If pElementsOfPositions.Count = 0 Then Exit Function
    End If
    
    ' choose the nearer/further intersection position
    Dim pPositionOfResult As IJDPosition:
    Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Nothing
    If Not pPointOfTrackPoint Is Nothing Then Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
    Set pPositionOfResult = Position_FromPositionsV1(pElementsOfPositions, pPositionOfTrackPoint, eTrackFlag)
    
    ' return result
    Set Position_AtSurfaceBodyAndWireBodyIntersection = pPositionOfResult
End Function
Public Function Position_AtCurvesMinimumDistance(pCurve1 As IJCurve, pCurve2 As IJCurve, ByRef pVectorBetweenCurves As IJDVector) As IJDPosition
    ' compute points on each curve
    Dim pPosition1 As IJDPosition
    Dim pPosition2 As IJDPosition
    If True Then
        Dim pModelBody1 As IJDModelBody: Set pModelBody1 = ModelBody_FromGeometry(pCurve1)
        Dim pModelBody2 As IJDModelBody: Set pModelBody2 = ModelBody_FromGeometry(pCurve2)
        Dim dMinimumDistance As Double
        On Error Resume Next
        Call pModelBody1.GetMinimumDistance(pModelBody2, pPosition1, pPosition2, dMinimumDistance)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' compute mid-point
    Dim pPositionOfResult As IJDPosition
    If True Then
        Dim pVector As IJDVector: Set pVector = pPosition2.Subtract(pPosition1)
        Set pPositionOfResult = pPosition1.Offset(Vector_Scale(pVector, 0.5))
    End If
    
    ' return result
    Call pVectorBetweenCurves.Set(pVector.x, pVector.y, pVector.z)
    Set Position_AtCurvesMinimumDistance = pPositionOfResult
End Function
Public Function Curve_AtSurfacesIntersection(pSurface1 As IJSurface, pSurface2 As IJSurface, _
                                             pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJCurve
    ' initialize result
    Set Curve_AtSurfacesIntersection = Nothing
    
    ' compute possible intersection curves
    Dim pElementsOfCurves As IJElements
    If True Then
        Dim eGeom3dIntersectConstants As Geom3dIntersectConstants
        On Error Resume Next
        Call pSurface1.Intersect(pSurface2, pElementsOfCurves, eGeom3dIntersectConstants)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
        If pElementsOfCurves Is Nothing Then Exit Function
        If pElementsOfCurves.Count = 0 Then Exit Function
    End If
    
    ' choose neare/further curve
    Dim pCurveOfResult As IJCurve: Set pCurveOfResult = Curve_FromCurves(pElementsOfCurves, pPointOfTrackPoint, iTrackFlag)
    
    ' return result
    Set Curve_AtSurfacesIntersection = pCurveOfResult
End Function
Function WireBody_FromWireBodies(pElementsOfWireBodies As IJElements, pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJWireBody
    ' initialize result
    Set WireBody_FromWireBodies = Nothing
    
    Dim pWireBodyOfResultCurve As IJWireBody
    If pPointOfTrackPoint Is Nothing Then
        Set pWireBodyOfResultCurve = pElementsOfWireBodies(1)
    Else
        Dim pPositionOfTrackPoint As IJDPosition: Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
        Dim dMinimumDistanceMini As Double: Let dMinimumDistanceMini = BIG_EXTENSION
        Dim dMinimumDistanceMaxi As Double: Let dMinimumDistanceMaxi = -BIG_EXTENSION
    
        Set pWireBodyOfResultCurve = pElementsOfWireBodies.Item(1)
        Dim i As Integer
        For i = 1 To pElementsOfWireBodies.Count
            Dim pWireBody As IJWireBody: Set pWireBody = pElementsOfWireBodies.Item(i)
            Dim pPositionOfProjection As IJDPosition
            Dim pVector As IJDVector
            On Error Resume Next
            Call GetTopologyToolBox().ProjectPointOnWireBody(pWireBody, pPositionOfTrackPoint, pPositionOfProjection, pVector)
            If Err.Number <> 0 Then
                ' projected point is outside the curve, compute distance to end points
                On Error GoTo 0
                Dim pPosition0 As IJDPosition: Set pPosition0 = Position_FromWireBody(pWireBody, 0)
                Dim pPosition1 As IJDPosition: Set pPosition1 = Position_FromWireBody(pWireBody, 1)
                Dim dMinimumDistance0 As Double: Let dMinimumDistance0 = pPositionOfTrackPoint.DistPt(pPosition0)
                Dim dMinimumDistance1 As Double: Let dMinimumDistance1 = pPositionOfTrackPoint.DistPt(pPosition1)
                If iTrackFlag = NEAR Then
                    If dMinimumDistance0 < dMinimumDistanceMini Then
                        Let dMinimumDistanceMini = dMinimumDistance0
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                    If dMinimumDistance1 < dMinimumDistanceMini Then
                        Let dMinimumDistanceMini = dMinimumDistance1
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                Else
                    If dMinimumDistance0 > dMinimumDistanceMaxi Then
                        Let dMinimumDistanceMaxi = dMinimumDistance0
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                    If dMinimumDistance1 > dMinimumDistanceMaxi Then
                        Let dMinimumDistanceMaxi = dMinimumDistance1
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                End If
            Else
                ' projected point is inside the curve
                On Error GoTo 0
                Dim dMinimumDistance As Double: Let dMinimumDistance = pPositionOfTrackPoint.DistPt(pPositionOfProjection)
                If iTrackFlag = NEAR Then
                    If dMinimumDistance < dMinimumDistanceMini Then
                        Let dMinimumDistanceMini = dMinimumDistance
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                Else
                    If dMinimumDistance > dMinimumDistanceMaxi Then
                        Let dMinimumDistanceMaxi = dMinimumDistance
                        Set pWireBodyOfResultCurve = pWireBody
                    End If
                End If
            End If
        Next
    End If
    
    ' return result
    Set WireBody_FromWireBodies = pWireBodyOfResultCurve
End Function
Sub WireBody_Orientate(pWireBody As IJWireBody, pLocalCoordinateSystem As IJLocalCoordinateSystem)
    Dim pVectorOfLine As IJDVector
    Set pVectorOfLine = Position_FromWireBody(pWireBody, 1).Subtract(Position_FromWireBody(pWireBody, 0))
    Let pVectorOfLine.Length = 1
            
    ' revert wirebody
    If IsVectorInversionNeeded(pVectorOfLine, pLocalCoordinateSystem) Then Call pWireBody.ReverseTopology
End Sub
Function Curve_FromCurves(pElementsOfCurves As IJElements, pPointOfTrackPoint As IJPoint, ByVal iTrackFlag As Integer) As IJCurve
    ' initialize result
    Set Curve_FromCurves = Nothing
    
    ' find nearer/further curve
    Dim pCurveOfResultCurve As IJCurve
    If pPointOfTrackPoint Is Nothing Then
        Set pCurveOfResultCurve = pElementsOfCurves(1)
    Else
        ' transform the collectionof curves into a collection of wirebodies
        Dim pElementsOfWireBodies As IJElements: Set pElementsOfWireBodies = New JObjectCollection
            If True Then
            Dim i As Integer
            For i = 1 To pElementsOfCurves.Count
                Dim pWireBody As IJWireBody: Set pWireBody = ModelBody_FromGeometry(pElementsOfCurves(i))
                Call pElementsOfWireBodies.Add(pWireBody)
            Next
        End If
        
        ' find nearer/further wirebody
        Dim pWireBodyOfResultWireBody As IJWireBody
        Set pWireBodyOfResultWireBody = WireBody_FromWireBodies(pElementsOfWireBodies, pPointOfTrackPoint, iTrackFlag)
        
        ' retrieve the mapping between the resulting wirebody and a curve in the collection of curves
        For i = 1 To pElementsOfWireBodies.Count
            If pWireBodyOfResultWireBody Is pElementsOfWireBodies(i) Then
                Set pCurveOfResultCurve = pElementsOfCurves(i)
                Exit For
            End If
        Next
    End If
    
    ' return result
    Set Curve_FromCurves = pCurveOfResultCurve
End Function
Function WireBody_AsLine(pWireBodyLinear As IJWireBody) As IJWireBody
    Dim pPositionOfStartPoint As IJDPosition
    Dim pPositionOfEndPoint As IJDPosition
    Call pWireBodyLinear.GetEndPoints(pPositionOfStartPoint, pPositionOfEndPoint)
        
    Set WireBody_AsLine = ModelBody_FromGeometry(Line_FromPositions(Nothing, pPositionOfStartPoint, pPositionOfEndPoint))
End Function
Public Function WireBody_GetReversed(ByVal pWireBodyToReverse As IJWireBody) As IJWireBody
    ' prepare result
    Dim pWireBodyReversed As IJWireBody: Set pWireBodyReversed = New DWireBody2: Call ModelBody_UpdateFromModelBody(pWireBodyReversed, pWireBodyToReverse)

    ' reverse the body
    Call pWireBodyReversed.ReverseTopology

    ' return the result
    Set WireBody_GetReversed = pWireBodyReversed
End Function
Public Sub PositionAndVector_FromTangentToWireBodyAtPosition(ByVal pWireBody As IJWireBody, ByVal pPositionToProject As AutoMath.IJDPosition, _
                                                          ByRef pPositionOfProjectedPoint As AutoMath.IJDPosition, ByRef pVectorOfTangent As AutoMath.IJDVector)
    ' initialize results
    Set pPositionOfProjectedPoint = Nothing
    Set pVectorOfTangent = Nothing

    ' do the projection
    If True Then
        On Error Resume Next
        Call GetTopologyToolBox().ProjectPointOnWireBody(pWireBody, pPositionToProject, pPositionOfProjectedPoint, pVectorOfTangent)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
End Sub
Public Function Vector_FromPlanarWireBody(pWireBody As IJWireBody) As IJDVector
    ' prepare result
    Dim pVector As IJDVector
    
    Dim bIsSelfIntersecting As Boolean
    Dim bIsPlanar As Boolean
    Call pWireBody.GetProperties(bIsSelfIntersecting, bIsPlanar)
    If bIsPlanar Then
        Dim pPositionOfPlane As IJDPosition
        Dim pVectorOfPlane As IJDVector
        Call pWireBody.IsPlanar(pPositionOfPlane, pVectorOfPlane)
        Set pVector = pVectorOfPlane
    End If
    
    ' return result
    Set Vector_FromPlanarWireBody = pVector
End Function
'
' SurfaceBody
'
Function IsObjectSurface(pObject As Object) As Boolean
    If TypeOf pObject Is IJSurface _
    Or TypeOf pObject Is IJSurfaceBody Then
        Let IsObjectSurface = True
    Else
        IsObjectSurface = False
    End If
End Function
Public Function Vector_FromSurfaceBody(pSurfaceBody As IJSurfaceBody, pPosition As IJDPosition) As IJDVector
    Dim pVectorOfNormal As IJDVector
    Call pSurfaceBody.GetNormalFromPosition(pPosition, pVectorOfNormal)
    Set Vector_FromSurfaceBody = pVectorOfNormal
End Function
Public Function Vector_FromTangentToWireBodyAtPosition(pWireBody As IJWireBody, pPositionToProject As IJDPosition) As IJDVector
    Dim pPositionProjected As IJDPosition
    Dim pVectorTangent As IJDVector
    If True Then
        On Error Resume Next
        Call GetTopologyToolBox().ProjectPointOnWireBody(pWireBody, pPositionToProject, pPositionProjected, pVectorTangent)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' return result
    Set Vector_FromTangentToWireBodyAtPosition = pVectorTangent
End Function
Public Function SurfaceBody_ByOffset(pSurfaceBody As IJSurfaceBody, dOffset As Double) As IJSurfaceBody
    ' initialize result
    Set SurfaceBody_ByOffset = Nothing
    
    ' prepare inputs
    Dim pSurfaceBodyToOffset As IJSurfaceBody
    If dOffset > 0 Then
        Set pSurfaceBodyToOffset = pSurfaceBody
    Else
        Set pSurfaceBodyToOffset = New DSurfaceBody2
        Call ModelBody_UpdateFromModelBody(pSurfaceBodyToOffset, pSurfaceBody)
        Call pSurfaceBodyToOffset.ReverseTopologyOrientation
    End If
    
    Dim pSurfaceBodyByOffset As IJSurfaceBody
    If True Then
        On Error Resume Next
        Call GetGeomOpsSurfDef().PlaceSurfaceByOffset(Nothing, pSurfaceBodyToOffset, Abs(dOffset), Nothing, pSurfaceBodyByOffset)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' return result
    Set SurfaceBody_ByOffset = pSurfaceBodyByOffset
End Function
Public Sub PositionAndVector_FromNormalToSurfaceBodyAtPosition(ByVal pSurfaceBody As IJSurfaceBody, ByVal pPositionToProject As AutoMath.IJDPosition, _
                                                               ByRef pPositionOfProjectedPoint As AutoMath.IJDPosition, ByRef pVectorOfNormal As AutoMath.IJDVector)
    ' initialize results
    Set pPositionOfProjectedPoint = Nothing
    Set pVectorOfNormal = Nothing

    ' do the projection
    If True Then
        On Error Resume Next
        Call GetTopologyToolBox().ProjectPointOnSurfaceBody(pSurfaceBody, pPositionToProject, pPositionOfProjectedPoint, pVectorOfNormal)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If
End Sub
Public Function SurfaceBody_ByOffset1(ByVal pSurfaceBodyToOffset As IJSurfaceBody, ByVal dOffset As Double) As IJSurfaceBody
    ' initialize result
    Dim pSurfaceBodyByOffset As IJSurfaceBody: Set pSurfaceBodyByOffset = Nothing

    If True Then
        On Error Resume Next
        Call GetGeomOpsSurfDef().PlaceSurfaceByOffset(Nothing, pSurfaceBodyToOffset, Abs(dOffset), Nothing, pSurfaceBodyByOffset)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set SurfaceBody_ByOffset1 = Nothing
            Exit Function
        End If
        On Error GoTo 0
    End If

    ' return result
    Set SurfaceBody_ByOffset1 = pSurfaceBodyByOffset
End Function
Public Function SurfaceBody_GetReversed(ByVal pSurfaceBodyToReverse As IJSurfaceBody) As IJSurfaceBody
    ' prepare result
    Dim pSurfaceBodyReversed As IJSurfaceBody: Set pSurfaceBodyReversed = New DSurfaceBody2: Call ModelBody_UpdateFromModelBody(pSurfaceBodyReversed, pSurfaceBodyToReverse)

    ' reverse the body
    Call pSurfaceBodyReversed.ReverseTopologyOrientation

    ' return the result
    Set SurfaceBody_GetReversed = pSurfaceBodyReversed
End Function
Function AreSurfacesOverlapping(oSurface1 As Object, oSurface2 As Object) As Boolean
    Dim bAreSurfacesOverlapping As Boolean
    Call GetTopologyIntersect().HasOverlappingGeometry(oSurface1, oSurface2, bAreSurfacesOverlapping)
    
    Let AreSurfacesOverlapping = bAreSurfacesOverlapping
End Function
Function AreModelBodiesSame(pModelBody0 As IJDModelBody, pModelBody1 As IJDModelBody) As Boolean
    Call pModelBody0.CompareTopologyAndGeometry(pModelBody1, AreModelBodiesSame)
End Function
Sub SurfaceBody_OrientateFollowingLine(pSurfaceBody As IJSurfaceBody, pLine As IJLine, bSameOrientation As Boolean)
    Dim pVectorOfSurface As IJDVector
    Set pVectorOfSurface = Vector_FromPlane(pSurfaceBody)
        
    Dim pVectorOfLine As IJDVector
    Set pVectorOfLine = Vector_FromLine(pLine)
    
    Dim bRelativeOrientation As Boolean: Let bRelativeOrientation = pVectorOfSurface.Dot(pVectorOfLine) > 0
    
    If Not bRelativeOrientation = bSameOrientation Then _
        Call pSurfaceBody.ReverseTopologyOrientation
End Sub
Sub SurfaceBody_OrientateFollowingVector(pSurfaceBody As IJSurfaceBody, pVectorOfLine As IJDVector, bSameOrientation As Boolean)
    Dim pVectorOfSurface As IJDVector
    Set pVectorOfSurface = Vector_FromPlane(pSurfaceBody)
        
    Dim bRelativeOrientation As Boolean: Let bRelativeOrientation = pVectorOfSurface.Dot(pVectorOfLine) > 0
    
    If Not bRelativeOrientation = bSameOrientation Then _
        Call pSurfaceBody.ReverseTopologyOrientation
End Sub
Function ModelBody_ByModelBodies(ByVal oResourceManager As Object, ByVal pElementsOfModelBodies As IMSCoreCollections.IJElements, Optional ByVal lIntersectionOption As Long = 0) As IJDModelBody
        ' prepare result
        Dim pModelBody As IJDModelBody: Set pModelBody = Nothing

        If True Then
            On Error Resume Next
            Call GetGeometryBoolean().PlaceMergedModelBodyFromCollection(oResourceManager, pElementsOfModelBodies, True, Nothing, pModelBody)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Set ModelBody_ByModelBodies = Nothing:
                Exit Function
            End If
            On Error GoTo 0
        End If

        ' try to merge the ribbons, if possible
        If lIntersectionOption = 2 Then
            Dim pElementsOfSurfaceBodiesOfRibbons As IMSCoreCollections.IJElements: Set pElementsOfSurfaceBodiesOfRibbons = New JObjectCollection
            Dim pWireBody As IJWireBody: Set pWireBody = Nothing
            Dim i As Integer
            For i = 1 To pElementsOfModelBodies.Count
                Set pWireBody = pElementsOfModelBodies(i)
                Dim pSurfaceBodyOfAttribute As IJSurfaceBody
                Set pSurfaceBodyOfAttribute = WireBody_GetSurfaceAttribute(pWireBody)
                If Not pSurfaceBodyOfAttribute Is Nothing And Not pElementsOfSurfaceBodiesOfRibbons Is Nothing Then
                    Call pElementsOfSurfaceBodiesOfRibbons.Add(pSurfaceBodyOfAttribute)
                Else
                    Set pElementsOfSurfaceBodiesOfRibbons = Nothing
                End If
            Next

            If Not pElementsOfSurfaceBodiesOfRibbons Is Nothing Then
                Dim pSurfaceBodyOfRibbons As IJSurfaceBody: Set pSurfaceBodyOfRibbons = Nothing
                On Error Resume Next
                Set pSurfaceBodyOfRibbons = New DSurfaceBody2
                Dim lNumPatches As Long
                Call GetAcisHelper().PlaceTolerantStitchedAndOrientedTopology(Nothing, pElementsOfSurfaceBodiesOfRibbons, False, Nothing, pSurfaceBodyOfRibbons, lNumPatches, False, 0.001)
                'Set pSurfaceBodyOfRibbons = ModelBody_ByModelBodies(oResourceManager, pElementsOfSurfaceBodiesOfRibbons, 0)
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Dim pWireBodyOfResult As IJWireBody: Set pWireBodyOfResult = pModelBody
                    Call GetGeomSeamHelper().SetSurfaceModelBodyAttribute(pWireBodyOfResult, pSurfaceBodyOfRibbons)
                Else
                    On Error GoTo 0
                End If
            End If
        End If

        ' return result
        Set ModelBody_ByModelBodies = pModelBody
    End Function
    Public Function WireBody_GetSurfaceAttribute(pWireBody As IJWireBody) As IJSurfaceBody
        ' prepare result
        Dim pSurfaceBody As IJSurfaceBody: Set pSurfaceBody = Nothing
        
        Dim bHasSurfaceBodyAttribute As Boolean
        Call GetGeomSeamHelper().HasSurfaceModelBodyAttribute(pWireBody, bHasSurfaceBodyAttribute)
        If bHasSurfaceBodyAttribute Then
           Call GetGeomSeamHelper().GetSurfaceModelBodyAttribute(pWireBody, pSurfaceBody)
        End If
        
        ' return result
        Set WireBody_GetSurfaceAttribute = pSurfaceBody
    End Function
    Public Function Distance_PointToSurfaceBody(pPoint As IJPoint, pSurfaceBody As IJSurfaceBody) As Double
        Dim pPositionOfPoint As IJDPosition: Set pPositionOfPoint = Position_FromPoint(pPoint)
        Dim pModelBody As IJDModelBody: Set pModelBody = pSurfaceBody
        Dim pPositionOfClosestPoint As IJDPosition
        Dim dDistance As Double
        Call pModelBody.GetMinimumDistanceFromPosition(pPositionOfPoint, pPositionOfClosestPoint, dDistance)
        Distance_PointToSurfaceBody = dDistance
'        Dim pPositionOfProjectedPoint As IJDPosition
'        Dim pVector As IJDVector
'        Call PositionAndVector_FromNormalToSurfaceBodyAtPosition(pSurfaceBody, pPositionOfPoint, pPositionOfProjectedPoint, pVector)
'        Distance_PointToSurfaceBody = pPositionOfPoint.DistPt(pPositionOfProjectedPoint)
    End Function
    Public Function Distance_SurfaceBodyToSurfaceBody(pSurfaceBody1 As IJSurfaceBody, pSurfaceBody2 As IJSurfaceBody) As Double
        Dim pModelBody1 As IJDModelBody: Set pModelBody1 = pSurfaceBody1
        Dim pPositionOfClosestPoint1 As IJDPosition
        Dim pPositionOfClosestPoint2 As IJDPosition
        Dim dDistance As Double
        Call pModelBody1.GetMinimumDistance(pSurfaceBody2, pPositionOfClosestPoint1, pPositionOfClosestPoint2, dDistance)
        Distance_SurfaceBodyToSurfaceBody = dDistance
'        Dim pPositionOfProjectedPoint As IJDPosition
'        Dim pVector As IJDVector
'        Call PositionAndVector_FromNormalToSurfaceBodyAtPosition(pSurfaceBody, pPositionOfPoint, pPositionOfProjectedPoint, pVector)
'        Distance_PointToSurfaceBody = pPositionOfPoint.DistPt(pPositionOfProjectedPoint)
    End Function
    Public Function Position_AtSurfaceBodiesIntersection(pSurfaceBody1 As IJSurfaceBody, pSurfaceBody2 As IJSurfaceBody, pSurfaceBody3 As IJSurfaceBody, pPointOfTrackPoint As IJPoint, eTrackFlag As eTrackFlag) As IJDPosition
        ' prepare result
        Dim pPositionOfIntersection As IJDPosition: Set pPositionOfIntersection = Nothing
        
        Dim oResult As Object
        Call GetTopologyIntersect().GetSliceBySurfaces(Nothing, pSurfaceBody1, pSurfaceBody2, pSurfaceBody3, oResult)
        
        Dim pModelBody As IJDModelBody: Set pModelBody = oResult
        Dim pElementsOfPositions As IJElements: Set pElementsOfPositions = pModelBody.GetPhysicalVertices
        
        If pElementsOfPositions.Count > 0 Then
            Dim pPositionOfTrackPoint As IJDPosition: If Not pPointOfTrackPoint Is Nothing Then Set pPositionOfTrackPoint = Position_FromPoint(pPointOfTrackPoint)
            Set pPositionOfIntersection = Position_FromPositionsV1(pElementsOfPositions, pPositionOfTrackPoint, eTrackFlag)
        End If
        
        ' return result
        Set Position_AtSurfaceBodiesIntersection = pPositionOfIntersection
    End Function
Public Sub SurfaceBody_Monikerize(pSurfaceBody As IJSurfaceBody, iOpt As Integer, iOpr As Integer, iCtx As Integer, iXid As Integer)
    Dim oGeomOpsToolBox As DGeomOpsToolBox: Set oGeomOpsToolBox = New DGeomOpsToolBox
    Dim lCountOfProcessedFaces As Long: Let lCountOfProcessedFaces = 0
    Call oGeomOpsToolBox.SetGenericAttributeOnSurfaceBodyFaces(pSurfaceBody, "JSOpt", iOpt, lCountOfProcessedFaces)
    Call oGeomOpsToolBox.SetGenericAttributeOnSurfaceBodyFaces(pSurfaceBody, "JSOpr", iOpr, lCountOfProcessedFaces)
    Call oGeomOpsToolBox.SetGenericAttributeOnSurfaceBodyFaces(pSurfaceBody, "JSCtx", iCtx, lCountOfProcessedFaces)
    Call oGeomOpsToolBox.SetGenericAttributeOnSurfaceBodyFaces(pSurfaceBody, "JSXid", iXid, lCountOfProcessedFaces)
End Sub

