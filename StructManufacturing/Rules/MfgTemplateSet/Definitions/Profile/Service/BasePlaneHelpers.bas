Attribute VB_Name = "BasePlaneHelpers"
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   BasePlane Helpers Bas Module
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Private Const MODULE = "GSCADStrMfgTemplate.BasePlaneService"
Private Const DISTANCETOLERANCE = 0.000001  'Distance Tolerance = 1E-06

Private Function PointsToArray(oPos1 As IJDPosition, oPos2 As IJDPosition, oPos3 As IJDPosition) As Double()
Const METHOD = "PointsToArray"
On Error GoTo ErrorHandler
    
    Dim oPointsDouble(8) As Double
    
    oPointsDouble(0) = oPos1.x
    oPointsDouble(1) = oPos1.y
    oPointsDouble(2) = oPos1.z
    
    oPointsDouble(3) = oPos2.x
    oPointsDouble(4) = oPos2.y
    oPointsDouble(5) = oPos2.z
    
    oPointsDouble(6) = oPos3.x
    oPointsDouble(7) = oPos3.y
    oPointsDouble(8) = oPos3.z
    
    
    PointsToArray = oPointsDouble
    
Exit Function
ErrorHandler:
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateAvgOfCornersPlane
' This method gets the two edges in the direction of template and use them to construct base plane
' If the edges are three, it uses the end points of one edge in the direction and the vertex
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateAvgOfCornersPlane(ByVal oProfilePart As IJProfilePart, ByVal strDirection As String, ByVal oSurfaceBody As IJSurfaceBody) As Plane3d
Const METHOD = "CreateAvgOfCornersPlane"
    On Error GoTo ErrorHandler
    
    'prepare collection of edges from actual surface
    Dim oPlateEdges As IJElements

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oPlateEdges = oGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, False)
    On Error GoTo ErrorHandler
    
    If oPlateEdges Is Nothing Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    If oPlateEdges.count = 0 Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    ' If plateEgdes is passed in, find the average plane from all corner points
    Dim oTempEdges As IJElements
                            
    Dim oMfgTemplateHelper As MfgTemplateHelper
    Set oMfgTemplateHelper = New MfgTemplateHelper
    Dim dAvgRootX As Double
    Dim dAvgRootY As Double
    Dim dAvgRootZ As Double
    Dim dAvgNormalX As Double
    Dim dAvgNormalY As Double
    Dim dAvgNormalZ As Double
    Dim nIndex As Long
    Dim oCornerPoints(3) As IJDPosition
    Dim oCornerPointsDouble() As Double
    
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    
    If oPlateEdges.count < 3 Then 'this is legal, use natural balance
        'Get True Natural balance method
        Set CreateAvgOfCornersPlane = GetTrueNaturalBalance(oSurfaceBody)
    ElseIf oPlateEdges.count = 3 Then 'triangular
        GetEndPoints oPlateEdges(1), oCornerPoints(0), oCornerPoints(1)
        GetEndPoints oPlateEdges(2), oCornerPoints(2), oCornerPoints(3)
        
        'among these four points, two will be same
        'find that and construct plane with three distinct points
        Dim oPlane As IJPlane
        Set oPlane = New Plane3d
        '0 or 1 will coincide with 2 or 3 - check it.
        Dim oPoint1 As IJDPosition
        Dim oPoint2 As IJDPosition
        Dim oPoint3 As IJDPosition
        If (AreSamePoints(oCornerPoints(0), oCornerPoints(2))) Then
            'consider 0,1,3
            Set oPoint1 = oCornerPoints(0)
            Set oPoint2 = oCornerPoints(1)
            Set oPoint3 = oCornerPoints(3)
        ElseIf (AreSamePoints(oCornerPoints(0), oCornerPoints(3))) Then
            'consider 0,1,2
            Set oPoint1 = oCornerPoints(0)
            Set oPoint2 = oCornerPoints(1)
            Set oPoint3 = oCornerPoints(2)
        ElseIf (AreSamePoints(oCornerPoints(1), oCornerPoints(2))) Then
            'consider 0,1,3
            Set oPoint1 = oCornerPoints(0)
            Set oPoint2 = oCornerPoints(1)
            Set oPoint3 = oCornerPoints(3)
        Else '1 and 3 are same
            'consider 0,1,2
            Set oPoint1 = oCornerPoints(0)
            Set oPoint2 = oCornerPoints(1)
            Set oPoint3 = oCornerPoints(2)
        End If
        
        'construct plane
        oCornerPointsDouble() = PointsToArray(oPoint1, oPoint2, oPoint3)
        oPlane.DefineByPoints 3, oCornerPointsDouble
        
        'get normal
        oPlane.GetNormal dAvgNormalX, dAvgNormalY, dAvgNormalZ
        
        'calculate avg root point
        dAvgRootX = (oPoint1.x + oPoint2.x + oPoint3.x) / 3#
        dAvgRootY = (oPoint1.y + oPoint2.y + oPoint3.y) / 3#
        dAvgRootZ = (oPoint1.z + oPoint2.z + oPoint3.z) / 3#
            
        oBasePlane.DefineByPointNormal dAvgRootX, dAvgRootY, dAvgRootZ, dAvgNormalX, dAvgNormalY, dAvgNormalZ
            
        Set CreateAvgOfCornersPlane = oBasePlane
        
    Else 'four or more edges
        On Error Resume Next
        Set oTempEdges = GetTheTwoEdgesParallelToProfileAxis(oSurfaceBody, oPlateEdges)
        On Error GoTo ErrorHandler
        
        If oTempEdges Is Nothing Then
            Set CreateAvgOfCornersPlane = Nothing
            Exit Function
        End If
        
        If oTempEdges.count < 2 Then
            Set CreateAvgOfCornersPlane = Nothing
            Exit Function
        End If
        
        Dim oPlanes(3) As IJPlane
        
        GetEndPoints oTempEdges(1), oCornerPoints(0), oCornerPoints(1)
        GetEndPoints oTempEdges(2), oCornerPoints(2), oCornerPoints(3)
    
        Set oPlanes(0) = New Plane3d
        oCornerPointsDouble() = PointsToArray(oCornerPoints(0), oCornerPoints(1), oCornerPoints(2))
        oPlanes(0).DefineByPoints 3, oCornerPointsDouble
    
        Set oPlanes(1) = New Plane3d
        oCornerPointsDouble() = PointsToArray(oCornerPoints(1), oCornerPoints(2), oCornerPoints(3))
        oPlanes(1).DefineByPoints 3, oCornerPointsDouble
    
        Set oPlanes(2) = New Plane3d
        oCornerPointsDouble() = PointsToArray(oCornerPoints(2), oCornerPoints(3), oCornerPoints(0))
        oPlanes(2).DefineByPoints 3, oCornerPointsDouble
    
        Set oPlanes(3) = New Plane3d
        oCornerPointsDouble() = PointsToArray(oCornerPoints(3), oCornerPoints(0), oCornerPoints(1))
        oPlanes(3).DefineByPoints 3, oCornerPointsDouble
    
        For nIndex = 0 To 3
            Dim x As Double, y As Double, z As Double
    
            oCornerPoints(nIndex).Get x, y, z
    
            dAvgRootX = dAvgRootX + x
            dAvgRootY = dAvgRootY + y
            dAvgRootZ = dAvgRootZ + z
    
            oPlanes(nIndex).GetNormal x, y, z
    
            dAvgNormalX = dAvgNormalX + x
            dAvgNormalY = dAvgNormalY + y
            dAvgNormalZ = dAvgNormalZ + z
    
        Next
    
        dAvgRootX = dAvgRootX / 4#
        dAvgRootY = dAvgRootY / 4#
        dAvgRootZ = dAvgRootZ / 4#
        dAvgNormalX = dAvgNormalX / 4#
        dAvgNormalY = dAvgNormalY / 4#
        dAvgNormalZ = dAvgNormalZ / 4#
    
        oBasePlane.DefineByPointNormal dAvgRootX, dAvgRootY, dAvgRootZ, dAvgNormalX, dAvgNormalY, dAvgNormalZ
    
        Set CreateAvgOfCornersPlane = oBasePlane
    
        Set oTempEdges = Nothing
    
        For nIndex = 0 To 3
            Set oPlanes(nIndex) = Nothing
            Set oCornerPoints(nIndex) = Nothing
        Next
    End If
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'checks whether given two points are same
Private Function AreSamePoints(oPoint1 As IJDPosition, oPoint2 As IJDPosition) As Boolean
Const METHOD = "AreSamePoints"
On Error GoTo ErrorHandler

    If (Abs(oPoint1.x - oPoint2.x) <= DISTANCETOLERANCE And Abs(oPoint1.y - oPoint2.y) <= DISTANCETOLERANCE And Abs(oPoint1.z - oPoint2.z) <= DISTANCETOLERANCE) Then
        AreSamePoints = True
    End If
    
Exit Function
ErrorHandler:
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateBPlaneOfPerpendicularXY
'   1. Project the shell plate on the XYPlane
'   2. Calculate the mid point of each butt line
'   3. Calcuate the cross section lines between the 3D shell plate at Middle Point
'   4. Calculate the cross point between the middle cross section line and upper/lower seam line
'   5. Template base plane is perpendicular to Template Plane include the points calculated in step 4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateBPlaneOfPerpendicularXY(ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oProfilePart As IJProfilePart, ByVal oSurfaceBody As IJSurfaceBody, ByVal bBaseSide As Boolean) As IJPlane
    Const METHOD = "CreateBPlaneOfPerpendicularXY"
    On Error GoTo ErrorHandler

    'prepare collection of edges from actual surface
    Dim oEdges As IJElements
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oEdges = oGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, bBaseSide)
    On Error GoTo ErrorHandler
    
    If oEdges Is Nothing Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    If oEdges.count = 0 Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1. Project the shell plate on the XYPlane
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oProjectedEdges As IJElements
    Set oProjectedEdges = New JObjectCollection
       
    Dim oXYPlane As IJPlane
    oGeomHelper.MakeTransientPlane 0, 0, 0, 0, 0, 1, oXYPlane
        
    Dim nCount As Long
    For nCount = 1 To oEdges.count
        Dim oTempProjectedEdge As IUnknown
        Set oTempProjectedEdge = GetProjectedCurveOnPlane(oEdges.Item(nCount), oXYPlane)
        oProjectedEdges.Add oTempProjectedEdge
        Set oTempProjectedEdge = Nothing
    Next nCount
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2. Calculate the mid point of each butt line
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oMidPoint1 As IJDPosition, oMidPoint2 As IJDPosition
        
    Dim strTemplateDirection As String
    strTemplateDirection = oProcessSettings.TemplateDirection
    
    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    If oProjectedEdges.count < 3 Then
        'Get True Natural balance method
        GetMidPointsOfButtsSpecial oSurfaceBody, oProjectedEdges, strTemplateDirection, oMidPoint1, oMidPoint2
    ElseIf oProjectedEdges.count = 3 Then
        GetTriangleMidPointsOfButt oProjectedEdges, strTemplateDirection, oMidPoint1, oMidPoint2, oRootPoint
    Else 'plate of four or more edges
  
    '   Get MidPoints of Butts
        GetQuadRangleMidPointsOfButt oSurfaceBody, oProjectedEdges, strTemplateDirection, oMidPoint1, oMidPoint2
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   3. Calcuate the cross section lines between the 3D shell plate at Middle at Middle Position of two butts.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oLine As New Line3d
    oLine.DefineBy2Points oMidPoint1.x, oMidPoint1.y, oMidPoint1.z, oMidPoint2.x, oMidPoint2.y, oMidPoint2.z
'   Create the Template Plane at Center Points of above line
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    dRootX = (oMidPoint1.x + oMidPoint2.x) / 2
    dRootY = (oMidPoint1.y + oMidPoint2.y) / 2
    dRootZ = (oMidPoint1.z + oMidPoint2.z) / 2
    
    ' Create Template Plane
    Dim oTemplatePlane As IJPlane
    Set oTemplatePlane = New Plane3d
    
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oLine.GetDirection dNormalX, dNormalY, dNormalZ
    oTemplatePlane.SetRootPoint dRootX, dRootY, dRootZ
    oTemplatePlane.SetNormal dNormalX, dNormalY, dNormalZ
              
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   4. Calculate the cross point between the middle cross section line and upper/lower seam line
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oOutPutCurve As Object
    
    ' GetInterSectPoints
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oTemplatePlane, oOutPutCurve, oStartPos, oEndPos
    On Error GoTo ErrorHandler

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   5. Template base plane is perpendicular to Template Plane include the points calculated in step 4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not oOutPutCurve Is Nothing Then
        Dim oUVector As New DVector
        oUVector.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
        oUVector.Length = 1
        
        Dim oVVector As New DVector
        oVVector.Set dNormalX, dNormalY, dNormalZ
        oVVector.Length = 1
        
        Dim oNormalVec As New DVector
        Set oNormalVec = oUVector.Cross(oVVector)
    
        Dim oBasePlane As IJPlane
        Set oBasePlane = New Plane3d
        
        oBasePlane.DefineByPointNormal oStartPos.x, oStartPos.y, oStartPos.z, oNormalVec.x, oNormalVec.y, oNormalVec.z
        
        Set CreateBPlaneOfPerpendicularXY = oBasePlane
    Else
        Set CreateBPlaneOfPerpendicularXY = Nothing
    End If
    
    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'checks whether given two points are same
Public Function GetTrueNaturalBalance(ByVal oSurfaceBody As Object) As IJPlane
Const METHOD = "GetTrueNaturalBalance"
On Error GoTo ErrorHandler

    Dim oGeomUtils          As IJTopologyLocate
    Set oGeomUtils = New TopologyLocate
    
    Dim oCenter As IJDPosition
    Dim oNormal As IJDPosition
   'calculate base plane
    oGeomUtils.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
    
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    
    oBasePlane.DefineByPointNormal oCenter.x, oCenter.y, oCenter.z, oNormal.x, oNormal.y, oNormal.z
    
    Set GetTrueNaturalBalance = oBasePlane
        
    Exit Function
ErrorHandler:
End Function
 
Public Function CreateParallelAxisPlane(ByVal oProfilePart As IJProfilePart, ByVal oSurfaceBody As IJSurfaceBody) As Plane3d
Const METHOD = "CreateParallelAxisPlane"
    On Error GoTo ErrorHandler
    
     Dim oCenter             As IJDPosition
     Dim oNormal             As IJDVector
     Dim oParallelAxis        As IJDVector
     Dim oTopoLocate         As IJTopologyLocate
     
     Set oTopoLocate = New TopologyLocate
     oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
     Set oTopoLocate = Nothing
    
     Dim oXVector As IJDVector
     Dim oYVector As IJDVector
     Dim oZVector As IJDVector
     
     Set oXVector = New DVector
     Set oYVector = New DVector
     Set oZVector = New DVector
     
     oXVector.Set 1, 0, 0
     oYVector.Set 0, 1, 0
     oZVector.Set 0, 0, 1
     
     'normalize the vectors
     oXVector.Length = 1#
     oYVector.Length = 1#
     oZVector.Length = 1#
     
     Dim dX As Double
     Dim dY As Double
     Dim dZ As Double
     
     dX = Abs(oXVector.Dot(oNormal))
     dY = Abs(oYVector.Dot(oNormal))
     dZ = Abs(oZVector.Dot(oNormal))
     
     If (dX > dY) And (dX > dZ) Then
         Set oParallelAxis = oXVector
     ElseIf (dY > dX) And (dY > dZ) Then
         Set oParallelAxis = oYVector
     ElseIf (dZ > dX) And (dZ > dY) Then
         Set oParallelAxis = oZVector
     End If
    
     Dim oBasePlane As IJPlane
     Set oBasePlane = New Plane3d
     oBasePlane.DefineByPointNormal oCenter.x, oCenter.y, oCenter.z, oParallelAxis.x, oParallelAxis.y, oParallelAxis.z
    
     Set CreateParallelAxisPlane = oBasePlane

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetBasePlaneFromCorners()
'   1. Prepare collection of edges from actual surface.
'   2. Construct min box for the surface in standard x, y, z directions.
'   3. Depending on the center of plate part, decide if it is on Port side or Starboard side.
'   4. Store the three points amongst the min box (bottom surface) points which make sense for
'       LowerForeCorners, UpperForeCorners, UpperAftCorners, LowerAftCorners.
'   3. Get all the 3D vertices of the surface body into a collection.
'   4. Get a 2D vertices collection by projecting all the 3D vertices onto the bottom surface of the min box.
'   5. Find the nearest 2d vertices from the previously stored box points.
'   6. Get back the corresponding 3D vertices from the 2D vertices.
'   7. Create a template base plane from these 3D vertices.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetBasePlaneFromCorners(ByVal oProfilePart As Object, strBasePlane As String, ByVal oSurfaceBody As IJSurfaceBody, ByVal bBaseSide As Boolean) As IJPlane
Const METHOD = "GetBasePlaneFromCorners"
On Error GoTo ErrorHandler

    Dim oMfgGeomHelper          As MfgGeomHelper
    Dim oPlateEdges             As IJElements
    Dim oVectorElems            As IJElements
    Dim oCenter                 As IJDPosition
    Dim oNormal                 As IJDVector
    Dim oTopoLocate             As IJTopologyLocate
    Dim oXVector                As IJDVector
    Dim oYVector                As IJDVector
    Dim oZVector                As IJDVector

    Dim oMinBoxPoints           As IJElements
    Dim Points(1 To 4)          As IJDPosition
    Dim oBoxRootPoint           As IJDPosition
    Dim oBoxPointColl           As Collection
    
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oTopoLocate = New TopologyLocate
    Set oVectorElems = New JObjectCollection
    Set oBoxPointColl = New Collection
    
    'Prepare collection of edges from actual surface
    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oPlateEdges = oMfgGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, bBaseSide)
    On Error GoTo ErrorHandler
    
    If oPlateEdges Is Nothing Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    If oPlateEdges.count = 0 Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
        
    'Get center point and normal vector of the surface
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
    Set oTopoLocate = Nothing
    
    'Create vectors of unit length
    Set oXVector = New DVector
    Set oYVector = New DVector
    Set oZVector = New DVector
    
    'Normalize the vectors
    oXVector.Set 1, 0, 0
    oYVector.Set 0, 1, 0
    oZVector.Set 0, 0, 1
    
    oXVector.Length = 1#
    oYVector.Length = 1#
    oZVector.Length = 1#

    oVectorElems.Add oXVector
    oVectorElems.Add oYVector
    oVectorElems.Add oZVector
  
    ' Construct min box for the surface in standard x, y, z directions
    Set oMinBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVectors(oPlateEdges, oVectorElems)

    Set Points(1) = oMinBoxPoints.Item(1)
    Set Points(2) = oMinBoxPoints.Item(2)
    Set Points(3) = oMinBoxPoints.Item(3)
    
    Set oBoxRootPoint = New DPosition
    oBoxRootPoint.x = (Points(1).x + Points(3).x) / 2
    oBoxRootPoint.y = (Points(1).y + Points(3).y) / 2
    oBoxRootPoint.z = (Points(1).z + Points(3).z) / 2
    
    Set Points(4) = New DPosition
    Points(4).x = (2 * oBoxRootPoint.x) - Points(2).x
    Points(4).y = (2 * oBoxRootPoint.y) - Points(2).y
    Points(4).z = (2 * oBoxRootPoint.z) - Points(2).z

    'Decide the three points amongst the min box (bottom surface) points which make sense for
    '   LowerForeCorners, UpperForeCorners, UpperAftCorners, LowerAftCorners
    If oCenter.y > 0 Then
     '''''''''''''''''''''''''''''''''''''''
    '' Port Side
    '''''''''''''''''''''''''''''''''''''''
        If strBasePlane = "Lower Fore Corners" Then
            oBoxPointColl.Add Points(1)
            oBoxPointColl.Add Points(2)
            oBoxPointColl.Add Points(3)
        ElseIf strBasePlane = "Upper Fore Corners" Then
            oBoxPointColl.Add Points(2)
            oBoxPointColl.Add Points(3)
            oBoxPointColl.Add Points(4)
        ElseIf strBasePlane = "Upper Aft Corners" Then
            oBoxPointColl.Add Points(3)
            oBoxPointColl.Add Points(4)
            oBoxPointColl.Add Points(1)
        ElseIf strBasePlane = "Lower Aft Corners" Then
            oBoxPointColl.Add Points(4)
            oBoxPointColl.Add Points(1)
            oBoxPointColl.Add Points(2)
        End If
                
    Else
    '''''''''''''''''''''''''''''''''''''''
    '' Starboard Side
    '''''''''''''''''''''''''''''''''''''''
        If strBasePlane = "Lower Fore Corners" Then
            oBoxPointColl.Add Points(2)
            oBoxPointColl.Add Points(3)
            oBoxPointColl.Add Points(4)
        ElseIf strBasePlane = "Upper Fore Corners" Then
            oBoxPointColl.Add Points(1)
            oBoxPointColl.Add Points(2)
            oBoxPointColl.Add Points(3)
        ElseIf strBasePlane = "Upper Aft Corners" Then
            oBoxPointColl.Add Points(4)
            oBoxPointColl.Add Points(1)
            oBoxPointColl.Add Points(2)
        ElseIf strBasePlane = "Lower Aft Corners" Then
            oBoxPointColl.Add Points(3)
            oBoxPointColl.Add Points(4)
            oBoxPointColl.Add Points(1)
        End If
            
    End If

    Dim o3DVertices                 As Collection
    Dim o2DVertices                 As Collection
    Dim o2DNearestVertices          As Collection
    Dim i                           As Integer
    Dim j                           As Integer
    Dim o3DPoint                    As IJDPosition
    Dim o2DPoint                    As IJDPosition
    Dim o3DFinalPoint(1 To 3)       As IJDPosition
    Dim oSGOModelBodyUtilities      As SGOModelBodyUtilities
    
    'Get all the 3D vertices of the surface body into a collection.
    Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
    oSGOModelBodyUtilities.GetVertices oSurfaceBody, o3DVertices
    
    Set o2DVertices = New Collection
    Set o2DNearestVertices = New Collection
    
    'Get a 2D vertices collection by projecting all the 3D vertices onto the bottom surface of the min box.
    For i = 1 To o3DVertices.count
        Set o2DPoint = New DPosition
        Set o3DPoint = o3DVertices.Item(i)
        o2DPoint.x = o3DPoint.x
        o2DPoint.y = o3DPoint.y
        o2DPoint.z = Points(1).z
        
        o2DVertices.Add o2DPoint
    Next
    
    'Find the nearest 2d vertex from the previously stored box points.
    Set o2DNearestVertices = GetNearestVerticesFromBoxPoints(oBoxPointColl, o2DVertices)
    
    'Get back the corresponding 3D vertices from the 2D vertices.
    For i = 1 To o2DNearestVertices.count
        Set o2DPoint = o2DNearestVertices.Item(i)
    
        For j = 1 To o3DVertices.count
            Set o3DPoint = o3DVertices.Item(j)
            If (o2DPoint.x = o3DPoint.x) And (o2DPoint.y = o3DPoint.y) Then
                Set o3DFinalPoint(i) = o3DPoint
                o3DVertices.Remove j
                Exit For
            End If
    
        Next j
    Next i
    
    Dim o3PointsPlane           As IJPlane
    Dim oCornerPointsDouble()       As Double
    Dim dAvgRootX                   As Double
    Dim dAvgRootY                   As Double
    Dim dAvgRootZ                   As Double
    Dim dAvgNormalX                 As Double
    Dim dAvgNormalY                 As Double
    Dim dAvgNormalZ                 As Double
    
    'construct plane
    oCornerPointsDouble() = PointsToArray(o3DFinalPoint(1), o3DFinalPoint(2), o3DFinalPoint(3))
        
    'Create a template base plane from these 3D vertices.
    Set o3PointsPlane = New Plane3d
    o3PointsPlane.DefineByPoints 3, oCornerPointsDouble
    
    'get normal
    o3PointsPlane.GetNormal dAvgNormalX, dAvgNormalY, dAvgNormalZ
    
    'calculate avg root point
    dAvgRootX = (o3DFinalPoint(1).x + o3DFinalPoint(2).x + o3DFinalPoint(3).x) / 3#
    dAvgRootY = (o3DFinalPoint(1).y + o3DFinalPoint(2).y + o3DFinalPoint(3).y) / 3#
    dAvgRootZ = (o3DFinalPoint(1).z + o3DFinalPoint(2).z + o3DFinalPoint(3).z) / 3#
        
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    
    oBasePlane.DefineByPointNormal dAvgRootX, dAvgRootY, dAvgRootZ, dAvgNormalX, dAvgNormalY, dAvgNormalZ
    
    Set GetBasePlaneFromCorners = oBasePlane
    
CleanUp:
    Erase Points
    Set oBoxPointColl = Nothing
    Set oMinBoxPoints = Nothing
    Set oVectorElems = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function GetNearestVerticesFromBoxPoints(ByVal oBoxPointColl As Collection, ByVal oVertices As Collection) As Collection
Const METHOD = "GetNearestVerticesFromBoxPoints"
On Error GoTo ErrorHandler

    Dim i                           As Integer
    Dim j                           As Integer
    Dim oBoxPoint                   As IJDPosition
    Dim oVertexPoint                As IJDPosition
    Dim tmpPoint                    As IJDPosition
    Dim dDistance                   As Double
    Dim dMinDistance                As Double
    Dim oNearestVextexColl          As Collection
    Set oNearestVextexColl = New Collection
    
    j = 1
    
NextBoxPoint:
    
    Set oBoxPoint = oBoxPointColl.Item(j)
    Set tmpPoint = oVertices.Item(1)
    Set oVertexPoint = tmpPoint

    dMinDistance = oBoxPoint.DistPt(tmpPoint)
    Set tmpPoint = Nothing

    For i = 2 To oVertices.count
        Set tmpPoint = oVertices.Item(i)
        dDistance = oBoxPoint.DistPt(tmpPoint)

        If dDistance < dMinDistance Then
            dMinDistance = dDistance
            Set oVertexPoint = tmpPoint
        End If
    Next

    oNearestVextexColl.Add oVertexPoint
    
    For i = 1 To oVertices.count
        If oVertexPoint Is oVertices.Item(i) Then
            oVertices.Remove (i)
            Exit For
        End If
    Next
    
    Set oVertexPoint = Nothing
    
    If j < oBoxPointColl.count Then
        j = j + 1
        GoTo NextBoxPoint
    End If
    
    Set GetNearestVerticesFromBoxPoints = oNearestVextexColl

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetParallelAxis(ByVal oSurfaceBody As IJSurfaceBody) As IJDVector
Const METHOD = "GetParallelAxis"
    On Error GoTo ErrorHandler
    
     Dim oCenter             As IJDPosition
     Dim oNormal             As IJDVector
     Dim oFinalNormal        As IJDVector
     Dim oTopoLocate         As IJTopologyLocate
     
     Set oTopoLocate = New TopologyLocate
     oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
     Set oTopoLocate = Nothing
    
     Dim oXVector As IJDVector
     Dim oYVector As IJDVector
     Dim oZVector As IJDVector
     
     Set oXVector = New DVector
     Set oYVector = New DVector
     Set oZVector = New DVector
     
     oXVector.Set 1, 0, 0
     oYVector.Set 0, 1, 0
     oZVector.Set 0, 0, 1
     
     'normalize the vectors
     oXVector.Length = 1#
     oYVector.Length = 1#
     oZVector.Length = 1#
     
     Dim dX As Double
     Dim dY As Double
     Dim dZ As Double
     
     dX = Abs(oXVector.Dot(oNormal))
     dY = Abs(oYVector.Dot(oNormal))
     dZ = Abs(oZVector.Dot(oNormal))
     
     If (dX > dY) And (dX > dZ) Then
         Set GetParallelAxis = oXVector
     ElseIf (dY > dX) And (dY > dZ) Then
         Set GetParallelAxis = oYVector
     ElseIf (dZ > dX) And (dZ > dY) Then
         Set GetParallelAxis = oZVector
     End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'This method gets the port surface based on template side supplied
Public Function GetPortSurface(oProfilePart As Object, ByVal strTemplateSide As String) As IJSurfaceBody
    Const METHOD = "GetPortSurface"
    On Error GoTo ErrorHandler
    
    Dim oConnectable As IJConnectable
    Dim oEnumPorts As IJElements
    Dim oStructPort As IJStructPort
    Dim oPort As IJPort
   
   If TypeOf oProfilePart Is IJConnectable Then
     
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New ProfilePartSupport
        Set oPartSupport.Part = oProfilePart
        Dim oLeafPlateSystem As IJSystem
        oPartSupport.IsSystemDerivedPart oLeafPlateSystem, False
              
        ' Set the connectable.
        Set oConnectable = oLeafPlateSystem
       
        ' Set the correct context.
        Dim eContextID As eUSER_CTX_FLAGS
      
        If strTemplateSide = "BaseSide" Then
            eContextID = CTX_NMINUS
        ElseIf strTemplateSide = "OffsetSide" Then
            eContextID = CTX_NPLUS
        Else
            Dim pTopoLocate As IJTopologyLocate
            Set pTopoLocate = New TopologyLocate
            
            Dim pNormalVec As IJDVector
            Dim pThicknessDirection As IJDVector
            Dim pPointOnSurface As IJDPosition
            
            pTopoLocate.FindApproxCenterAndNormal oLeafPlateSystem, pPointOnSurface, pNormalVec
            
            Dim pPlateUtils As IJPlateAttributes
            Set pPlateUtils = New PlateUtils
            
            pPlateUtils.GetPlateThicknessDirVec oLeafPlateSystem, pPointOnSurface, pThicknessDirection
            
            If strTemplateSide = "MoldedSide" Then
                If pThicknessDirection.Dot(pNormalVec) > 0 Then
                    eContextID = CTX_NMINUS
                Else
                    eContextID = CTX_NPLUS
                End If
            ElseIf strTemplateSide = "AntiMoldedSide" Then
                If pThicknessDirection.Dot(pNormalVec) > 0 Then
                    eContextID = CTX_NPLUS
                Else
                    eContextID = CTX_NMINUS
                End If
            End If
        End If
        
        Dim ePortType As PortType
        ePortType = PortFace
        
        ' Get its edge ports.
        oConnectable.enumPorts oEnumPorts, ePortType
        
        Dim i As Integer
        For i = 1 To oEnumPorts.count
            Dim varElm As Variant
            Set varElm = oEnumPorts.Item(i)
            If TypeOf varElm Is IJStructPort Then
         
                Set oStructPort = Nothing
                Set oStructPort = varElm
                
                ' Gather if correct side.
                If oStructPort.ContextID = eContextID Then
                    ' Add to collection of possible.
                    If TypeOf oStructPort Is IJPort Then
                        Set oPort = oStructPort
                        ' There is only one port that should be found here.
                        Exit For
                    End If
                End If
            End If
        Next
        
        If Not oPort Is Nothing Then
           ' Assign the port found.
           Set GetPortSurface = oPort.Geometry
        Else
           Set GetPortSurface = Nothing
        End If
        
    End If
   
CleanUp:
    Set oConnectable = Nothing
    Set oEnumPorts = Nothing
    Set oStructPort = Nothing
    Set oPort = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


