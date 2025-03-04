Attribute VB_Name = "ControlLineHelpers"
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Control Line Helpers Bas Module
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************


Option Explicit

Private Const MODULE = "GSCADStrMfgTemplate.ControlLineService"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateBCtlLineOfNormalPlate
'   1. Get the centre of the Profile Surface
'   2. Get the SightPlane at the Cntre base don the Direction
'       2.1. If Direction is Perp To Axis,Get SightPlane Along Axis
'       2.2  If Direction is Along Axis, Get SightPlane Perp To Axis
'   3. Get Control Line calculated by intersecting between Plate Part and SightPlane
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateBCtlLineOfNormalProfile(ByVal oProfilePart As IJProfilePart, ByVal strDirection As String, ByVal oSurfaceBody As IJSurfaceBody, ByVal pBasePlane As Plane3d) As IUnknown
    Const METHOD = "CreateBCtlLineOfNormalProfile"
    On Error GoTo ErrorHandler
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  1. Create SightPlane
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'get the centre
    Dim oCentre As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCentre, oNormal
    Set oTopoLocate = Nothing
    
    'Get the longest edge vector (Alogn Axis Vector)
    
    ''prepare collection of edges from actual surface
    Dim oEdgeElements As IJElements
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oEdgeElements = oGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, False)
    On Error GoTo ErrorHandler
    
    If oEdgeElements Is Nothing Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    If oEdgeElements.count = 0 Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If

    Dim oLongestEdges As IJElements
    Set oLongestEdges = GetTheTwoEdgesParallelToProfileAxis(oSurfaceBody, oEdgeElements)
    
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    GetEndPoints oLongestEdges.Item(1), oStartPos, oEndPos
    
    Dim oAlongAxisVector As IJDVector
    Set oAlongAxisVector = oEndPos.Subtract(oStartPos)
       
    'Get the baseplane normal
    Dim oBasePlaneVec As DVector
    Set oBasePlaneVec = New DVector '
    Dim dX As Double, dY As Double, dZ As Double
    pBasePlane.GetNormal dX, dY, dZ
    oBasePlaneVec.Set dX, dY, dZ
    oBasePlaneVec.Length = 1
    
    'Get the cross product of above two - this is PerpToAxis Vector
    Dim oPerpToAxisVec As IJDVector
    oAlongAxisVector.Length = 1
    Set oPerpToAxisVec = oBasePlaneVec.Cross(oAlongAxisVector)
        
    Dim oSightPlane As IJPlane
         
    If strDirection = "PerpToAxis" Then  'if Direction is PerpToAxis
        'construct plane at centre passing thorugh above two
        oGeomHelper.MakeTransientPlane oCentre.x, oCentre.y, oCentre.z, oPerpToAxisVec.x, oPerpToAxisVec.y, oPerpToAxisVec.z, oSightPlane
    End If
    
    If strDirection = "AlongAxis" Then  'if Direction is AlongAxis
        'get cross-product of PerpToAxis and BasePlane Normal
        Dim oAlongAxisVector2 As IJDVector
        
        Set oAlongAxisVector2 = oBasePlaneVec.Cross(oPerpToAxisVec)
        'construct plane at centre with above result
        oGeomHelper.MakeTransientPlane oCentre.x, oCentre.y, oCentre.z, oAlongAxisVector2.x, oAlongAxisVector2.y, oAlongAxisVector2.z, oSightPlane
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2. Get Control Line calculated by intersecting between Plate Part and SightPlane
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
    Dim oControlLine As IUnknown
        
    Dim oStartPosition As IJDPosition, oEndPosition As IJDPosition
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oSightPlane, oControlLine, oStartPosition, oEndPosition
    On Error GoTo ErrorHandler
    
    Dim oCS As IJComplexString
    If Not oControlLine Is Nothing Then 'make sure it is CS
        Dim oMfgMGHelper As IJMfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper
        
        'if the intersection of oSurfaceBody, oSightPlane results in ambuigity or a wirbody of multiple lumps ( for plates with opening)
        'the WireBodyToComplexString would return nothing. In such cases, get the ComplexStrings from the wirebody and assign one of them as BCL
        On Error Resume Next
        oMfgMGHelper.WireBodyToComplexString oControlLine, oCS
        On Error GoTo ErrorHandler
                
        If oCS Is Nothing Then
            Dim oCSElems As IJElements
            oMfgMGHelper.WireBodyToComplexStrings oControlLine, oCSElems
            Set oCS = oCSElems.Item(1)
        End If
    End If
    
    Set CreateBCtlLineOfNormalProfile = oCS
    
    Set oSurfaceBody = Nothing
    
    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CreateBCtlLineOfTriangle
'   1. Get Vertex across from the Edge which is parallel to the Frame
'   2. Get Center Reference Frame and Get Intersection Points
'   3. Create YZ Plane at Step2's Point
'   4. Get Intersection Points and Curve between the YZ Palne and BasePlane
'   5. Create SightPlane using Step4's point and Curve and Get Control Line
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateBCtlLineOfTriangle(oProcessSettings As IJMfgTemplateProcessSettings, oEdgeElements As IJElements, oSurfaceBody As IJSurfaceBody, pBasePlane As Plane3d, oEndPtColl As Collection)
    Const METHOD = "CreateBCtlLineOfTriangle"
    On Error GoTo ErrorHandler
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1. Get Vertex across from the Edge which is parallel to the Frame
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Check the Template Orientation'
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    
    Dim oVertex As IJDPosition
    Set oVertex = New DPosition
    Dim nCount  As Long
    
    For nCount = 1 To oEdgeElements.count
        GetEndPoints oEdgeElements.Item(nCount), oStartPos, oEndPos
        If Not IsEqualPoint(oStartPos, oEndPtColl.Item(3)) Then
            If Not IsEqualPoint(oStartPos, oEndPtColl.Item(4)) Then
                oVertex.Set oStartPos.x, oStartPos.y, oStartPos.z
                Exit For
            End If
        End If
        
        If Not IsEqualPoint(oEndPos, oEndPtColl.Item(3)) Then
            If Not IsEqualPoint(oEndPos, oEndPtColl.Item(4)) Then
                oVertex.Set oEndPos.x, oEndPos.y, oEndPos.z
                Exit For
            End If
        End If
        Set oStartPos = Nothing
        
    Next nCount
          
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2. Get Center Reference Frame and Get Intersection Points
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    GetCenterFrameFromPlatePart oSurfaceBody, oProcessSettings, oStartPos, oEndPos
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   3. Create Plane at Step2's Point
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strDirection As String
    Dim oTempPlane As Plane3d
    strDirection = oProcessSettings.TemplateDirection
    
     If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
         oGeomHelper.MakeTransientPlane oStartPos.x, oStartPos.y, oStartPos.z, 0, 1, 0, oTempPlane
     ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
         oGeomHelper.MakeTransientPlane oStartPos.x, oStartPos.y, oStartPos.z, 1, 0, 0, oTempPlane
     Else 'Z - Direction(WaterLine)
         oGeomHelper.MakeTransientPlane oStartPos.x, oStartPos.y, oStartPos.z, 0, 0, 1, oTempPlane
     End If
     
'    Dim oYZPlane As Plane3d
'    oGeomHelper.MakeTransientPlane oStartPos.x, oStartPos.y, oStartPos.z, 1, 0, 0, oYZPlane
      
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Dim oOutPutCurve As IUnknown
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   4. Get Intersection Points and Curve between the Temp Palne and BasePlane
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oTempPlane, pBasePlane, oOutPutCurve, oStartPos, oEndPos
    On Error GoTo ErrorHandler
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   5. Create SightPlane using Step4's point and Curve and Get Control Line
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oSightPlane As IJPlane
    oGeomHelper.MakeTransientPlane oVertex.x, oVertex.y, oVertex.z, oEndPos.x - oStartPos.x, _
                                                            oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z, oSightPlane
    
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    
    Dim oControlLine As IUnknown
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oSightPlane, oControlLine, oStartPos, oEndPos
    On Error GoTo ErrorHandler
    
    Set CreateBCtlLineOfTriangle = oControlLine
   

    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'********************************************************************************************************************************
' Function Name:    CreateBCtlLineOfTrainglePlateNew
'
' Abstract:         This method computes the base control line for traingle plate. It proposes
'                   the line joining midpoint of an edge(along the Direction)of the plate and the opposite vertex as BCL
'
' Inputs:           strDirection -- Template Direction
'                   oEdgeElements -- Edges of the traingle plate
'                   oSurfaceBody -- Plate surface on the tempalte side
'                   pBasePlane -- BasePlane of the template
'
' Output:           The Base Control Line computed
'
' Algorithm:
'                   1. Get the Edge of the plate that is in the template direction
'                   2. Get mid point of the edge
'                   3. Get the vertex opposite to it
'                   4. Create SightPlane using BasePlane's Normal and  midpoint of the Edge and the vertex
'                   5. Get Control Line calculated by intersecting between Plate Part and SightPlane
'*********************************************************************************************************************************
Public Function CreateBCtlLineOfTrainglePlateNew(ByVal strDirection As String, ByVal oEdgeElements As IJElements, ByVal oSurfaceBody As IJSurfaceBody, ByVal pBasePlane As Plane3d) As IUnknown
    Const METHOD = "CreateBCtlLineOfTrainglePlateNew"
    On Error GoTo ErrorHandler
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   1. Get the Edge of the plate that is in the template direction
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim oControlLine As IUnknown
    Dim oSightPlane As IJPlane
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    
    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    Dim oBestEdge As Object
    Set oBestEdge = GetBestEdge(oEdgeElements, strDirection, oRootPoint)
    If oBestEdge Is Nothing Then Exit Function
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   2. Get mid point of the edge
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim oMidPoint1 As IJDPosition, oVertex As IJDPosition
    Set oMidPoint1 = New DPosition
    
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    GetEndPoints oBestEdge, oStartPos, oEndPos
    oMidPoint1.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   3. Get the vertex opposite to it
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set oVertex = GetVertexNotOnEdgeForTrainglePlate(oEdgeElements, oBestEdge)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   4. Create SightPlane using BasePlane's Normal and  midpoint of the Edge and the vertex
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oUVec As DVector
    Set oUVec = New DVector '
    Dim dUX As Double, dUY As Double, dUZ As Double
    pBasePlane.GetNormal dUX, dUY, dUZ
    oUVec.Set dUX, dUY, dUZ
    oUVec.Length = 1
    
    Dim oVVec As DVector
    Set oVVec = New DVector
    oVVec.Set oMidPoint1.x - oVertex.x, oMidPoint1.y - oVertex.y, oMidPoint1.z - oVertex.z
    oVVec.Length = 1
    
    Dim oNormalVec As DVector
    Set oNormalVec = oUVec.Cross(oVVec)
    
    oGeomHelper.MakeTransientPlane oVertex.x, oVertex.y, oVertex.z, oNormalVec.x, oNormalVec.y, oNormalVec.z, oSightPlane
 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   5. Get Control Line calculated by intersecting between Plate Part and SightPlane
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oStartPosition As IJDPosition, oEndPosition As IJDPosition
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oSightPlane, oControlLine, oStartPosition, oEndPosition
    On Error GoTo ErrorHandler
    
    Set CreateBCtlLineOfTrainglePlateNew = oControlLine
    Set oSurfaceBody = Nothing
    
    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateBCtlLineStemStern
' Control Line is the CenterLine i.e. Intersection Line between ProfilePart and y=0 plane
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateBCtlLineStemStern(ByVal oSurfaceBody As IJSurfaceBody) As IUnknown
    Const METHOD = "CreateBCtlLineStemStern"
    On Error GoTo ErrorHandler
    
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim oXZPlane As Plane3d 'Y=0 plane
    oGeomHelper.MakeTransientPlane 0, 0, 0, 0, 1, 0, oXZPlane
    
    Dim oOutPutCurve As IUnknown
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oXZPlane, oOutPutCurve, oStartPos, oEndPos
    On Error GoTo ErrorHandler
    
    Dim oCS As IJComplexString
    If Not oOutPutCurve Is Nothing Then
        Dim oMfgMGHelper As IJMfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper
        oMfgMGHelper.WireBodyToComplexString oOutPutCurve, oCS
    End If
    
    Set CreateBCtlLineStemStern = oCS
    
CleanUp:
    Set oXZPlane = Nothing
    Set oOutPutCurve = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateBCtlLineOfPerpendicularXY
'   1. Project the Edge on to XY Plane and Get Projected Edges
'   2. Calculate the mid point of each butt line
'   3. Create TemplatePlane at Middle Point and calcuate the Intersection Points between the 3D shell plate
'   4. Create SightPlane as Middle Point and Get Intersection Line between this SightPlane and ProfilePart
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateBCtlLineOfPerpendicularXY(ByVal oProfilePart As IJProfilePart, ByVal strTemplateDirection As String, ByVal oSurfaceBody As IJSurfaceBody, ByVal pBasePlane As Plane3d, ByVal bBaseSide As Boolean) As IUnknown
    Const METHOD = "CreateBCtlLineOfPerpendicularXY"
    On Error GoTo ErrorHandler
    
    'prepare collection of edges from actual surface
    Dim oEdgeElements As IJElements
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oEdgeElements = oGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, bBaseSide)
    On Error GoTo ErrorHandler
    
    If oEdgeElements Is Nothing Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If
    
    If oEdgeElements.count = 0 Then
        Call StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToGetEgesFromPlatePart, , "RULES")
        Exit Function
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1. Project the Edge on to XY Plane and Get Projected Edges
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oProjectedEdges As IJElements
    Set oProjectedEdges = New JObjectCollection
    
       
    Dim oXYPlane As IJPlane
    oGeomHelper.MakeTransientPlane 0, 0, 0, 0, 0, 1, oXYPlane
        
    Dim nCount As Long
    For nCount = 1 To oEdgeElements.count
        Dim oTempProjectedEdge As IUnknown
        Set oTempProjectedEdge = GetProjectedCurveOnPlane(oEdgeElements.Item(nCount), oXYPlane)
                   
        oProjectedEdges.Add oTempProjectedEdge
        Set oTempProjectedEdge = Nothing
    Next nCount

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2. Calculate the mid point of each butt line
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oMidPoint1 As IJDPosition, oMidPoint2 As IJDPosition
    Set oMidPoint1 = New DPosition
    Set oMidPoint2 = New DPosition
    
    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    
    If oProjectedEdges.count < 3 Then
        GetMidPointsOfButtsSpecialForBCL oProfilePart, oSurfaceBody, oProjectedEdges, strTemplateDirection, pBasePlane, oMidPoint1, oMidPoint2
    ElseIf oProjectedEdges.count = 3 Then
        GetTriangleMidPointsOfButt oProjectedEdges, strTemplateDirection, oMidPoint1, oMidPoint2, oRootPoint
    Else ' plate of four or more edges
    '   Get MidPoints of Butts
        GetQuadRangleMidPointsOfButt oSurfaceBody, oProjectedEdges, strTemplateDirection, oMidPoint1, oMidPoint2
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   3. Create TemplatePlane at Middle Point and calcuate the Intersection Points between the 3D shell plate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oLine As New Line3d
    oLine.DefineBy2Points oMidPoint1.x, oMidPoint1.y, oMidPoint1.z, oMidPoint2.x, oMidPoint2.y, oMidPoint2.z
    
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
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   5. Create SightPlane as Middle Point and Get Intersection Line between this SightPlane and ProfilePart
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not oOutPutCurve Is Nothing Then
        Dim oNormalVec As New DVector
        oNormalVec.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
        oNormalVec.Length = 1
        
        dRootX = (oEndPos.x + oStartPos.x) / 2
        dRootY = (oEndPos.y + oStartPos.y) / 2
        dRootZ = (oEndPos.z + oStartPos.z) / 2
        
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oOutPutCurve = Nothing
    End If
            
    Dim oSightPlane As IJPlane
    Set oSightPlane = New Plane3d
    
    oSightPlane.SetRootPoint dRootX, dRootY, dRootZ
    oSightPlane.SetNormal oNormalVec.x, oNormalVec.y, oNormalVec.z
          
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oSightPlane, oOutPutCurve, oStartPos, oEndPos
    On Error GoTo ErrorHandler
    
    Dim oCS As IJComplexString
    If Not oOutPutCurve Is Nothing Then
        Dim oMfgMGHelper As IJMfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper
        oMfgMGHelper.WireBodyToComplexString oOutPutCurve, oCS
    End If
    
    Set CreateBCtlLineOfPerpendicularXY = oCS
 
    Exit Function
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CreateControlLineAtAvgPointAvgNormal(ByVal oTemplateSetData As TemSetData, ByVal oProfilePart As Object, ByVal pProcessSettings As Object, ByVal pBasePlane As Object) As Object
    Const METHOD = "CreateControlLineAtAvgPointAvgNormal"
    On Error GoTo ErrorHandler
    
    'prepare collection of edges from actual surface
    Dim oEdgeElements As IJElements
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim bBaseSide As Boolean
    
    'get molded surface
    Dim ePlateThicknessSide As PlateThicknessSide
    If bBaseSide Then
        ePlateThicknessSide = PlateBaseSide
    Else
        ePlateThicknessSide = PlateOffsetSide
    End If
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oGeomHelper.GetSurfaceFromPlate(oProfilePart, TRUE_MOLD, ePlateThicknessSide, 0)

    On Error Resume Next 'expected that there can be errors when surface is bad
    Set oEdgeElements = oGeomHelper.GetPlatePartEdgesInIJElements(oSurfaceBody, False)
    On Error GoTo ErrorHandler
    
    'call the extended method
    Set CreateControlLineAtAvgPointAvgNormal = CreateControlLineAtAvgPointAvgNormal_Ex(oTemplateSetData, oProfilePart, pProcessSettings, pBasePlane, oEdgeElements)
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
    
Public Function CreateControlLineAtAvgPointAvgNormal_Ex(oTemplateSetData As TemSetData, pProfilePart As Object, pProcessSettings As Object, pBasePlane As Object, oEdgeElements As IJElements) As Object
    Const METHOD = "CreateControlLineAtAvgPointAvgNormal_Ex"
    On Error GoTo ErrorHandler
            
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       2. Check the Template Type and If it is a special plate or not
'           2.1. Template Type is PerPendicualrXY(Near to Stem/Stern) or
'               Pentagonal plate or Triangular plate which edge is not parallel to Frame line,
'           => Perpendicular XY Algorithm
'           2.2. Template Type is Stem/Stern
'           => Stem/Stern Type, There is no triangular and pentagonal plate(only quadrangle plate).
'           => refer to Analysis- DefineCtlLineAndBPlane.doc
'           => Exactly symmetrical to Center line
'           2.3. Other Case
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oSurfaceBody As IUnknown
    Set oSurfaceBody = oTemplateSetData.SurfaceBody

    Dim oControlLine As IUnknown
'    Dim ePlateType As enumPlateType

    Dim dRX As Double, dRY As Double, dRZ As Double
    Dim dNX As Double, dNY As Double, dNZ As Double
 
    oGeomHelper.GetPlatePartAvgPointAvgNormal oSurfaceBody, False, _
                    dRX, dRY, dRZ, dNX, dNY, dNZ
    
    Dim strDirection As String
    Dim oPlateNormal As DVector
    Dim oDirVec As DVector
    Dim oBCLPlaneNormal As DVector
    Dim oBCLPlane As Plane3d
    Dim oStartPosition As IJDPosition, oEndPosition As IJDPosition
    
    strDirection = oTemplateSetData.Direction

    Set oPlateNormal = New DVector
    oPlateNormal.Set dNX, dNY, dNZ
    oPlateNormal.Length = 1
    
    Set oDirVec = New DVector
    If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
        oDirVec.Set 0, 1, 0
    ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
        oDirVec.Set 1, 0, 0
    Else 'Z - Direction(WaterLine)
        oDirVec.Set 0, 0, 1
    End If
    oDirVec.Length = 1
    
    Set oBCLPlaneNormal = oDirVec.Cross(oPlateNormal)
    oBCLPlaneNormal.Get dNX, dNY, dNZ
    oBCLPlaneNormal.Length = 1
    
    Set oBCLPlane = New Plane3d
    oBCLPlane.DefineByPointNormal dRX, dRY, dRZ, dNX, dNY, dNZ
        
    If oEdgeElements Is Nothing Then 'directly intesect the surface with BCLPlane
        oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oControlLine, oStartPosition, oEndPosition
        Set CreateControlLineAtAvgPointAvgNormal_Ex = oControlLine
        Exit Function
    End If
    
    If oEdgeElements.count < 3 Then  'directly intesect the surface with BCLPlane
        oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oControlLine, oStartPosition, oEndPosition
        Set CreateControlLineAtAvgPointAvgNormal_Ex = oControlLine
        Exit Function
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' try to connect mid points of extreme edges
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim nIndex As Long
    Dim oEdges(2) As Object    ' we only need two edges
    Dim oEdge As IUnknown
    Dim nEdgesFound  As Long
    Dim oIntersection As Object
    nEdgesFound = 0
    For nIndex = 1 To oEdgeElements.count
        If nEdgesFound < 2 Then
            Set oEdge = oEdgeElements.Item(nIndex)
            Set oIntersection = oGeomHelper.IntersectCurveWithPlane(oEdge, oBCLPlane)
            If Not oIntersection Is Nothing Then
                nEdgesFound = nEdgesFound + 1
                Set oEdges(nEdgesFound) = oEdge
            End If
            Set oIntersection = Nothing
            Set oEdge = Nothing
        End If
    Next
    
    If nEdgesFound < 2 Then 'directly intesect the surface with BCLPlane
        oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oControlLine, oStartPosition, oEndPosition
        Set CreateControlLineAtAvgPointAvgNormal_Ex = oControlLine
        Exit Function
    End If
    
    'BCL will be along the line joining the mid points of these two found edges
    Dim oMidPoint1 As IJDPosition
    Dim oMidPoint2 As IJDPosition
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    
    Set oMidPoint1 = New DPosition
    Set oMidPoint2 = New DPosition
    
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    GetEndPoints oEdges(1), oStartPos, oEndPos
    oMidPoint1.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    GetEndPoints oEdges(2), oStartPos, oEndPos
    oMidPoint2.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    Set oStartPos = Nothing
    Set oEndPos = Nothing
       
    Dim oUVec As DVector
    Set oUVec = New DVector '
    Dim dUX As Double, dUY As Double, dUZ As Double
    Dim oBasePlane As IJPlane
    Set oBasePlane = pBasePlane
    oBasePlane.GetNormal dUX, dUY, dUZ
    Set oBasePlane = Nothing
    oUVec.Set dUX, dUY, dUZ
    oUVec.Length = 1
    
    Dim oVVec As DVector
    Set oVVec = New DVector
    oVVec.Set oMidPoint1.x - oMidPoint2.x, oMidPoint1.y - oMidPoint2.y, oMidPoint1.z - oMidPoint2.z
    oVVec.Length = 1
    
    Dim oNormalVec As DVector
    Set oNormalVec = oUVec.Cross(oVVec)
    Set oBCLPlane = Nothing
    oGeomHelper.MakeTransientPlane oMidPoint1.x, oMidPoint1.y, oMidPoint1.z, oNormalVec.x, oNormalVec.y, oNormalVec.z, oBCLPlane
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
    oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oControlLine, oStartPosition, oEndPosition
    

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   3.      Return Control Line
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set CreateControlLineAtAvgPointAvgNormal_Ex = oControlLine
    Set oTemplateSetData = Nothing
    Set oGeomHelper = Nothing
    Set oPlateNormal = Nothing
    Set oDirVec = Nothing
    Set oBCLPlaneNormal = Nothing
    Set oBCLPlane = Nothing
    Set oStartPosition = Nothing
    Set oEndPosition = Nothing
    Set oControlLine = Nothing

    Exit Function
    
ISWP_ErrorHandler:     'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FailedToIntersectSurfaceWithPlane, , "RULES")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


 
