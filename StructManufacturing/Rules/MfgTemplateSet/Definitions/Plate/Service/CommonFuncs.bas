Attribute VB_Name = "CommonFuncs"
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Bas file with functions to support Service
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new bas
'************************************************************************************************************

Option Explicit
Private Const MODULE = "CommonFuncs"
Private Const EPSILON = 0.001
Private Const EPSILON_PLANAR = 0.0005
Private Const FRAME_EDGE_DIST = 0.009 ' consider frame is on edge if dist < FRAME_EDGE_DIST
Private Const DISTANCETOLERANCE = 0.000001  'Distance Tolerance = 1E-06
Private Const COS45 = 0.7072 '1/sqrt(2)

Public Sub GetCenterFrameFromPlatePart(oSurfaceBody As IJSurfaceBody, oProcessSettings As IJMfgTemplateProcessSettings, ByRef oStartPosition As IJDPosition, ByRef oEndPosition As IJDPosition)
    Const METHOD = "GetCenterFrameFromPlatePart"
    On Error GoTo ErrorHandler

        Dim oGeomHelper As New MfgGeomHelper
        Dim oOutPutCurve As Object
        Dim oFramesColl As New Collection
        Dim nCenterFrameNum As Integer

        ' Get The Reference Planes within the Plate Range
        ' Second Arguments is the line of which direction is same as the Frames' direction, X-Direction
        ' Later it will be determined by the Reference Data
        Dim oVector As New DVector
        Dim bIsTemplateOnFrame As Boolean
        bIsTemplateOnFrame = True
        Dim strTempPosition As String
        strTempPosition = oProcessSettings.TemplatePositionEven
        If strTempPosition = "PositionEven" Then
           bIsTemplateOnFrame = False
        End If

        Dim strTemplateType As String
        strTemplateType = oProcessSettings.TemplateType

        If strTemplateType = "Frame" Then 'Check the Direction Longitudinal, Transversal, Waterline
             oVector.Set 1, 0, 0
        Else 'Frame
            Dim strDirection As String
            strDirection = oProcessSettings.TemplateDirection
            If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
                oVector.Set 0, 1, 0
            ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
                oVector.Set 1, 0, 0
            Else 'Z - Direction(WaterLine)
                oVector.Set 0, 0, 1
            End If
        End If
'        Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
'        Dim dRootX As Double, dRootY As Double, dRootZ As Double
'        oGeomHelper.GetPlatePartAvgPointAvgNormal pPlatePart, bBaseSide, dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ
'        oVector.Set dNormalX, dNormalY, dNormalZ
'        oVector.length = 1#

        oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oVector, oFramesColl
        Set oVector = Nothing

        ' Find out the Center Frame(Reference Plane) among the Reference Planes
        If oFramesColl.count Mod 2 = 0 Then nCenterFrameNum = oFramesColl.count / 2   ' even number
        If oFramesColl.count Mod 2 = 1 Then nCenterFrameNum = (oFramesColl.count + 1) / 2 ' odd number

        ' Get Intersection Points between CenterFrame and Base or Offset Pland
        ' Make the transient plane defined by Middle Point Btw Intersection Points and Normal Vector by Intersection Points
        On Error Resume Next
        If oFramesColl.count > 0 Then
            oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oFramesColl.Item(nCenterFrameNum), oOutPutCurve, oStartPosition, oEndPosition
            If oOutPutCurve Is Nothing Then
                oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oFramesColl.Item(nCenterFrameNum + 1), oOutPutCurve, oStartPosition, oEndPosition
            End If
        End If
        ' Release Objects
        Set oGeomHelper = Nothing
        Set oOutPutCurve = Nothing
        Set oFramesColl = Nothing


    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'Public Function GetFarthestEdgePointAlongVector(oCollEdgePoint As Collection, oBasePlaneNormal As IJDVector) As IJDPosition
'    Const METHOD = "GetFarthestEdgePointAlongVector"
'    On Error GoTo ErrorHandler
'
'        Dim nNumEdgePoint As Integer
'        Dim nIndex As Integer
'        nNumEdgePoint = oCollEdgePoint.count
'
'        ' Create Temporary Plane which is above the EdgePoint along the givenVector to Get the Distance btw EdgePoint and plane
'        Dim oTempPlane As New Plane3d
'        Dim oTempRootPt As New DPosition
'
'        ' Create Dummy Plane to Check the Distance From the Edge Position
'        oTempRootPt.x = oCollEdgePoint.Item(1).x - 5# * oBasePlaneNormal.x
'        oTempRootPt.y = oCollEdgePoint.Item(1).y - 5# * oBasePlaneNormal.y
'        oTempRootPt.z = oCollEdgePoint.Item(1).z - 5# * oBasePlaneNormal.z
'
'        oTempPlane.SetNormal oBasePlaneNormal.x, oBasePlaneNormal.y, oBasePlaneNormal.z
'        oTempPlane.SetRootPoint oTempRootPt.x, oTempRootPt.y, oTempRootPt.z
'
'        Dim oGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
'        Dim oProjectedPosition As IJDPosition
'
'        ' Initially FirstEdgePoint is set as a FarthestEdgePoint
'        Dim dMinDistanceBtwPnts As Double
'        dMinDistanceBtwPnts = 100#
'        Set GetFarthestEdgePointAlongVector = oCollEdgePoint.Item(1)
'
'        For nIndex = 1 To nNumEdgePoint
'            Dim oProjectLine As Line3d
'            Dim dTempDistance As Double
'            Set oProjectLine = CreateInfiniteLine(oCollEdgePoint.Item(nIndex), oBasePlaneNormal)
''            oProjectLine.SetDirection -oBasePlaneNormal.x, -oBasePlaneNormal.y, -oBasePlaneNormal.z
''            oProjectLine.SetStartPoint oCollEdgePoint.Item(nIndex).x, oCollEdgePoint.Item(nIndex).y, oCollEdgePoint.Item(nIndex).z
'
'            ' Get a ProjectedPosition Position and Checked the Distance
'            On Error Resume Next
'            Set oProjectedPosition = oGeomHelper.IntersectCurveWithPlane(oProjectLine, oTempPlane)
'
'           ' oIntersector.PlaceIntersectionObject Nothing, oProjectedLine, oTempPlane, Nothing, oProjectedPosition
'            If Not oProjectedPosition Is Nothing Then
'                dTempDistance = oProjectedPosition.DistPt(oCollEdgePoint.Item(nIndex))
'
'                If dTempDistance < dMinDistanceBtwPnts Then
'                    dMinDistanceBtwPnts = dTempDistance
'                    Set GetFarthestEdgePointAlongVector = Nothing
'                    Set GetFarthestEdgePointAlongVector = oCollEdgePoint.Item(nIndex)
'                End If
'             End If
'        Next nIndex
'
'
'        'Release Objects
'        Set oTempPlane = Nothing
'        Set oTempRootPt = Nothing
'        Set oGeomHelper = Nothing
'        Set oProjectedPosition = Nothing
'
'    Exit Function
'ErrorHandler:
'
'    Err.Raise LogError(Err, MODULE, METHOD).Number
'
'End Function

Public Function GetInterSectionPointBtwLineAndPlane(oPoint As IJDPosition, oProjectVec As IJDVector, oPlane As Plane3d) As IJDPosition
    Const METHOD = "GetInterSectionPointBtwLineAndPlane"
    On Error GoTo ErrorHandler

    Dim oLine As Line3d
    Set oLine = CreateInfiniteLine(oPoint, oProjectVec)
    Dim oOutputPoint As IJDPosition

    On Error Resume Next
    Dim oGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
    Set oOutputPoint = oGeomHelper.IntersectCurveWithPlane(oLine, oPlane)

    Set GetInterSectionPointBtwLineAndPlane = oOutputPoint

    Set oLine = Nothing
    Set oGeomHelper = Nothing
    Set oOutputPoint = Nothing

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Public Function GetProjectedPointOnPlanebyGivenPointAndVector(oPlane As Plane3d, oPosition As IJDPosition, Optional oVector As IJDVector = Nothing) As IJDPosition
'    Const METHOD = "GetProjectedPointOnPlanebyGivenPointandVector"
'    On Error GoTo ErrorHandler
'    Dim strErrorMsg As String
'
'     'check if inputs are okay
'    If oPlane Is Nothing Then
'        strErrorMsg = strErrorMsg & "oPlane is Nothing "
'        GoTo ErrorHandler
'    End If
'    If oPosition Is Nothing Then
'        strErrorMsg = strErrorMsg & "oPosition is Nothing "
'        GoTo ErrorHandler
'    End If
'
'    'if user did not give vector, project in plane's normal direction
'    If oVector Is Nothing Then
'        Dim dNX As Double, dNY As Double, dNZ As Double
'        Set oVector = New DVector
'        oPlane.GetNormal dNX, dNY, dNZ
'        oVector.Set dNX, dNY, dNZ
'    End If
'
'    Dim oProjectLine As Line3d
'    ' Create Very Very Long Line for the Intersection
'    Set oProjectLine = CreateInfiniteLine(oPosition, oVector)
'    Dim oProjectedPoint As IJDPosition
'
'    On Error Resume Next
'    Dim oGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
'    Set oProjectedPoint = oGeomHelper.IntersectCurveWithPlane(oProjectLine, oPlane)
'
'    If Not oProjectedPoint Is Nothing Then
'        Set GetProjectedPointOnPlanebyGivenPointAndVector = oProjectedPoint
'    Else
'        Set GetProjectedPointOnPlanebyGivenPointAndVector = Nothing
'    End If
'
'    'release Objects
'    Set oGeomHelper = Nothing
'    Set oProjectedPoint = Nothing
'    Set oProjectLine = Nothing
'
'    Exit Function
'ErrorHandler:
'
'    Err.Raise LogError(Err, MODULE, METHOD).Number
'
'End Function


Public Function GetProjectedPointOnPlane(oPosition As IJDPosition, oPlane As Plane3d, Optional oVector As IJDVector = Nothing) As IJDPosition
    Const METHOD = "GetProjectedPointOnPlane"
    On Error GoTo ErrorHandler
    Dim strErrorMsg As String

     'check if inputs are okay
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If
    If oPosition Is Nothing Then
        strErrorMsg = strErrorMsg & "oPosition is Nothing "
        GoTo ErrorHandler
    End If


    'if user did not give vector, project in plane's normal direction
    Dim bFreeVector As Boolean
    bFreeVector = False
    If oVector Is Nothing Then
        Dim dNX As Double, dNY As Double, dNZ As Double
        Set oVector = New DVector
        oPlane.GetNormal dNX, dNY, dNZ
        oVector.Set dNX, dNY, dNZ
        bFreeVector = True
    End If

    Dim oProjectLine As Line3d
    ' Create Very Very Long Line for the Intersection
    Set oProjectLine = CreateInfiniteLine(oPosition, oVector)
    Dim oProjectedPoint As IJDPosition

    On Error Resume Next
    Dim oGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
    Set oProjectedPoint = oGeomHelper.IntersectCurveWithPlane(oProjectLine, oPlane)

    Set GetProjectedPointOnPlane = oProjectedPoint

    'release Objects
    Set oGeomHelper = Nothing
    Set oProjectedPoint = Nothing
    Set oProjectLine = Nothing
    If bFreeVector Then Set oVector = Nothing
    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function


'Public Sub CorrectBasePlaneNormal(ByRef oBasePlaneNormal As IJDVector, ByVal oRootPoint As IJDPosition, ByVal oControlLine As Object)
'    Const METHOD = "CorrectBasePlaneNormal"
'    On Error GoTo ErrorHandler
'
'        'Create Transient Plane to Correct BasePlane's Normal
'        Dim oMidPointColl As New Collection
'        Dim oControlLineColl As New Collection
'        oControlLineColl.Add oControlLine
'
'        Dim oGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
'        oGeomHelper.GetMidPtCollOfEdges oControlLineColl, oMidPointColl
'        Set oControlLineColl = Nothing
'
'        Dim oProjectedPoint As IJDPosition
'        Dim oProjectVector As IJDVector
'        Dim oTransientPlane As New Plane3d
'         oTransientPlane.DefineByPointNormal oRootPoint.x, oRootPoint.y, oRootPoint.z, _
'                                                oBasePlaneNormal.x, oBasePlaneNormal.y, oBasePlaneNormal.z
'
'        Set oProjectedPoint = GetProjectedPointOnPlanebyGivenPointAndVector(oTransientPlane, oMidPointColl.Item(1), oBasePlaneNormal)
'        Set oProjectVector = oProjectedPoint.Subtract(oMidPointColl.Item(1))
'
'        'Check the Marking Side later
'        If oProjectVector.Dot(oBasePlaneNormal) > 0 Then
'            oBasePlaneNormal.x = -oBasePlaneNormal.x
'            oBasePlaneNormal.y = -oBasePlaneNormal.y
'            oBasePlaneNormal.z = -oBasePlaneNormal.z
'        End If
'
'        'Release Objects
'        Set oTransientPlane = Nothing
'        Set oMidPointColl = Nothing
'        Set oControlLineColl = Nothing
'        Set oGeomHelper = Nothing
'        Set oProjectedPoint = Nothing
'        Set oProjectVector = Nothing
'
'     Exit Sub
'ErrorHandler:
'
'   Err.Raise LogError(Err, MODULE, METHOD).Number
'End Sub

Public Function CreateInfiniteLine(oPosition As IJDPosition, oDirection As IJDVector) As Line3d
Const METHOD = "CreateInfiniteLine"
    On Error GoTo ErrorHandler

    ' Create very very long line for the Intersection
    Dim oLine As New Line3d
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double

    dStartX = oPosition.x - 1000 * oDirection.x
    dStartY = oPosition.y - 1000 * oDirection.y
    dStartZ = oPosition.z - 1000 * oDirection.z

    dEndX = oPosition.x + 1000 * oDirection.x
    dEndY = oPosition.y + 1000 * oDirection.y
    dEndZ = oPosition.z + 1000 * oDirection.z

    oLine.DefineBy2Points dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    Set CreateInfiniteLine = oLine

    Set oLine = Nothing

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Check the Template Type and If it is a special plate or not
'           2.1. Template Type is PerPendicualrXY(Near to Stem/Stern) or
'               Pentagonal plate or Triangular plate which edge is not parallel to Frame line
'           => Perpendicular XY Algorithm
'           2.2. Template Type is Stem/Stern
'           => Stem/Stern Type, There is no triangular and pentagonal plate(only quadrangle plate).
'           => refer to Analysis- DefineCtlLineAndBPlane.doc
'           => Exactly symmetrical to Center line
'           2.3. Other Case
'               Check the BasePlane Setting Value
'               - BySystem - Check the area of Plate Part if more than one third(33%)  =>  Flat part is Base Plane
'                                                                                    otherwise => Make the BasePlane using the diagonal
'               - NormalToPlate - Make the BasePlane using PinJigs True Natural balance method
'               - ParallelToPlate - Flat part is BasePlane
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckAndGetPlateType(ByVal oProcessSettings As IJMfgTemplateProcessSettings, ByVal oSurfaceBody As IUnknown) As enumPlateType
    Const METHOD = "CheckAndGetPlateType"
    On Error GoTo ErrorHandler

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim strTemplateType As String
    strTemplateType = oProcessSettings.TemplateType
    Dim strBasePlane As String
    strBasePlane = oProcessSettings.BasePlane

    Dim ePlateType As enumPlateType

    If strTemplateType = "Stem/Stern" Or strTemplateType = "CenterLine" Then
        ePlateType = SymToCenterLine
    ElseIf strTemplateType = "PerpendicularXY" Then
        ePlateType = PerpendicularXY
    ElseIf strTemplateType = "Box" Or strTemplateType = "UserDefined Box" Or _
         strTemplateType = "UserDefined Box With Edges" Then
        ePlateType = Box
    Else
        Dim ratio As Double
        Dim oPlane As Plane3d
        On Error Resume Next 'based on surface complexity it might fail
        oGeomHelper.GetMostPlanarFaceOfSheetBody2 oSurfaceBody, ratio, oPlane
        If Err.Number < 0 Then
            LogMessage Err, MODULE, METHOD
        End If
        On Error GoTo ErrorHandler

        If ratio > 0.33 Then ' more than one third
            ePlateType = MostFlatPlate
        Else
            ePlateType = NormalPlate
        End If

    End If

    CheckAndGetPlateType = ePlateType

   Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub CheckAndGetEdgeParallelToFrame(oEdges As IJElements, oProcessSettings As IJMfgTemplateProcessSettings, ByRef IsParallel As Boolean, ByRef oEdge As IUnknown, ByRef count As Long)

    Const METHOD = "IsEdgeParallelToFrame"
    On Error GoTo ErrorHandler

    ' Reference Frame's normal is X - direction
    Dim oVector As IJDVector
    Set oVector = New DVector
    Dim strDirection As String
    strDirection = oProcessSettings.TemplateDirection

     If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
         oVector.Set 0, 1, 0
     ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
         oVector.Set 1, 0, 0
     Else 'Z - Direction(WaterLine)
         oVector.Set 0, 0, 1
     End If
'    oVector.Set 1, 0, 0
    IsParallel = False

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim nCount As Long
    For nCount = 1 To oEdges.count
        Dim oStartPosition As IJDPosition, oEndPosition As IJDPosition
        Dim oTempedge As IUnknown

        Set oTempedge = oEdges.Item(nCount)
        GetEndPoints oTempedge, oStartPosition, oEndPosition

        Dim oVec As IJDVector
        Set oVec = New DVector

        oVec.Set oEndPosition.x - oStartPosition.x, oEndPosition.y - oStartPosition.y, oEndPosition.z - oStartPosition.z
        'Check if One of Edge is parallel to Frame- It means Dot Product is almost zero
        If Abs(oVector.Dot(oVec)) < EPSILON Then ' parallel to Frame
             IsParallel = True
             Set oEdge = oEdges.Item(nCount)
             count = nCount
        End If
    Next nCount

    Set oVector = Nothing
    Set oGeomHelper = Nothing

   Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Public Sub GetEndPoints(ByRef oEdge As IUnknown, ByRef oStartPosition As IJDPosition, ByRef oEndPosition As IJDPosition)
    Const METHOD = "GetEndPoints"
    On Error GoTo ErrorHandler

    If TypeOf oEdge Is IJCurve Then

        Dim oCurve As IJCurve
        Set oCurve = oEdge

        Dim dX1 As Double, dY1 As Double, dZ1 As Double, dX2 As Double, dY2 As Double, dZ2 As Double
        oCurve.EndPoints dX1, dY1, dZ1, dX2, dY2, dZ2
        Set oCurve = Nothing

        Set oStartPosition = New DPosition
        oStartPosition.Set dX1, dY1, dZ1

        Set oEndPosition = New DPosition
        oEndPosition.Set dX2, dY2, dZ2

    ElseIf TypeOf oEdge Is IJWireBody Then

        Dim oWire As IJWireBody
        Set oWire = oEdge

        oWire.GetEndPoints oStartPosition, oEndPosition

        Set oWire = Nothing

    End If

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function IsEqualPoint(oPos1 As IJDPosition, oPos2 As IJDPosition) As Boolean
    Const METHOD = "IsEqualPoint"
    On Error GoTo ErrorHandler

        If Abs(oPos1.x - oPos2.x) < EPSILON And Abs(oPos1.y - oPos2.y) < EPSILON And Abs(oPos1.z - oPos2.z) < EPSILON Then
            IsEqualPoint = True
        Else
            IsEqualPoint = False
        End If

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetProjectedCurveOnPlane(oCurveToProject As IUnknown, oPlane As IJPlane) As IUnknown
    Const METHOD = "GetProjectedCurveOnPlane"
    On Error GoTo ErrorHandler

    Dim oProject As IMSModelGeomOps.Project
    Set oProject = New IMSModelGeomOps.Project

    Dim oPOM As IUnknown
    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oVector As New DVector
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oPlane.GetNormal dNormalX, dNormalY, dNormalZ
    oVector.Set dNormalX, dNormalY, dNormalZ

    Dim oProjectedCurve As IUnknown
    oProject.CurveAlongVectorOnToSurface oPOM, oPlane, oCurveToProject, oVector, Nothing, oProjectedCurve

    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    Dim oCS As IJComplexString

    oMfgMGHelper.WireBodyToComplexString oProjectedCurve, oCS

    If oCS Is Nothing Then
        GoTo ErrorHandler
    Else
        Set GetProjectedCurveOnPlane = oCS
    End If

    ' get rid of the persistent wirebody
    Dim oObject As IJDObject
    Set oObject = oProjectedCurve
    oObject.Remove

    Set oProject = Nothing
    Set oPOM = Nothing
    Set oVector = Nothing
    Set oMfgMGHelper = Nothing

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub GetQuadRangleMidPointsOfButt(ByVal oSurfaceBody As IJSurfaceBody, ByVal oEdges As IJElements, strTemplateDirection As String, ByRef oMidPoint1 As IJDPosition, ByRef oMidPoint2 As IJDPosition)
    Const METHOD = "GetQuadRangleMidPointsOfButt"
    On Error GoTo ErrorHandler

    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oTempEdges As IJElements
    Set oTempEdges = GetEdgesAtGivenDirectionAndPlateNew3(oSurfaceBody, oEdges, strTemplateDirection)

    GetEndPoints oTempEdges.Item(1), oStartPos, oEndPos
    Set oMidPoint1 = New DPosition
    oMidPoint1.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    Set oStartPos = Nothing
    Set oEndPos = Nothing

    GetEndPoints oTempEdges.Item(2), oStartPos, oEndPos
    Set oMidPoint2 = New DPosition
    oMidPoint2.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oTempEdges = Nothing

    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub GetTriangleMidPointsOfButt(oEdges As IJElements, strTemplateDirection As String, oMidPoint1 As IJDPosition, oMidPoint2 As IJDPosition, oRootPoint As IJDPosition)
    Const METHOD = "GetTriangleMidPointsOfButt"
    On Error GoTo ErrorHandler

    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    
    Dim oBestEdge As Object
    Set oBestEdge = GetBestEdge(oEdges, strTemplateDirection, oRootPoint)

    GetEndPoints oBestEdge, oStartPos, oEndPos
    Set oMidPoint1 = New DPosition
    oMidPoint1.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2

    Set oMidPoint2 = GetVertexNotOnEdgeForTrainglePlate(oEdges, oBestEdge)

    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'This method returns two points on the curve(single butt) lying opposite along normal to template direction
'Algorithm
'Get the mid point and normal of surface
'Get the vector normal to template plane
'Get the cross of plate normal with above
'make a plane with mid point of surface and above result
'intersect the surface with this plane
Public Sub GetMidPointsOfButtsSpecial(oSurfaceBody As IJSurfaceBody, oEdges As IJElements, strTemplateDirection As String, oMidPoint1 As IJDPosition, oMidPoint2 As IJDPosition)
    Const METHOD = "GetMidPointsOfButtsSpecial"
    On Error GoTo ErrorHandler
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    'get center of surface
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    oMfgGeomHelper.GetPlatePartAvgPointAvgNormal oSurfaceBody, False, dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ

    'get template direction vector
    Dim oDirVec As IJDVector
    Set oDirVec = New DVector
    If strTemplateDirection = "Longitudinal" Then 'X - Direction(Buttock)
        oDirVec.Set 0, 1, 0
    ElseIf strTemplateDirection = "Transversal" Then 'Y - Direction(Frame)
        oDirVec.Set 1, 0, 0
    Else 'Z - Direction(WaterLine)
        oDirVec.Set 0, 0, 1
    End If
    oDirVec.Length = 1
    
    Dim oPlateNormal As IJDVector
    Set oPlateNormal = New DVector
    oPlateNormal.Set dNormalX, dNormalY, dNormalZ
    
    Dim oBCLPlaneNormal As IJDVector
    
    Set oBCLPlaneNormal = oPlateNormal.Cross(oDirVec)
    
    Dim oBCLPlane As IJPlane
    Set oBCLPlane = New Plane3d
    oBCLPlane.DefineByPointNormal dRootX, dRootY, dRootZ, oBCLPlaneNormal.x, oBCLPlaneNormal.y, oBCLPlaneNormal.z

    Dim oOtputcurve As Object
    oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oOtputcurve, oMidPoint1, oMidPoint2
    
    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'This method returns two points on the curve(single butt) lying opposite along normal to template direction
'Algorithm
'Get the mid point and normal of BasePlane
'Get the vector normal to template plane
'Get the cross of Base plane normal with above
'make a plane with mid point of surface and above result
'intersect the surface with this plane
'(This method  is similar to GetMidPointsOfButtsSpecial. In this method, we use the normal of the BasePlane rather than the Plate normal.
'It is observed that the plate normal is not reliable for plates with curvature and the BCL is not proper in such cases)
Public Sub GetMidPointsOfButtsSpecialForBCL(oPlatePart As IJPlatePart, oSurfaceBody As IJSurfaceBody, oEdges As IJElements, strTemplateDirection As String, pBasePlane As Plane3d, oMidPoint1 As IJDPosition, oMidPoint2 As IJDPosition)
    Const METHOD = "GetMidPointsOfButtsSpecialForBCL"
    On Error GoTo ErrorHandler
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    'get baseplane
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double

    pBasePlane.GetNormal dNormalX, dNormalY, dNormalZ
    
    'get template direction vector
    Dim oDirVec As IJDVector
    Set oDirVec = New DVector
    If strTemplateDirection = "Longitudinal" Then 'X - Direction(Buttock)
        oDirVec.Set 0, 1, 0
    ElseIf strTemplateDirection = "Transversal" Then 'Y - Direction(Frame)
        oDirVec.Set 1, 0, 0
    Else 'Z - Direction(WaterLine)
        oDirVec.Set 0, 0, 1
    End If
    oDirVec.Length = 1
    
    Dim oBasePlaneNormal As IJDVector
    Set oBasePlaneNormal = New DVector
    oBasePlaneNormal.Set dNormalX, dNormalY, dNormalZ
    
    Dim oBCLPlaneNormal As IJDVector
    
    Set oBCLPlaneNormal = oBasePlaneNormal.Cross(oDirVec)
    
    'Set the COG of the plate as the root point of the BCL plane to ensure that the plane intersects the plate in most cases
    'get COG of the plate
    Dim oCOGPosition As IJDPosition
    Dim oCommonStructUtils As IJStructSymbolTools
    Set oCommonStructUtils = New StructSymbolTools
    
    Call oCommonStructUtils.GetCenterOfGravityOfEntity(oPlatePart, oCOGPosition)

    Dim oBCLPlane As IJPlane
    Set oBCLPlane = New Plane3d
    oBCLPlane.DefineByPointNormal oCOGPosition.x, oCOGPosition.y, oCOGPosition.z, oBCLPlaneNormal.x, oBCLPlaneNormal.y, oBCLPlaneNormal.z

    Dim oOtputcurve As Object
    oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oBCLPlane, oOtputcurve, oMidPoint1, oMidPoint2
    
    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateVirtualQuadrangle
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateVirtualQuadrangle(oSurfaceBody As IJSurfaceBody, oProcessSettings As IJMfgTemplateProcessSettings, ByRef oEndPtColl As Collection, oEdgeElements As IJElements)
    Const METHOD = "CreateVirtualQuadrangle"
    On Error GoTo ErrorHandler

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim oVector As New DVector
    Dim strDirection As String
    strDirection = oProcessSettings.TemplateDirection

     If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
         oVector.Set 0, 1, 0
     ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
         oVector.Set 1, 0, 0
     Else 'Z - Direction(WaterLine)
         oVector.Set 0, 0, 1
     End If
     'Vector.Set 1, 0, 0 'Frame( Get Reference Planes along X-direction)

    Dim oFramesColl As New Collection
    oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oVector, oFramesColl
    Set oVector = Nothing
    
    'Find out the Edge parallel to Frame and the other Edge comprises the Virtual Rectangle
    Dim bParallel As Boolean
    Dim oEdge As IUnknown
    Dim Index As Long
    CheckAndGetEdgeParallelToFrame oEdgeElements, oProcessSettings, bParallel, oEdge, Index
    Set oEdgeElements = Nothing

    Dim dDistance As Double
    dDistance = 0
    Dim oVirtualStartPos As IJDPosition, oVirtualEndPos As IJDPosition

    Dim nCount As Long
    For nCount = 1 To oFramesColl.count
        Dim oCurve As Object
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition

        On Error Resume Next
        oGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oFramesColl.Item(nCount), oCurve, oStartPos, oEndPos

        Dim oTempStartPos As IJDPosition, oTempEndPos As IJDPosition
        Dim dTempDistance As Double
        If Not oCurve Is Nothing Then
            oGeomHelper.GetDistBet2Curves oEdge, oCurve, dTempDistance, oTempStartPos, oTempEndPos
            If dTempDistance > dDistance Then
                dDistance = dTempDistance
                Set oVirtualStartPos = oStartPos
                Set oVirtualEndPos = oEndPos
            End If
            Set oCurve = Nothing
            Set oStartPos = Nothing
            Set oEndPos = Nothing
        End If
    Next nCount

    oEndPtColl.Add oVirtualStartPos
    oEndPtColl.Add oVirtualEndPos

    Dim oDirVec1 As IJDVector
    Set oDirVec1 = oVirtualEndPos.Subtract(oVirtualStartPos)

    Set oVirtualStartPos = Nothing
    Set oVirtualEndPos = Nothing

    GetEndPoints oEdge, oVirtualStartPos, oVirtualEndPos
    Dim oDirVec2 As IJDVector
    Set oDirVec2 = oVirtualEndPos.Subtract(oVirtualStartPos)

    If oDirVec1.Dot(oDirVec2) > 0 Then ' same dirction
        oEndPtColl.Add oVirtualEndPos
        oEndPtColl.Add oVirtualStartPos
    Else 'opposit direction
        oEndPtColl.Add oVirtualStartPos
        oEndPtColl.Add oVirtualEndPos
    End If

     'Release Objects
    Set oFramesColl = Nothing
    Set oGeomHelper = Nothing
    Set oVirtualStartPos = Nothing
    Set oVirtualEndPos = Nothing

    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Longitudinal direction - X Dir
' Transversal direction - Y Dir
' Waterline direction - Z Dir
' Note - Longitudinal and WaterLine are also parallel to X Direction
'            - Transversal is parallel to Y Direction
'             Check each Edge and Do Dot Product btw X Dir(also Y Dir) and Edge Direction
'             Bigger one( Dot Product) means Edge is along the this Direction
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEdgesAtGivenDirection(oEdgeElements As IJElements, strTemplateDirection As String) As IJElements
    Const METHOD = "GetEdgesAtGivenDirection"
    On Error GoTo ErrorHandler

    Dim PI As Double
    PI = 3.141592
    Dim dCriteria As Double
'    dCriteria = 1 - Cos(PI / 4) ' 45 degree

    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1

    Dim dNumEdges As Double
    dNumEdges = oEdgeElements.count

    Dim oXDirEdges As IJElements
    Set oXDirEdges = New JObjectCollection

    Dim oYDirEdges As IJElements
    Set oYDirEdges = New JObjectCollection

    Dim oZDirEdges As IJElements
    Set oZDirEdges = New JObjectCollection

    Dim dDot1 As Double, dDot2 As Double, dDot3 As Double
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oDirVec As IJDVector
    Dim nCount As Long, nXDir As Integer, nYDir As Integer, nZDir As Integer

    For nCount = 1 To oEdgeElements.count
        GetEndPoints oEdgeElements.Item(nCount), oStartPos, oEndPos

        Set oDirVec = oEndPos.Subtract(oStartPos)
        oDirVec.Length = 1

        dDot1 = Abs(oDirVec.Dot(oXDir))
        dDot2 = Abs(oDirVec.Dot(oYDir))
        dDot3 = Abs(oDirVec.Dot(oZDir))

        If dDot1 > dDot2 Then ' And Abs(dDot1 - dDot2) > dCriteria Then
            If dDot1 > dDot3 Then 'And Abs(dDot1 - dDot3) > dCriteria Then
                oXDirEdges.Add oEdgeElements.Item(nCount)
            Else
                oZDirEdges.Add oEdgeElements.Item(nCount)
            End If
        Else
            If dDot2 > dDot3 Then 'And Abs(dDot2 - dDot3) > dCriteria Then
                oYDirEdges.Add oEdgeElements.Item(nCount)
            Else
                oZDirEdges.Add oEdgeElements.Item(nCount)
            End If
        End If

        Set oStartPos = Nothing
        Set oEndPos = Nothing
    Next nCount

    ' This is the routine for Special plate
    If oXDirEdges.count = dNumEdges Or oYDirEdges.count = dNumEdges Or oZDirEdges.count = dNumEdges Then
        ReArrangeEdges oXDirEdges, oYDirEdges, oZDirEdges
    End If

'    If oEdgeElements.count = 4 Then
        If strTemplateDirection = "Longitudinal" Then
            If oXDirEdges.count > 0 Then
                 Set GetEdgesAtGivenDirection = oXDirEdges
             ElseIf oZDirEdges.count > 0 Then
                 Set GetEdgesAtGivenDirection = oZDirEdges
             End If
        ElseIf strTemplateDirection = "Waterline" Then
            If oXDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirection = oXDirEdges
            ElseIf oYDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirection = oYDirEdges
            Else
            End If
        Else ' Transversal
            If oYDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirection = oYDirEdges
            ElseIf oZDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirection = oZDirEdges
            End If
        End If

'    ElseIf oEdgeElements.count = 5 Then
'        ' To be implemented
'    End If

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Public Function GetEdgePoints(oEdges As IJElements) As IJElements
'    Const METHOD = "GetEdgePoints"
'    On Error GoTo ErrorHandler
'
'    Dim i As Long, j As Long, InternalCount1 As Long, InternalCount2 As Long, nCount As Long
'    Dim IsEqual1 As Boolean, IsEqual2 As Boolean
'
'    Dim oEdgePoints As IJElements
'    Set oEdgePoints = New JObjectCollection
'
'    Dim oPivotPoint As IJDPosition
'    Dim oTempStartPos As IJDPosition, oTempEndPos As IJDPosition
'    GetEndPoints oEdges.Item(1), oTempStartPos, oTempEndPos
'    Set oPivotPoint = oTempStartPos
'    oEdgePoints.Add oPivotPoint
'    nCount = 1
'
'    For i = 2 To oEdges.count
'       For j = 1 To oEdges.count
'            If j = nCount Then
'            'do noting
'            Else
'                Set oTempStartPos = Nothing
'                Set oTempEndPos = Nothing
'                GetEndPoints oEdges.Item(j), oTempStartPos, oTempEndPos
'
'                 IsEqual1 = IsEqualPoint(oPivotPoint, oTempStartPos)
'                 IsEqual2 = IsEqualPoint(oPivotPoint, oTempEndPos)
'
'                 If IsEqual1 = True And IsEqual2 = False Then
'                     Set oPivotPoint = Nothing
'                     Set oPivotPoint = oTempEndPos
'                     oEdgePoints.Add oPivotPoint
'                     nCount = j
'                     Exit For
'                 ElseIf IsEqual1 = False And IsEqual2 = True Then
'                    Set oPivotPoint = Nothing
'                    Set oPivotPoint = oTempStartPos
'                    oEdgePoints.Add oPivotPoint
'                    nCount = j
'                    Exit For
'                 End If
'             End If
'        Next j
'    Next i
'
'    Set GetEdgePoints = oEdgePoints
'
'    Set oEdgePoints = Nothing
'
'    Exit Function
'ErrorHandler:
'
'End Function
'Public Function GetClosestEdgeFromPosition(dX As Double, dY As Double, dZ As Double, oEdgeElement As IJElements) As Object
'    Const METHOD = "GetClosestEdgeFromPosition"
'    On Error GoTo ErrorHandler
'
'    Dim oPosition As IJDPosition
'    Set oPosition = New DPosition
'    oPosition.Set dX, dY, dZ
'    Dim nCount As Long, count As Long
'    Dim dDistance As Double
'    dDistance = 1000
'
'    For nCount = 1 To oEdgeElement.count
'        Dim oMidPosition As IJDPosition, oStartPos As IJDPosition, oEndPos As IJDPosition
'        Dim dTempDistance As Double
'
'        GetEndPoints oEdgeElement.Item(nCount), oStartPos, oEndPos
'        Set oMidPosition = New DPosition
'        oMidPosition.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
'        dTempDistance = oPosition.DistPt(oMidPosition)
'        If dTempDistance < dDistance Then
'            dDistance = dTempDistance
'            count = nCount
'        End If
'        Set oMidPosition = Nothing
'        Set oStartPos = Nothing
'        Set oEndPos = Nothing
'    Next nCount
'
'    Set GetClosestEdgeFromPosition = oEdgeElement.Item(count)
'
'    Exit Function
'ErrorHandler:
'
'    Err.Raise LogError(Err, MODULE, METHOD).Number
'End Function

Public Sub ReArrangeEdges(oXDirEdges As IJElements, oYDirEdges As IJElements, oZDirEdges As IJElements)
    Const METHOD = "ReArrangeEdges"
    On Error GoTo ErrorHandler

    Dim nNumEdges As Long
    Dim eDirection As enumDirection
'    If Not oXDirEdges.count = 0 Then
'        nNumEdges = oXDirEdges.count
'        eDirection = eXDir
'    ElseIf Not oYDirEdges.count = 0 Then
'        nNumEdges = oYDirEdges.count
'        eDirection = eYDir
'    Else
'        nNumEdges = oZDirEdges.count
'        eDirection = eZDir
'    End If

    Dim nXNumEdges As Long
    Dim nYNumEdges As Long
    Dim nZNumEdges As Long
    nXNumEdges = oXDirEdges.count
    nYNumEdges = oYDirEdges.count
    nZNumEdges = oZDirEdges.count

    If nXNumEdges > nYNumEdges Then
        If nXNumEdges > nZNumEdges Then
            nNumEdges = nXNumEdges
            eDirection = eXDir
        Else
            nNumEdges = nZNumEdges
            eDirection = eZDir
        End If
    Else
        If nYNumEdges > nZNumEdges Then
            nNumEdges = nYNumEdges
            eDirection = eYDir
        Else
            nNumEdges = nZNumEdges
            eDirection = eZDir
        End If
    End If

    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1

    Dim oTempXDirEdges As IJElements
    Set oTempXDirEdges = New JObjectCollection

    Dim oTempYDirEdges As IJElements
    Set oTempYDirEdges = New JObjectCollection

    Dim oTempZDirEdges As IJElements
    Set oTempZDirEdges = New JObjectCollection

    Dim dTempDot1 As Double, dTempDot2 As Double, dTempDot3 As Double
    Dim dDot1 As Double, dDot2 As Double, dDot3 As Double
    dDot1 = 0
    dDot2 = 0
    dDot3 = 0

    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oDirVec As IJDVector
    Dim nIndex As Long, i As Long, j As Long
    Dim eSecondaryDir As enumDirection
    Dim dSecondaryDot As Double
    dSecondaryDot = 0

    Dim Index() As Integer
    ReDim Index(1 To nNumEdges)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check X-Direction
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If eDirection = eXDir Then
        For i = 1 To nNumEdges
            GetEndPoints oXDirEdges.Item(i), oStartPos, oEndPos
            oTempXDirEdges.Add oXDirEdges.Item(i)

            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            If dTempDot2 > dTempDot3 And dTempDot2 > dSecondaryDot Then
                eSecondaryDir = eYDir
                dSecondaryDot = dTempDot2
            ElseIf dTempDot3 > dTempDot2 And dTempDot3 > dSecondaryDot Then
                eSecondaryDir = eZDir
                dSecondaryDot = dTempDot3
            End If
        Next i

        'We just pick up two edges which are most oriented along Xdirection.

'        oYDirEdges.Clear
'        oZDirEdges.Clear
        For i = 1 To nNumEdges
            GetEndPoints oXDirEdges.Item(i), oStartPos, oEndPos
            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            Index(i) = -1
            If eSecondaryDir = eYDir Then
                 If dTempDot2 > dTempDot3 Then
                    oYDirEdges.Add oTempXDirEdges.Item(i)
                    Index(i) = i
                End If
            ElseIf eSecondaryDir = eZDir Then

                If dTempDot3 > dTempDot2 Then
                    oZDirEdges.Add oTempXDirEdges.Item(i)
                    Index(i) = i
                End If
            End If
        Next i

        oXDirEdges.Clear
        For i = 1 To oTempXDirEdges.count
            If Index(i) = -1 Then
                oXDirEdges.Add oTempXDirEdges.Item(i)
            End If
        Next i
    End If
    oTempXDirEdges.Clear
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check Y-Direction
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If eDirection = eYDir Then
        For i = 1 To nNumEdges
            GetEndPoints oYDirEdges.Item(i), oStartPos, oEndPos
            oTempYDirEdges.Add oYDirEdges.Item(i)

            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            If dTempDot1 > dTempDot3 And dTempDot1 > dSecondaryDot Then
                eSecondaryDir = eXDir
                dSecondaryDot = dTempDot1
            ElseIf dTempDot3 > dTempDot1 And dTempDot3 > dSecondaryDot Then
                eSecondaryDir = eZDir
                dSecondaryDot = dTempDot3
            End If
        Next i

        'We just pick up two edges which are most oriented along Xdirection.

'        oXDirEdges.Clear
'        oZDirEdges.Clear
        For i = 1 To nNumEdges
            GetEndPoints oYDirEdges.Item(i), oStartPos, oEndPos
            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            Index(i) = -1
            If eSecondaryDir = eXDir Then
                 If dTempDot1 > dTempDot3 Then
                    oXDirEdges.Add oTempYDirEdges.Item(i)
                    Index(i) = i
                End If
            ElseIf eSecondaryDir = eZDir Then
                If dTempDot3 > dTempDot1 Then
                    oZDirEdges.Add oTempYDirEdges.Item(i)
                    Index(i) = i
                End If
            End If
        Next i

        oYDirEdges.Clear
        For i = 1 To oTempYDirEdges.count
            If Index(i) = -1 Then
                oYDirEdges.Add oTempYDirEdges.Item(i)
            End If
        Next i
    End If
    oTempYDirEdges.Clear
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Check Z-Direction
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If eDirection = eZDir Then
        For i = 1 To nNumEdges
            GetEndPoints oZDirEdges.Item(i), oStartPos, oEndPos
            oTempZDirEdges.Add oZDirEdges.Item(i)

            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            If dTempDot1 > dTempDot2 And dTempDot1 > dSecondaryDot Then
                eSecondaryDir = eXDir
                dSecondaryDot = dTempDot1
            ElseIf dTempDot2 > dTempDot1 And dTempDot2 > dSecondaryDot Then
                eSecondaryDir = eYDir
                dSecondaryDot = dTempDot2
            End If
        Next i

        'We just pick up two edges which are most oriented along Xdirection.
'        oXDirEdges.Clear
'        oYDirEdges.Clear
        For i = 1 To nNumEdges
            GetEndPoints oZDirEdges.Item(i), oStartPos, oEndPos
            Set oDirVec = oEndPos.Subtract(oStartPos)
            oDirVec.Length = 1

            dTempDot1 = Abs(oDirVec.Dot(oXDir))
            dTempDot2 = Abs(oDirVec.Dot(oYDir))
            dTempDot3 = Abs(oDirVec.Dot(oZDir))

            Set oStartPos = Nothing
            Set oEndPos = Nothing

            Index(i) = -1
            If eSecondaryDir = eXDir Then
                 If dTempDot1 > dTempDot2 Then
                    oXDirEdges.Add oTempZDirEdges.Item(i)
                    Index(i) = i
                End If
            ElseIf eSecondaryDir = eYDir Then
                If dTempDot2 > dTempDot1 Then
                    oYDirEdges.Add oTempZDirEdges.Item(i)
                    Index(i) = i
                End If
            End If
        Next i

        oZDirEdges.Clear
        For i = 1 To oTempZDirEdges.count
            If Index(i) = -1 Then
                oZDirEdges.Add oTempZDirEdges.Item(i)
            End If
        Next i
    End If
    oTempZDirEdges.Clear

    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Longitudinal direction - X Dir
' Transversal direction - Y Dir
' Waterline direction - Z Dir
' Note - Longitudinal and WaterLine are also parallel to X Direction
'            - Transversal is parallel to Y Direction
'             Check each Edge and Do Dot Product btw X Dir(also Y Dir) and Edge Direction
'             Bigger one( Dot Product) means Edge is along the this Direction
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEdgesAtGivenDirectionAndPlate(oSurfaceBody As IJSurfaceBody, oEdgeElements As IJElements, strTemplateDirection As String) As IJElements
    Const METHOD = "GetEdgesAtGivenDirectionAndPlate"
    On Error GoTo ErrorHandler

    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1

    Dim dNumEdges As Double
    dNumEdges = oEdgeElements.count

    Dim oXDirEdges As IJElements
    Set oXDirEdges = New JObjectCollection

    Dim oYDirEdges As IJElements
    Set oYDirEdges = New JObjectCollection

    Dim oZDirEdges As IJElements
    Set oZDirEdges = New JObjectCollection

    Dim dDot1 As Double, dDot2 As Double, dDot3 As Double
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oDirVec As IJDVector
    Dim nCount As Long, nXDir As Integer, nYDir As Integer, nZDir As Integer
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double, dRootX As Double, dRootY As Double, dRootZ As Double
    oGeomHelper.GetPlatePartAvgPointAvgNormal oSurfaceBody, False, dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ
    Dim eDir As enumDirection

    If Abs(dNormalX) > Abs(dNormalY) Then
        If Abs(dNormalX) > Abs(dNormalZ) Then
            eDir = eXDir
        Else
            eDir = eZDir
        End If
    Else
        If Abs(dNormalY) > Abs(dNormalZ) Then
            eDir = eYDir
        Else
            eDir = eZDir
        End If
    End If

    For nCount = 1 To oEdgeElements.count
        GetEndPoints oEdgeElements.Item(nCount), oStartPos, oEndPos

        If eDir = eXDir Then 'project Points onto YZ plane passing origin point
            oStartPos.x = 0
            oEndPos.x = 0
        ElseIf eDir = eYDir Then 'project Points onto XZ plane passing origin point
            oStartPos.y = 0
            oEndPos.y = 0
        Else 'project Points onto XY plane passing origin point
            oStartPos.z = 0
            oEndPos.z = 0
        End If

        Set oDirVec = oEndPos.Subtract(oStartPos)
        oDirVec.Length = 1

        dDot1 = Abs(oDirVec.Dot(oXDir))
        dDot2 = Abs(oDirVec.Dot(oYDir))
        dDot3 = Abs(oDirVec.Dot(oZDir))

        If dDot1 > dDot2 Then 'And Abs(dDot1 - dDot2) > dCriteria Then
            If dDot1 > dDot3 Then 'And Abs(dDot1 - dDot3) > dCriteria Then
                oXDirEdges.Add oEdgeElements.Item(nCount)
            Else
                oZDirEdges.Add oEdgeElements.Item(nCount)
            End If
        Else
            If dDot2 > dDot3 Then 'And Abs(dDot2 - dDot3) > dCriteria Then
                oYDirEdges.Add oEdgeElements.Item(nCount)
            Else
                oZDirEdges.Add oEdgeElements.Item(nCount)
            End If
        End If

        Set oStartPos = Nothing
        Set oEndPos = Nothing
    Next nCount

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   This is the routine for some special case :
    '   All edges are along one direction or three out of four are along one direction
    '   In this case, the control line can't be specified. So Rearrange the Edges.

    If oXDirEdges.count = dNumEdges Or oYDirEdges.count = dNumEdges Or oZDirEdges.count = dNumEdges Then
        ReArrangeEdges oXDirEdges, oYDirEdges, oZDirEdges
    End If

    If oEdgeElements.count = 4 Then
        If oXDirEdges.count = 1 Or oYDirEdges.count = 1 Or oZDirEdges.count = 1 Then
            ReArrangeEdges oXDirEdges, oYDirEdges, oZDirEdges
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If oEdgeElements.count = 3 Or oEdgeElements.count = 4 Then
        If strTemplateDirection = "Longitudinal" Then
            If oXDirEdges.count > 0 Then
                 Set GetEdgesAtGivenDirectionAndPlate = oXDirEdges
             ElseIf oZDirEdges.count > 0 Then
                 Set GetEdgesAtGivenDirectionAndPlate = oZDirEdges
             End If
        ElseIf strTemplateDirection = "Waterline" Then
            If oYDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirectionAndPlate = oYDirEdges
            ElseIf oXDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirectionAndPlate = oXDirEdges
            Else
            End If
        Else ' Transversal
            If oYDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirectionAndPlate = oYDirEdges
            ElseIf oZDirEdges.count > 0 Then
                Set GetEdgesAtGivenDirectionAndPlate = oZDirEdges
            End If
        End If
    ElseIf oEdgeElements.count > 4 Then
        If strTemplateDirection = "Longitudinal" Then
            If oXDirEdges.count >= 2 Then
                 SortEdgesForPentagonalPlate oXDirEdges, eXDir
                 Set GetEdgesAtGivenDirectionAndPlate = oXDirEdges
             ElseIf oZDirEdges.count >= 2 Then
                 SortEdgesForPentagonalPlate oZDirEdges, eZDir
                 Set GetEdgesAtGivenDirectionAndPlate = oZDirEdges
             End If
        ElseIf strTemplateDirection = "Waterline" Then
            If oXDirEdges.count >= 2 Then
                SortEdgesForPentagonalPlate oXDirEdges, eXDir
                Set GetEdgesAtGivenDirectionAndPlate = oXDirEdges
            ElseIf oYDirEdges.count >= 2 Then
                SortEdgesForPentagonalPlate oYDirEdges, eYDir
                Set GetEdgesAtGivenDirectionAndPlate = oYDirEdges
            End If
        Else ' Transversal
            If oYDirEdges.count >= 2 Then
                SortEdgesForPentagonalPlate oYDirEdges, eYDir
                Set GetEdgesAtGivenDirectionAndPlate = oYDirEdges
            ElseIf oZDirEdges.count >= 2 Then
                SortEdgesForPentagonalPlate oZDirEdges, eZDir
                Set GetEdgesAtGivenDirectionAndPlate = oZDirEdges
            End If
        End If
    End If

    Exit Function
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub SortEdgesForPentagonalPlate(ByRef oEdges As IJElements, ByVal eDirection As enumDirection)
    Const METHOD = "SortAndChooseBestTwo"
    On Error GoTo ErrorHandler

    Dim nCount As Long, i As Long, j As Long, nIndex As Long
    Dim bPushed() As Boolean
    nCount = oEdges.count

    ReDim bPushed(1 To nCount)
    For i = 1 To nCount
        bPushed(i) = False
    Next i

    Dim oDirection As IJDVector
    Set oDirection = New DVector

    Select Case eDirection
        Case eXDir
            oDirection.Set 1, 0, 0
            oDirection.Length = 1

        Case eYDir
            oDirection.Set 0, 1, 0
            oDirection.Length = 1

        Case eZDir
            oDirection.Set 0, 0, 1
            oDirection.Length = 1
    End Select

    Dim oTempEdges As IJElements
    Set oTempEdges = oEdges.Clone

    Dim dTempDot As Double, dDot As Double
    dDot = 0
    dTempDot = 0

    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oTempVector As IJDVector
    oEdges.Clear

    For i = 1 To nCount
        For j = 1 To nCount
            If bPushed(j) = False Then
                GetEndPoints oTempEdges.Item(j), oStartPos, oEndPos
                Set oTempVector = oEndPos.Subtract(oStartPos)
                oTempVector.Length = 1
                dTempDot = Abs(oTempVector.Dot(oDirection))

                If dTempDot > dDot Then
                    dDot = dTempDot
                    nIndex = j
                End If

                dTempDot = 0
                Set oStartPos = Nothing
                Set oEndPos = Nothing
                Set oTempVector = Nothing
            End If
        Next j

        oEdges.Add oTempEdges.Item(nIndex)

        bPushed(nIndex) = True
        dDot = 0
    Next i
    oTempEdges.Clear

    Exit Sub
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
'Public Function WireBodyToComplexString(ByVal oWireBody As IJWireBody) As IJComplexString
'   Const METHOD = "WireBodyToComplexString"
'   On Error GoTo ErrorHandler
'
'    Dim oMathHelper As IJMfgMGHelper
'    Dim oCS As IJComplexString
'
'    If Not oWireBody Is Nothing Then
'        Set oMathHelper = New MfgMGHelper
'        If Not oMathHelper Is Nothing Then
'            oMathHelper.WireBodyToComplexString oWireBody, oCS
'            If Not oCS Is Nothing Then
'                Set WireBodyToComplexString = oCS
'            Else
'                GoTo ErrorHandler
'            End If
'        Else
'            GoTo ErrorHandler
'        End If
'    Else
'        GoTo ErrorHandler
'    End If
'
'    Set oMathHelper = Nothing
'    Set oCS = Nothing
'    Exit Function
'
'ErrorHandler:
'    Set WireBodyToComplexString = Nothing
'
'    Err.Raise LogError(Err, MODULE, METHOD).Number
'End Function


Public Function MaxOfThree(ByVal dOne As Double, ByVal dTwo As Double, ByVal dThree As Double) As Double
    If dOne >= dTwo Then
        If dOne >= dThree Then
            MaxOfThree = dOne
        Else
            MaxOfThree = dThree
        End If
    Else
        If dTwo >= dThree Then
            MaxOfThree = dTwo
        Else
            MaxOfThree = dThree
        End If
    End If
End Function

Public Function MinOfThree(ByVal dOne As Double, ByVal dTwo As Double, ByVal dThree As Double) As Double
    If dOne <= dTwo Then
        If dOne <= dThree Then
            MinOfThree = dOne
        Else
            MinOfThree = dThree
        End If
    Else
        If dTwo <= dThree Then
            MinOfThree = dTwo
        Else
            MinOfThree = dThree
        End If
    End If
End Function

Public Function GetSurfaceWithoutFeatures(oPlatePart As IJPlatePart, bBaseSide As Boolean) As IUnknown
Const METHOD = "GetSurfaceWithoutFeatures"
On Error Resume Next

    Set GetSurfaceWithoutFeatures = Nothing 'initliaze return value to nothing
    If oPlatePart Is Nothing Then Exit Function ' check if there is input else return nothing

    Dim oSurfaceBody As IJSurfaceBody

    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oPartSupport As IJPartSupport

    Set oPartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport.Part = oPlatePart
    Set oPlatePartSupport = oPartSupport

    ' there is an error in oPlatePartSupport which gets surface in opposite way
    If Not oPlatePartSupport Is Nothing Then
        If bBaseSide Then
            oPlatePartSupport.GetSurfaceWithoutFeatures PlateBaseSide, oSurfaceBody
        Else
            oPlatePartSupport.GetSurfaceWithoutFeatures PlateOffsetSide, oSurfaceBody
        End If
    End If
    Set GetSurfaceWithoutFeatures = oSurfaceBody

CleanUp:
    ' delete all local objects
    Set oPlatePartSupport = Nothing
    Set oSurfaceBody = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetIntersectionBetweenCurves(oObject1 As Object, oObject2 As Object) As IJDPosition
Const METHOD = "GetIntersectionBetweenCurves"
On Error Resume Next

    Dim strErrorMsg As String

    'check if inputs are okay
    If oObject1 Is Nothing Then
        strErrorMsg = strErrorMsg & "oObject1 is Nothing "
        GoTo ErrorHandler
    End If
    If oObject2 Is Nothing Then
        strErrorMsg = strErrorMsg & "oObject2 is Nothing "
        GoTo ErrorHandler
    End If

    Dim oMfgGeomHelper As MfgGeomHelper
    Dim oMfgMGHelper As MfgMGHelper

    Set oMfgGeomHelper = New MfgGeomHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    Dim oCurve1 As IJCurve
    Dim oCurve2 As IJCurve

    If TypeOf oObject1 Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oObject1, oCurve1
    Else
        Set oCurve1 = oObject1
    End If

    If TypeOf oObject2 Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oObject2, oCurve2
    Else
        Set oCurve2 = oObject2
    End If

    If oCurve1 Is Nothing Or oCurve2 Is Nothing Then GoTo CleanUp

    Dim oIntersectionObject As Object
    Dim oIntersectionPos As IJDPosition
    Dim oIntersectionElems As IJElements
    Set oIntersectionObject = oMfgGeomHelper.IntersectCurveWithCurve(oCurve1, oCurve2)
    Dim dDist As Double
    Dim oPoint1 As IJDPosition, oPoint2 As IJDPosition


    If Not oIntersectionObject Is Nothing Then
        If TypeOf oIntersectionObject Is IJDPosition Then
            Set oIntersectionPos = oIntersectionObject
        Else
            Set oIntersectionElems = oIntersectionObject
            Set oIntersectionPos = oIntersectionElems.Item(1)
            Set oIntersectionElems = Nothing
        End If
    Else
        'sometimes IntersectCurveWithCurve can fail so see if distance method works
        oMfgGeomHelper.GetDistBet2Curves oCurve1, oCurve2, dDist, oPoint1, oPoint2
        If dDist < 0.001 Then
            Set oIntersectionPos = oPoint1
        End If
    End If

    Set GetIntersectionBetweenCurves = oIntersectionPos

CleanUp:

    ' delete all local objects
    Set oPoint1 = Nothing
    Set oPoint2 = Nothing
    Set oCurve1 = Nothing
    Set oCurve2 = Nothing
    Set oIntersectionPos = Nothing
    Set oIntersectionObject = Nothing
    Set oIntersectionPos = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgMGHelper = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetIntersectionBetweenCurveAndPlane(oCurve As Object, oPlane As IJPlane) As IJDPosition
Const METHOD = "GetIntersectionBetweenCurveAndPlane"
On Error Resume Next

    Dim strErrorMsg As String

    'check if inputs are okay
    If oCurve Is Nothing Then
        strErrorMsg = strErrorMsg & "oCurve is Nothing "
        GoTo ErrorHandler
    End If
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If


    Dim oMfgGeomHelper As MfgGeomHelper
    Dim oMfgMGHelper As MfgMGHelper

    Set oMfgGeomHelper = New MfgGeomHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper



    Dim oIntersectionObject As Object
    Dim oIntersectionPos As IJDPosition
    Dim oIntersectionElems As IJElements

    Dim oCurve1 As IJCurve

    If TypeOf oCurve Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oCurve, oCurve1
    Else
        Set oCurve1 = oCurve
    End If

    Set oIntersectionObject = oMfgGeomHelper.IntersectCurveWithPlane(oCurve1, oPlane)

    If Not oIntersectionObject Is Nothing Then
        If TypeOf oIntersectionObject Is IJDPosition Then
            Set oIntersectionPos = oIntersectionObject
        Else
            Set oIntersectionElems = oIntersectionObject
            Set oIntersectionPos = oIntersectionElems.Item(1)
            Set oIntersectionElems = Nothing
        End If
    End If

    Set GetIntersectionBetweenCurveAndPlane = oIntersectionPos

    ' delete all local objects
    Set oIntersectionObject = Nothing
    Set oIntersectionPos = Nothing
    Set oMfgMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IsEdgeOnPlane(oEdge As Object, oPlane As Object, oControlLine As Object) As Boolean
    Const METHOD = "IsEdgeOnPlane"
    On Error GoTo ErrorHandler
    Dim strErrorMsg As String

    If oPlane Is Nothing Then
        strErrorMsg = "oFrame is Nothing"
        GoTo ErrorHandler
    End If
    If oEdge Is Nothing Then
        strErrorMsg = "oEdge is Nothing"
        GoTo ErrorHandler
    End If
    If oControlLine Is Nothing Then
        strErrorMsg = "oControlLine is Nothing"
        GoTo ErrorHandler
    End If

    Dim oIntersectionPoint1 As IJDPosition
    Dim oIntersectionPoint2 As IJDPosition

    IsEdgeOnPlane = False

     'check if the edge and plane are near (check w.r.t to control line)
    ' checking intersection with edge and plane fails sometimes
    ' hence check dist between intersection points w.r.t controlline
    Set oIntersectionPoint1 = GetIntersectionBetweenCurveAndPlane(oControlLine, oPlane)
    If oIntersectionPoint1 Is Nothing Then GoTo CleanUp

    Set oIntersectionPoint2 = GetIntersectionBetweenCurves(oControlLine, oEdge)
    If oIntersectionPoint2 Is Nothing Then GoTo CleanUp


    If Not oIntersectionPoint1.DistPt(oIntersectionPoint2) < FRAME_EDGE_DIST Then GoTo CleanUp

    Dim oIJPlane As IJPlane
    Dim oPlaneVec As IJDVector, oEdgeVec As IJDVector
    Dim dNX As Double, dNY As Double, dNZ As Double
    Dim bEdgePlanar As Boolean

    Set oIJPlane = oPlane
    Set oPlaneVec = New DVector
    
    oIJPlane.GetNormal dNX, dNY, dNZ
    oPlaneVec.Set dNX, dNY, dNZ
    
    IsEdgeOnPlane = IsCurveOnGivenPlane(oEdge, oPlaneVec)

CleanUp:
    Set oIntersectionPoint1 = Nothing
    Set oIntersectionPoint2 = Nothing
    Set oPlaneVec = Nothing
    Set oEdgeVec = Nothing
    Set oIJPlane = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetEdgeOnFrame(oFrame As Object, oEdgeElems As IJElements, oControlLine As Object) As Object
    Const METHOD = "IsFrameOnEdge"
    On Error GoTo ErrorHandler
    Dim strErrorMsg As String

    If oFrame Is Nothing Then
        strErrorMsg = "oFrame is Nothing"
        GoTo ErrorHandler
    End If
    If oEdgeElems Is Nothing Then
        strErrorMsg = "oEdgeElems is Nothing"
        GoTo ErrorHandler
    End If
    If oControlLine Is Nothing Then
        strErrorMsg = "oControlLine is Nothing"
        GoTo ErrorHandler
    End If

    Dim nEdgeIndex As Long
    Dim oPlateEdge As Object
    Dim bEdgeOnPlane As Boolean

    For nEdgeIndex = 1 To oEdgeElems.count
        Set oPlateEdge = oEdgeElems.Item(nEdgeIndex)
        bEdgeOnPlane = IsEdgeOnPlane(oPlateEdge, oFrame, oControlLine)
        If bEdgeOnPlane Then
            Set GetEdgeOnFrame = oPlateEdge
            Exit For
        End If
    Next

CleanUp:
    Set oPlateEdge = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IsEdgeTemplateDefined(oTemplateDataElems As IJElements, oControlPoint As IJDPosition) As Boolean
Const METHOD = "CreateFinitePlaneFromPlane"
On Error GoTo ErrorHandler
    Dim oTemplateData As TemplateData
    Dim nIndex As Long

    IsEdgeTemplateDefined = False

    For nIndex = 1 To oTemplateDataElems.count
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        If oControlPoint.DistPt(oTemplateData.ControlPoint) < 2 * EPSILON Then '.001 is too small
            IsEdgeTemplateDefined = True
            GoTo CleanUp
        End If
        Set oTemplateData = Nothing
    Next
CleanUp:
    Set oTemplateData = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IsCurvePlanar(ByRef oObject As Object, ByRef bIsLine As Boolean, ByRef oPlanarDirVec As IJDVector) As Boolean
 Const METHOD = "IsCurvePlanar"
On Error GoTo ErrorHandler
    bIsLine = False
    Set oPlanarDirVec = Nothing
    
    Dim oCurveElems As IJElements
    Set oCurveElems = New JObjectCollection

    Dim oCS As IJComplexString
    Dim oObjUnKnown As IUnknown
    Dim oMfgGeomHelper As MfgGeomHelper
    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    If TypeOf oObject Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oObject, oCS
        Set oObjUnKnown = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oCS)
        oCurveElems.Add oObjUnKnown
        Set oObjUnKnown = Nothing
    ElseIf TypeOf oObject Is IJCurve Then
        oCurveElems.Add oObject
    ElseIf TypeOf oObject Is IJComplexString Then
        oCurveElems.Add oObject
    End If

    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElems)

    Dim Length1 As Double, length2 As Double, length3 As Double
    Dim Points(1 To 4) As IJDPosition
    Dim oVector1 As IJDVector, oVector2 As IJDVector, oVector3 As IJDVector

    Set Points(1) = oBoxPoints.Item(1)
    Set Points(2) = oBoxPoints.Item(2)
    Set Points(3) = oBoxPoints.Item(3)
    Set Points(4) = oBoxPoints.Item(4)

    Length1 = Points(1).DistPt(Points(2))
    length2 = Points(2).DistPt(Points(3))
    length3 = Points(1).DistPt(Points(4))

    Set oVector1 = Points(2).Subtract(Points(1))
    Set oVector2 = Points(3).Subtract(Points(2))
    Set oVector3 = Points(4).Subtract(Points(1))
    
    oVector1.Length = 1
    oVector2.Length = 1
    oVector3.Length = 1

    IsCurvePlanar = False
    Dim dMinLength As Double

    dMinLength = MinOfThree(Length1, length2, length3)
    If dMinLength < 0.01 Then
        IsCurvePlanar = True
        
        If Length1 < 0.01 Then
            Set oPlanarDirVec = oVector1
            If length2 < 0.01 Then
                bIsLine = True
                Set oPlanarDirVec = Nothing 'it is line you can not construct a plane
            End If
            If length3 < 0.01 Then 'Hope input is not a point
                bIsLine = True
                Set oPlanarDirVec = Nothing 'it is line you can not construct a plane
            End If
        ElseIf length2 < 0.01 Then
            Set oPlanarDirVec = oVector2
            If Length1 < 0.01 Then
                bIsLine = True
                Set oPlanarDirVec = Nothing
            End If
            If length3 < 0.01 Then
                bIsLine = True
                Set oPlanarDirVec = Nothing
            End If
        Else 'If length3 < 0.01 Then
            Set oPlanarDirVec = oVector3
            If Length1 < 0.01 Then
                bIsLine = True
                Set oPlanarDirVec = Nothing
            End If
            If length2 < 0.01 Then
                bIsLine = True
                Set oPlanarDirVec = Nothing
            End If
        End If
    Else
       IsCurvePlanar = False
    End If
CleanUp:
    Set Points(1) = Nothing
    Set Points(2) = Nothing
    Set Points(3) = Nothing
    Set Points(4) = Nothing
    Set oBoxPoints = Nothing
    Set oCurveElems = Nothing
    Set oVector1 = Nothing
    Set oVector2 = Nothing
    Set oVector3 = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgMGHelper = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function IsCurveOnGivenPlane(ByVal oObject As Object, ByVal oGivenPlaneNormalVec As IJDVector) As Boolean
Const METHOD = "IsCurveOnGivenPlane"
On Error GoTo ErrorHandler
    Dim oCurveElems As IJElements
    Set oCurveElems = New JObjectCollection

    Dim oCS As IJComplexString
    Dim oObjUnKnown As IUnknown
    Dim oMfgGeomHelper As MfgGeomHelper
    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

    If TypeOf oObject Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oObject, oCS
        Set oObjUnKnown = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oCS)
        oCurveElems.Add oObjUnKnown
        Set oObjUnKnown = Nothing
    ElseIf TypeOf oObject Is IJCurve Then
        oCurveElems.Add oObject
    ElseIf TypeOf oObject Is IJComplexString Then
        oCurveElems.Add oObject
    End If

    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElems, oGivenPlaneNormalVec)

    Dim Points(1 To 4) As IJDPosition

    Set Points(1) = oBoxPoints.Item(1)
    Set Points(2) = oBoxPoints.Item(2)
    Set Points(3) = oBoxPoints.Item(3)
    Set Points(4) = oBoxPoints.Item(4)

    'distance between point1 and point4 should be negligible
    If Points(1).DistPt(Points(4)) < EPSILON_PLANAR Then
        IsCurveOnGivenPlane = True
    End If
CleanUp:
    Set Points(1) = Nothing
    Set Points(2) = Nothing
    Set Points(3) = Nothing
    Set Points(4) = Nothing
    Set oBoxPoints = Nothing
    Set oCurveElems = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgMGHelper = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SnapPlaneToCS(ByRef oPlane As Plane3d, dToleranceDegrees As Double)
Const METHOD = "SnapPlaneToCS"
On Error GoTo ErrorHandler

    Dim strErrorMsg As String
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If
    Dim dNX As Double
    Dim dNY As Double
    Dim dNZ As Double

    Dim dRX As Double
    Dim dRY As Double
    Dim dRZ As Double

    Dim oPlaneNormal As IJDVector
    Set oPlaneNormal = New DVector

    oPlane.GetNormal dNX, dNY, dNZ

    oPlane.GetRootPoint dRX, dRY, dRZ

    oPlaneNormal.Set dNX, dNY, dNZ

    Dim oX As IJDVector
    Set oX = New DVector
    Dim oY As IJDVector
    Set oY = New DVector
    Dim oZ As IJDVector
    Set oZ = New DVector

    oX.Set 1, 0, 0
    oY.Set 0, 1, 0
    oZ.Set 0, 0, 1

    Dim dToleranceRadians As Double
    If dToleranceDegrees > 360# Then dToleranceDegrees = dToleranceDegrees - 360
    'dToleranceRadians = (dToleranceDegrees * 0.0087301587)     ' =  dToleranceDegrees * pi / 360
    'pi radians is 180 degrees (not 360)
    dToleranceRadians = (dToleranceDegrees * 0.01745329252)     ' =  dToleranceDegrees * pi / 180

    ' The logic of all three if statements below is wrong.  The LHS on each inequality is
    ' difference in cosine values.  Right side is angular measurement.
'    If (1 - Abs(oX.Dot(oPlaneNormal))) < dToleranceRadians Then
'        MsgBox "Plane with normal (" & dNX & ", " & dNY & ", " & dNZ & ") reset to X axis!"
'        oPlaneNormal.Set 1, 0, 0
'    ElseIf (1 - Abs(oY.Dot(oPlaneNormal))) < dToleranceRadians Then
'        MsgBox "Plane with normal (" & dNX & ", " & dNY & ", " & dNZ & ") reset to Y axis!"
'        oPlaneNormal.Set 0, 1, 0
'    ElseIf (1 - Abs(oZ.Dot(oPlaneNormal))) < dToleranceRadians Then
'        MsgBox "Plane with normal (" & dNX & ", " & dNY & ", " & dNZ & ") reset to Z axis!"
'        oPlaneNormal.Set 0, 0, 1
'    End If

    oPlaneNormal.Get dNX, dNY, dNZ

    oPlane.DefineByPointNormal dRX, dRY, dRZ, dNX, dNY, dNZ
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Longitudinal direction - X Dir
' Transversal direction - Y Dir
' Waterline direction - Z Dir
' Note - Longitudinal and WaterLine are also parallel to X Direction
'            - Transversal is parallel to Y Direction
'             Check each Edge and Do Dot Product btw X Dir(also Y Dir) and Edge Direction
'             Bigger one( Dot Product) means Edge is along the this Direction
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEdgesAtGivenDirectionAndPlateNew(oSurfaceBody As IJSurfaceBody, oEdgeElements As IJElements, strDirection As String) As IJElements
    Const METHOD = "GetEdgesAtGivenDirection"
    On Error GoTo ErrorHandler

    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1

    Dim dNumEdges As Double
    dNumEdges = oEdgeElements.count

    Dim oXDirEdges As IJElements
    Set oXDirEdges = New JObjectCollection

    Dim oYDirEdges As IJElements
    Set oYDirEdges = New JObjectCollection

    Dim oZDirEdges As IJElements
    Set oZDirEdges = New JObjectCollection

    Dim oOutputEdges As IJElements
    Set oOutputEdges = New JObjectCollection

    Dim dDot1 As Double, dDot2 As Double, dDot3 As Double
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oDirVec As IJDVector
    Dim nCount As Long, nXDir As Integer, nYDir As Integer, nZDir As Integer
    Dim oEdge As IUnknown

    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    Dim eDir As enumDirection

    If strDirection = "Transversal" Then
        If Abs(oNormal.y) > Abs(oNormal.z) Then
            eDir = eYDir
        Else
            eDir = eZDir
        End If
    ElseIf strDirection = "Longitudinal" Then
        If Abs(oNormal.x) > Abs(oNormal.z) Then
            eDir = eXDir
        Else
            eDir = eZDir
        End If
    ElseIf strDirection = "Waterline" Then
        If Abs(oNormal.x) > Abs(oNormal.y) Then
            eDir = eXDir
        Else
            eDir = eYDir
        End If
    End If

    For nCount = 1 To oEdgeElements.count
        Set oEdge = oEdgeElements.Item(nCount)
        GetEndPoints oEdge, oStartPos, oEndPos

        If eDir = eXDir Then 'project Points onto YZ plane passing origin point
            oStartPos.x = 0
            oEndPos.x = 0
        ElseIf eDir = eYDir Then 'project Points onto XZ plane passing origin point
            oStartPos.y = 0
            oEndPos.y = 0
        Else 'project Points onto XY plane passing origin point
            oStartPos.z = 0
            oEndPos.z = 0
        End If

        Set oDirVec = oEndPos.Subtract(oStartPos)
        oDirVec.Length = 1

        dDot1 = Abs(oDirVec.Dot(oXDir))
        dDot2 = Abs(oDirVec.Dot(oYDir))
        dDot3 = Abs(oDirVec.Dot(oZDir))

        If dDot1 > dDot2 Then 'And Abs(dDot1 - dDot2) > dCriteria Then
            If dDot1 > dDot3 Then 'And Abs(dDot1 - dDot3) > dCriteria Then
                oXDirEdges.Add oEdge
            Else
                oZDirEdges.Add oEdge
            End If
        Else
            If dDot2 > dDot3 Then 'And Abs(dDot2 - dDot3) > dCriteria Then
                oYDirEdges.Add oEdge
            Else
                oZDirEdges.Add oEdge
            End If
        End If

        Set oStartPos = Nothing
        Set oEndPos = Nothing
    Next nCount

    Dim oLowEdge As IUnknown
    Dim oHighEdge As IUnknown

    nCount = 0

    If strDirection = "Transversal" Then
        If oYDirEdges.count > oZDirEdges.count Then
            While oYDirEdges.count < 2
                ' we have a problem, there are not enough edges in required dir. Odd shaped plate
                ' so find edge that is closest to required dir
                Set oEdge = FindBestEdgeNormalToGivenDir(oXDirEdges, oXDir)
                nCount = oXDirEdges.GetIndex(oEdge)
                oYDirEdges.Add oEdge
                oXDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oYDirEdges, oRootPoint, oXDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oYDirEdges, oRootPoint, oXDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oYDirEdges, oXDirEdges, oXDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oYDirEdges, oRootPoint, oXDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oYDirEdges, oRootPoint, oXDir)
            End If
        Else
            While oZDirEdges.count < 2
                ' we have a problem, there are not enough edges in required dir. Odd shaped plate
                ' so find edge that is closest to required dir
                Set oEdge = FindBestEdgeNormalToGivenDir(oXDirEdges, oXDir)
                nCount = oXDirEdges.GetIndex(oEdge)
                oZDirEdges.Add oEdge
                oXDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oZDirEdges, oRootPoint, oXDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oZDirEdges, oRootPoint, oXDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oZDirEdges, oXDirEdges, oXDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oZDirEdges, oRootPoint, oXDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oZDirEdges, oRootPoint, oXDir)
            End If
        End If
    ElseIf strDirection = "Longitudinal" Then
        If oXDirEdges.count > oZDirEdges.count Then
            While oXDirEdges.count < 2
                Set oEdge = FindBestEdgeNormalToGivenDir(oYDirEdges, oYDir)
                nCount = oYDirEdges.GetIndex(oEdge)
                oXDirEdges.Add oEdge
                oYDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oXDirEdges, oRootPoint, oYDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oXDirEdges, oRootPoint, oYDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oXDirEdges, oYDirEdges, oYDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oXDirEdges, oRootPoint, oYDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oXDirEdges, oRootPoint, oYDir)
            End If
        Else
            While oZDirEdges.count < 2
                Set oEdge = FindBestEdgeNormalToGivenDir(oYDirEdges, oYDir)
                nCount = oYDirEdges.GetIndex(oEdge)
                oZDirEdges.Add oEdge
                oYDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oZDirEdges, oRootPoint, oYDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oZDirEdges, oRootPoint, oYDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oZDirEdges, oYDirEdges, oYDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oZDirEdges, oRootPoint, oYDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oZDirEdges, oRootPoint, oYDir)
            End If
        End If
    ElseIf strDirection = "Waterline" Then
        If oXDirEdges.count > oYDirEdges.count Then
            While oXDirEdges.count < 2
                Set oEdge = FindBestEdgeNormalToGivenDir(oZDirEdges, oZDir)
                nCount = oZDirEdges.GetIndex(oEdge)
                oXDirEdges.Add oEdge
                oZDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oXDirEdges, oRootPoint, oZDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oXDirEdges, oRootPoint, oZDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oXDirEdges, oZDirEdges, oZDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oXDirEdges, oRootPoint, oZDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oXDirEdges, oRootPoint, oZDir)
            End If
        Else
            While oYDirEdges.count < 2
                Set oEdge = FindBestEdgeNormalToGivenDir(oZDirEdges, oZDir)
                nCount = oZDirEdges.GetIndex(oEdge)
                oYDirEdges.Add oEdge
                oZDirEdges.Remove nCount
            Wend
            Set oLowEdge = GetEdgeLowSideofPoint(oYDirEdges, oRootPoint, oZDir)
            Set oHighEdge = GetEdgeHighSideofPoint(oYDirEdges, oRootPoint, oZDir)
            
            'Check if adjacent edges are added to HighEdge and LowEdge, Remove one of the adjescent edge and add the opposite edge
            If dNumEdges > 3 And CheckForAdjacentEdgesAndRemove(oLowEdge, oHighEdge, oYDirEdges, oZDirEdges, oZDir) = True Then
                Set oLowEdge = GetEdgeLowSideofPoint(oYDirEdges, oRootPoint, oZDir)
                Set oHighEdge = GetEdgeHighSideofPoint(oYDirEdges, oRootPoint, oZDir)
            End If
        End If
    End If
        
    oOutputEdges.Add oLowEdge
    oOutputEdges.Add oHighEdge

    Set GetEdgesAtGivenDirectionAndPlateNew = oOutputEdges

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function CheckForAdjacentEdgesAndRemove(oLowEdge As IUnknown, oHighEdge As IUnknown, oEdgeColl As IJElements, oRemEdgeColl As IJElements, oDir As IJDVector) As Boolean
    
    Dim oLowStartPos As IJDPosition
    Dim oLowEndPos As IJDPosition
    
    Dim oHighStartPos As IJDPosition
    Dim oHighEndPos As IJDPosition
    
    GetEndPoints oLowEdge, oLowStartPos, oLowEndPos
    GetEndPoints oHighEdge, oHighStartPos, oHighEndPos
    
    'Compare the end points
    If (Abs(oLowStartPos.x - oHighEndPos.x) <= DISTANCETOLERANCE And Abs(oLowStartPos.y - oHighEndPos.y) <= DISTANCETOLERANCE And Abs(oLowStartPos.z - oHighEndPos.z) <= DISTANCETOLERANCE) Then
        CheckForAdjacentEdgesAndRemove = True
    ElseIf (Abs(oLowEndPos.x - oHighStartPos.x) <= DISTANCETOLERANCE And Abs(oLowEndPos.y - oHighStartPos.y) <= DISTANCETOLERANCE And Abs(oLowEndPos.z - oHighStartPos.z) <= DISTANCETOLERANCE) Then
        CheckForAdjacentEdgesAndRemove = True
    Else
        Exit Function
    End If
    
    Dim nCount As Long
    Dim oEdge As IUnknown
    'Find the best edge and keep it in the edge coll.
    Set oEdge = FindBestEdgeNormalToGivenDir(oEdgeColl, oDir)
    oEdgeColl.RemoveElements oEdgeColl
    oEdgeColl.Add oEdge
    
     'We need to remove the adjescent edge to the best edge(One among High/Low edges) if any from the oRemEdgeColl
    For nCount = 1 To oRemEdgeColl.count
        Dim oOppEdge As IUnknown
        Set oOppEdge = oRemEdgeColl.Item(nCount)
        
        Dim oOppEdgeStartPos As IJDPosition
        Dim oOppEdgeLowEndPos As IJDPosition
        
        Dim oEdgeStartPos As IJDPosition
        Dim oEdgeEndPos As IJDPosition
        
        GetEndPoints oOppEdge, oOppEdgeStartPos, oOppEdgeLowEndPos
        GetEndPoints oEdge, oEdgeStartPos, oEdgeEndPos
        'Compare the end points
        If (Abs(oOppEdgeStartPos.x - oEdgeEndPos.x) <= DISTANCETOLERANCE And Abs(oOppEdgeStartPos.y - oEdgeEndPos.y) <= DISTANCETOLERANCE And Abs(oOppEdgeStartPos.z - oEdgeEndPos.z) <= DISTANCETOLERANCE) Then
            oRemEdgeColl.Remove oOppEdge
        ElseIf (Abs(oOppEdgeLowEndPos.x - oEdgeStartPos.x) <= DISTANCETOLERANCE And Abs(oOppEdgeLowEndPos.y - oEdgeStartPos.y) <= DISTANCETOLERANCE And Abs(oOppEdgeLowEndPos.z - oEdgeStartPos.z) <= DISTANCETOLERANCE) Then
            oRemEdgeColl.Remove oOppEdge
        End If
    Next
    
    Set oEdge = Nothing
    'Find the opposite edge in the other edge collection and add it
    'in the edge coll that is used for creation of base plane.
    Set oEdge = FindBestEdgeNormalToGivenDir(oRemEdgeColl, oDir)
    nCount = oRemEdgeColl.GetIndex(oEdge)
    oEdgeColl.Add oEdge
    oRemEdgeColl.Remove nCount
    Set oEdge = Nothing
      
    Exit Function
    
End Function

Public Function GetEdgeLowSideofPoint(oEdgeCol As IJElements, oPoint As IJDPosition, oDirVec As IJDVector) As Object
    Dim nCount As Long
    Dim oEdge As IUnknown
    Dim dMaxDist As Double
    Dim dDist As Double
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oMidPos As IJDPosition
    Dim oVec As IJDVector

'    MsgBox "GetEdgeLowSideofPoint:" & vbCrLf & _
'           "Root Point  = " & oPoint.x & ", " & oPoint.y & ", " & oPoint.z & vbCrLf & _
'           "Unit vector = " & oDirVec.x & ", " & oDirVec.y & ", " & oDirVec.z

    For nCount = 1 To oEdgeCol.count
        Set oEdge = oEdgeCol.Item(nCount)
        GetEndPoints oEdge, oStartPos, oEndPos
        Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
        Set oVec = oMidPos.Subtract(oPoint)

'        MsgBox "Edge No. " & nCount & vbCrLf & _
'               "Start Pnt = " & oStartPos.x & ", " & oStartPos.y & ", " & oStartPos.z & vbCrLf & _
'               "End Point = " & oEndPos.x & ", " & oEndPos.y & ", " & oEndPos.z & vbCrLf & _
'               "Mid Point = " & oMidPos.x & ", " & oMidPos.y & ", " & oMidPos.z & vbCrLf & _
'               "Mid to Root Vec = " & oVec.x & ", " & oVec.y & ", " & oVec.z

        If oVec.Dot(oDirVec) < 0 Then
            dDist = oMidPos.DistPt(oPoint)
            If dDist > dMaxDist Then
                Set GetEdgeLowSideofPoint = oEdge
                dMaxDist = dDist
            End If
        End If
    Next
End Function


Public Function GetEdgeHighSideofPoint(oEdgeCol As IJElements, oPoint As IJDPosition, oDirVec As IJDVector) As Object
    Dim nCount As Long
    Dim oEdge As IUnknown
    Dim dMaxDist As Double
    Dim dDist As Double
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oMidPos As IJDPosition
    Dim oVec As IJDVector

'    MsgBox "GetEdgeHighSideofPoint:" & vbCrLf & _
'           "Root Point  = " & oPoint.x & ", " & oPoint.y & ", " & oPoint.z & vbCrLf & _
'           "Unit vector = " & oDirVec.x & ", " & oDirVec.y & ", " & oDirVec.z

    For nCount = 1 To oEdgeCol.count
        Set oEdge = oEdgeCol.Item(nCount)
        GetEndPoints oEdge, oStartPos, oEndPos
        Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
        Set oVec = oMidPos.Subtract(oPoint)

'        MsgBox "Edge No. " & nCount & vbCrLf & _
'               "Start Pnt = " & oStartPos.x & ", " & oStartPos.y & ", " & oStartPos.z & vbCrLf & _
'               "End Point = " & oEndPos.x & ", " & oEndPos.y & ", " & oEndPos.z & vbCrLf & _
'               "Mid Point = " & oMidPos.x & ", " & oMidPos.y & ", " & oMidPos.z & vbCrLf & _
'               "Mid to Root Vec = " & oVec.x & ", " & oVec.y & ", " & oVec.z

        If oVec.Dot(oDirVec) > 0 Then
            dDist = oMidPos.DistPt(oPoint)
            If dDist > dMaxDist Then
                Set GetEdgeHighSideofPoint = oEdge
                dMaxDist = dDist
            End If
        End If
    Next
End Function

''Public Function GetDistBtwnEdgeAndPoint(oEdgeStartPt As IJDPosition, oEdgeEndPt As IJDPosition, oPoint As IJDPosition) As Double
''
''    Dim oMidX As Double
''    Dim oMidY As Double
''    Dim oMidZ As Double
''
''    Dim oMid As IJDPosition
''    Set oMid = DPosition
''
''    oMid.x = (oEdgeStartPt.x + oEdgeEndPt.x) / 2
''    oMid.y = (oEdgeStartPt.y + oEdgeEndPt.y) / 2
''    oMid.z = (oEdgeStartPt.z + oEdgeEndPt.z) / 2
''
''    GetDistBtwnEdgeAndPoint = oMid.DistPt(oPoint)
''
''End Function


Public Function GetEdgeMidPoint(oEdgeStartPt As IJDPosition, oEdgeEndPt As IJDPosition) As IJDPosition

    Dim oMidX As Double
    Dim oMidY As Double
    Dim oMidZ As Double

    Dim oMid As IJDPosition
    Set oMid = New DPosition

    oMid.x = (oEdgeStartPt.x + oEdgeEndPt.x) / 2
    oMid.y = (oEdgeStartPt.y + oEdgeEndPt.y) / 2
    oMid.z = (oEdgeStartPt.z + oEdgeEndPt.z) / 2

    Set GetEdgeMidPoint = oMid

    Set oMid = Nothing

End Function

Public Function FindBestEdgeNormalToGivenDir(oEdgeCol As IJElements, oDirVec As IJDVector) As Object
    Dim nCount As Long
    Dim oEdge As IUnknown
    Dim dMinDot As Double
    Dim dDot As Double
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oMidPos As IJDPosition
    Dim oVec As IJDVector

    dMinDot = 1
    For nCount = 1 To oEdgeCol.count
        Set oEdge = oEdgeCol.Item(nCount)
        GetEndPoints oEdge, oStartPos, oEndPos
        Set oVec = oEndPos.Subtract(oStartPos)
        oVec.Length = 1
        oDirVec.Length = 1
        dDot = Abs(oVec.Dot(oDirVec))
        If dDot < dMinDot Then
            Set FindBestEdgeNormalToGivenDir = oEdge
            dMinDot = dDot
        End If
    Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This method gets the edge that is best among the edges for given direction. Then it gets
'edge parallel to this best edge
'Where as the old method gets the first edge which is almost in the given direction
'(does not guaranteed to be best) and strives to make sure that the two edges are farthest from center
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEdgesAtGivenDirectionAndPlateNew2(oSurfaceBody As IJSurfaceBody, oEdgeElements As IJElements, strDirection As String) As IJElements
    Const METHOD = "GetEdgesAtGivenDirectionAndPlateNew2"
    On Error GoTo ErrorHandler

    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1

    Dim dNumEdges As Double
    dNumEdges = oEdgeElements.count

    Dim oXDirEdges As IJElements
    Set oXDirEdges = New JObjectCollection

    Dim oYDirEdges As IJElements
    Set oYDirEdges = New JObjectCollection

    Dim oZDirEdges As IJElements
    Set oZDirEdges = New JObjectCollection

    Dim oOutputEdges As IJElements
    Set oOutputEdges = New JObjectCollection

    Dim dDot1 As Double, dDot2 As Double, dDot3 As Double
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oDirVec As IJDVector
    Dim nCount As Long, nXDir As Integer, nYDir As Integer, nZDir As Integer
    Dim oEdge As IUnknown

    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    Dim eDir As enumDirection

    Dim oBestEdge As Object
    Dim oSecondBestEdge As Object
    
    Set oBestEdge = GetBestEdge(oEdgeElements, strDirection, oRootPoint)
    oEdgeElements.Remove oBestEdge
    
    Set oSecondBestEdge = GetBestOppositeEdge(oEdgeElements, oBestEdge, oRootPoint, strDirection)

    Dim oBCLDirVec As IJDVector
    If strDirection = "Transversal" Then
        Set oBCLDirVec = oXDir
    ElseIf strDirection = "Longitudinal" Then
        Set oBCLDirVec = oYDir
    ElseIf strDirection = "Waterline" Then '
        Set oBCLDirVec = oZDir
    End If
    
    Dim oMidPos As IJDPosition
    GetEndPoints oBestEdge, oStartPos, oEndPos
    Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
    Dim oTempVec As IJDVector
    Set oTempVec = oMidPos.Subtract(oRootPoint)
    If oTempVec.Dot(oBCLDirVec) < 0 Then
        oOutputEdges.Add oBestEdge
        oOutputEdges.Add oSecondBestEdge
    Else
        oOutputEdges.Add oSecondBestEdge
        oOutputEdges.Add oBestEdge
    End If
    
    Set GetEdgesAtGivenDirectionAndPlateNew2 = oOutputEdges

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function AreEdgesAdjacent(oEdge1 As IUnknown, oEdge2 As IUnknown) As Boolean
    Const METHOD = "AreEdgesAdjacent"
    On Error GoTo ErrorHandler
    
    Dim oLowStartPos As IJDPosition
    Dim oLowEndPos As IJDPosition
    
    Dim oHighStartPos As IJDPosition
    Dim oHighEndPos As IJDPosition
    
    GetEndPoints oEdge1, oLowStartPos, oLowEndPos
    GetEndPoints oEdge2, oHighStartPos, oHighEndPos
    
    'Compare the end points
    If (Abs(oLowStartPos.x - oHighEndPos.x) <= DISTANCETOLERANCE And Abs(oLowStartPos.y - oHighEndPos.y) <= DISTANCETOLERANCE And Abs(oLowStartPos.z - oHighEndPos.z) <= DISTANCETOLERANCE) Then
        AreEdgesAdjacent = True
    ElseIf (Abs(oLowEndPos.x - oHighStartPos.x) <= DISTANCETOLERANCE And Abs(oLowEndPos.y - oHighStartPos.y) <= DISTANCETOLERANCE And Abs(oLowEndPos.z - oHighStartPos.z) <= DISTANCETOLERANCE) Then
        AreEdgesAdjacent = True
    Else
        Exit Function
    End If
      
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetBestEdge(oEdgeElements As IJElements, strDirection As String, oRootPoint As IJDPosition) As Object
    Const METHOD = "GetBestEdge"
    On Error GoTo ErrorHandler
    
   'get template direction vector
    Dim oDirVec As IJDVector
    Set oDirVec = New DVector
    If strDirection = "Longitudinal" Then 'X - Direction
        oDirVec.Set 0, 1, 0
    ElseIf strDirection = "Transversal" Then 'Y - Direction
        oDirVec.Set 1, 0, 0
    Else 'Z - Direction
        oDirVec.Set 0, 0, 1
    End If
    oDirVec.Length = 1
    
    Dim oTempedge As Object
    Dim oTempEdgeMidPos As IJDPosition
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oVecToEdge As IJDVector
    
    Dim dMaxDotProduct As Double
    dMaxDotProduct = 0
    
    Dim oEdgesNotInGivenDir As IJElements
    Set oEdgesNotInGivenDir = New JObjectCollection
    For Each oTempedge In oEdgeElements
         
        If IsEdgeInGivenDirection(oTempedge, strDirection) Then
        GetEndPoints oTempedge, oStartPos, oEndPos
        Set oTempEdgeMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
            Set oVecToEdge = oTempEdgeMidPos.Subtract(oRootPoint)
            oVecToEdge.Length = 1
        
            If Abs(oVecToEdge.Dot(oDirVec)) > dMaxDotProduct Then
            Set GetBestEdge = oTempedge
                dMaxDotProduct = Abs(oVecToEdge.Dot(oDirVec))
            End If
        Else
           oEdgesNotInGivenDir.Add oTempedge
        End If
    Next
       
    If GetBestEdge Is Nothing Then
        'Rare case of plate where all edges are not in given direction
        'Get the best one of them
        Set GetBestEdge = GetEdgeLeastInGivenDirection(oEdgesNotInGivenDir, strDirection)
    End If
       
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetBestOppositeEdge(oEdgeElements As IJElements, oEdge As Object, oRootPoint As IJDPosition, strDirection As String) As Object
    Const METHOD = "GetBestOppositeEdge"
    On Error GoTo ErrorHandler
    
    Dim oTempedge As Object
    Dim oTempEdgeMidPos As IJDPosition
    Dim oBestEdgeMidPos As IJDPosition
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oVecToBestEdge As IJDVector
    Dim oVecToEdge As IJDVector
        
    'Get the best edge vector
    GetEndPoints oEdge, oStartPos, oEndPos
    Set oBestEdgeMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
    Set oVecToBestEdge = oRootPoint.Subtract(oBestEdgeMidPos)
    oVecToBestEdge.Length = 1
    
    Dim oEdgesNotInGivenDir As IJElements
    Set oEdgesNotInGivenDir = New JObjectCollection
    
    Dim dMinDotProduct As Double
    dMinDotProduct = 1
    For Each oTempedge In oEdgeElements
        GetEndPoints oTempedge, oStartPos, oEndPos
            If (Not (AreEdgesAdjacent(oEdge, oTempedge))) Then
            If (IsEdgeInGivenDirection(oTempedge, strDirection)) Then
                Set oTempEdgeMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
                Set oVecToEdge = oRootPoint.Subtract(oTempEdgeMidPos)
                oVecToEdge.Length = 1

                If oVecToBestEdge.Dot(oVecToEdge) < dMinDotProduct Then
                    Set GetBestOppositeEdge = oTempedge
                    dMinDotProduct = oVecToBestEdge.Dot(oVecToEdge)
                End If
            Else
                oEdgesNotInGivenDir.Add oTempedge
            End If
        End If
    Next

    'try on the rest of the edges
    If GetBestOppositeEdge Is Nothing Then
        For Each oTempedge In oEdgesNotInGivenDir
            GetEndPoints oTempedge, oStartPos, oEndPos
            Set oTempEdgeMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
            Set oVecToEdge = oRootPoint.Subtract(oTempEdgeMidPos)
            oVecToEdge.Length = 1

            If oVecToBestEdge.Dot(oVecToEdge) < dMinDotProduct Then
                Set GetBestOppositeEdge = oTempedge
                dMinDotProduct = oVecToBestEdge.Dot(oVecToEdge)
            End If
    Next
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function IsPointAtCurveEnd(oPoint As IJDPosition, oCurve As Object) As Boolean
    Const METHOD = "IsPointAtCurveEnd"
    On Error GoTo ErrorHandler
    
    Dim oStartPt As IJDPosition
    Dim oEndPt As IJDPosition
    
    GetEndPoints oCurve, oStartPt, oEndPt
    
    If IsEqualPoint(oStartPt, oPoint) Then
        IsPointAtCurveEnd = True
        Exit Function
    End If
    
    If IsEqualPoint(oEndPt, oPoint) Then
        IsPointAtCurveEnd = True
        Exit Function
    End If
        
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************************************************************************
' Method Name:    GetVertexNotOnEdgeForTrainglePlate
'
' Abstract:         This method returns the vertex of the triangle plate that is not on the given edge
'
' Inputs:           oTrianglePlateEdges -- the edges of the triangle plate
'                   oGivenEdge -- the given edge for check
'
' Output:           the vertex not on given edge
'
' Algorithm:
'                   1. get end points of the given edge
'                   2. get the vertex on plate edges that is not one of the above end points
'*********************************************************************************************************************************
Public Function GetVertexNotOnEdgeForTrainglePlate(oTrianglePlateEdges As IJElements, oGivenEdge As Object) As IJDPosition
    Const METHOD = "GetVertexNotOnEdgeForTrainglePlate"
    On Error GoTo ErrorHandler
    
    Dim oVertex1 As IJDPosition
    Dim oVertex2 As IJDPosition
    
    'get end points of the given edge
    GetEndPoints oGivenEdge, oVertex1, oVertex2
    
    'get the vertex on plate edges that is not one of the above end points
    Dim i As Long
    For i = 1 To oTrianglePlateEdges.count
        Dim oStartPos As IJDPosition
        Dim oEndPos As IJDPosition
        
        'get end points of the given edge
        GetEndPoints oTrianglePlateEdges.Item(i), oStartPos, oEndPos
        
        If Not (IsEqualPoint(oStartPos, oVertex1)) Then
            If Not (IsEqualPoint(oStartPos, oVertex2)) Then
                'startpos is vertex3
                Set GetVertexNotOnEdgeForTrainglePlate = oStartPos
                Exit For
            End If
        End If
    Next
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IsEdgeInGivenDirection(oEdge As Object, strDirection As String)
    Const METHOD = "IsEdgeInGivenDirection"
    On Error GoTo ErrorHandler
    
    Dim oXDir As IJDVector, oYDir As IJDVector, oZDir As IJDVector
    Set oXDir = New DVector
    oXDir.Set 1, 0, 0

    Set oYDir = New DVector
    oYDir.Set 0, 1, 0

    Set oZDir = New DVector
    oZDir.Set 0, 0, 1
    
    'create edge vector
    Dim oAlongEdgeVec As IJDVector
    Dim oStartPos As IJDPosition, oEndPosition As IJDPosition
    GetEndPoints oEdge, oStartPos, oEndPosition
    Set oAlongEdgeVec = oStartPos.Subtract(oEndPosition)
    oAlongEdgeVec.Length = 1
    
    Dim dDotX As Double, dDotY As Double, dDotZ As Double
    dDotX = Abs(oAlongEdgeVec.Dot(oXDir))
    dDotY = Abs(oAlongEdgeVec.Dot(oYDir))
    dDotZ = Abs(oAlongEdgeVec.Dot(oZDir))
    
    If strDirection = "Transversal" Then
        If (dDotY > dDotX) Or (dDotZ > dDotX) Then
            IsEdgeInGivenDirection = True
        Else
            IsEdgeInGivenDirection = False
        End If
    ElseIf strDirection = "Longitudinal" Then
        If (dDotX > dDotY) Or (dDotZ > dDotY) Then
            IsEdgeInGivenDirection = True
        Else
            IsEdgeInGivenDirection = False
        End If
    Else 'WaterLine
        If (dDotX > dDotZ) Or (dDotY > dDotZ) Then
            IsEdgeInGivenDirection = True
        Else
            IsEdgeInGivenDirection = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetEdgeLeastInGivenDirection(oEdges As IJElements, strDirection As String) As Object
    Const METHOD = "GetEdgeLeastInGivenDirection"
    On Error GoTo ErrorHandler
    
    Dim oDirVec As IJDVector
    Set oDirVec = New DVector
    If strDirection = "Transversal" Then
        oDirVec.Set 1, 0, 0
    ElseIf strDirection = "Longitudinal" Then
        oDirVec.Set 0, 1, 0
    Else 'WaterLine
        oDirVec.Set 0, 0, 1
    End If
    
    Dim dLeastDot As Double
    dLeastDot = 1
    Dim oEdge As Object
    For Each oEdge In oEdges
        'create edge vector
        Dim oAlongEdgeVec As IJDVector
        Dim oStartPos As IJDPosition, oEndPosition As IJDPosition
        GetEndPoints oEdge, oStartPos, oEndPosition
        Set oAlongEdgeVec = oStartPos.Subtract(oEndPosition)
        oAlongEdgeVec.Length = 1
        
        Dim dDot As Double
        dDot = Abs(oAlongEdgeVec.Dot(oDirVec))
        
        If dDot < dLeastDot Then
            dLeastDot = dDot
            Set GetEdgeLeastInGivenDirection = oEdge
        End If
    Next

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Get the edges on AftSide and ForeSide based on the root point and BCL direction(normal to Template direction)
'Among the AftSide edges, get the good aft edges(those are more in Template direction than along template normal)
'If these are zero, consider all aft edges as good aft edges
'Among these good aft edges, get the Best aft edge(edge vector to which is most parallel to BCL)
'Similarly, get the good Fore edges(those are more in Template direction than along template normal)
'Among these get the Best Fore edge( the edge vector to which is more parallel to vector to aft)
'NOTE:we carry out above calculations on 2D (show we project the edges)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEdgesAtGivenDirectionAndPlateNew3(oSurfaceBody As IJSurfaceBody, oEdgeElements As IJElements, strDirection As String) As IJElements
    Const METHOD = "GetEdgesAtGivenDirectionAndPlateNew3"
    On Error GoTo ErrorHandler

    Dim oOutputEdges As IJElements
    Set oOutputEdges = New JObjectCollection

    Dim oRootPoint As IJDPosition
    Dim oNormal As IJDVector
    
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oRootPoint, oNormal
    Set oTopoLocate = Nothing
    
    Dim eProjDir As enumDirection
    Dim oBCLDir As IJDVector
    Set oBCLDir = New DVector
    
    If strDirection = "Transversal" Then
        If Abs(oNormal.y) > Abs(oNormal.z) Then
            eProjDir = eYDir
        Else
            eProjDir = eZDir
        End If
        oBCLDir.Set 1, 0, 0
    ElseIf strDirection = "Longitudinal" Then
        If Abs(oNormal.x) > Abs(oNormal.z) Then
            eProjDir = eXDir
        Else
            eProjDir = eZDir
        End If
        oBCLDir.Set 0, 1, 0
    ElseIf strDirection = "Waterline" Then
        If Abs(oNormal.x) > Abs(oNormal.y) Then
            eProjDir = eXDir
        Else
            eProjDir = eYDir
        End If
        oBCLDir.Set 0, 0, 1
    End If
    
    'get low side(Aft) and high side(Fore) edges
    Dim oAftSideEdges As IJElements
    Dim oForeSideEdges As IJElements
    GetAftAndForeEdges oEdgeElements, eProjDir, oBCLDir, oRootPoint, oAftSideEdges, oForeSideEdges

    'Get the good edges on low side
    Dim oGoodAftEdges As IJElements
    Set oGoodAftEdges = GetEdgesNormalToGivenDirection(oAftSideEdges, eProjDir, oBCLDir)
    If oGoodAftEdges.count = 0 Then
       Set oGoodAftEdges = oAftSideEdges
    End If
    
    'Get the best Aft edge
    Dim oBestAftEdge As Object
    Set oBestAftEdge = GetEdgeMostNormalToGivenDir(oGoodAftEdges, eProjDir, oBCLDir, oRootPoint)
    
    ''Get its opposite edge - Opposite edge is one of the Highedges
    'It should not be adjacent
    RemoveEdgesAdjacentToEdge oForeSideEdges, oBestAftEdge
        
    If oForeSideEdges.count = 0 Then
        'Alledges - BestAftEdge - BestAftEdgeNeighbors
        oForeSideEdges.AddElements oEdgeElements
        oForeSideEdges.Remove oBestAftEdge
        RemoveEdgesAdjacentToEdge oForeSideEdges, oBestAftEdge
    End If
            
    'Get the good edges on Fore side
    Dim oGoodForeEdges As IJElements
    Set oGoodForeEdges = GetEdgesNormalToGivenDirection(oForeSideEdges, eProjDir, oBCLDir)
    If oGoodForeEdges.count = 0 Then
       Set oGoodForeEdges = oForeSideEdges
    End If
    
    'GetBestEdgeVector
    Dim oVecToBestAftEdge As IJDVector
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    GetEndPoints oBestAftEdge, oStartPos, oEndPos

    If eProjDir = eXDir Then 'project Points onto YZ plane passing origin point
        oStartPos.x = 0
        oEndPos.x = 0
        oRootPoint.x = 0
    ElseIf eProjDir = eYDir Then 'project Points onto XZ plane passing origin point
        oStartPos.y = 0
        oEndPos.y = 0
        oRootPoint.y = 0
    Else 'project Points onto XY plane passing origin point
        oStartPos.z = 0
        oEndPos.z = 0
        oRootPoint.z = 0
    End If

    Dim oMidPoint As IJDPosition
    Set oMidPoint = GetEdgeMidPoint(oStartPos, oEndPos)
    
    Dim oVecToEdge As IJDVector
    Set oVecToBestAftEdge = oRootPoint.Subtract(oMidPoint)
    oVecToBestAftEdge.Length = 1
    
    'get the best Fore side edge
    Dim oBestForeEdge As Object
    Set oBestForeEdge = GetEdgeMostNormalToGivenDir(oGoodForeEdges, eProjDir, oVecToBestAftEdge, oRootPoint)
           
    oOutputEdges.Add oBestAftEdge
    oOutputEdges.Add oBestForeEdge

    Set GetEdgesAtGivenDirectionAndPlateNew3 = oOutputEdges

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'From the given edges, get the edges that are on Aft side and those on the Fore side
Private Sub GetAftAndForeEdges(ByVal oEdges As IJElements, ByVal eProjDir As enumDirection, ByVal oBCLDir As IJDVector, ByVal oRootPoint As IJDPosition, ByRef oAftEdges As IJElements, ByRef oForeEdges As IJElements)
    Const METHOD = "GetAftAndForeEdges"
    On Error GoTo ErrorHandler
    
    Set oAftEdges = New JObjectCollection
    Set oForeEdges = New JObjectCollection
  
    'get the Aft and Fore side edges
    Dim nCount As Long
    For nCount = 1 To oEdges.count
        Dim oEdge As Object
        Set oEdge = oEdges.Item(nCount)
        
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        GetEndPoints oEdge, oStartPos, oEndPos

        If eProjDir = eXDir Then 'project Points onto YZ plane passing origin point
            oStartPos.x = 0
            oEndPos.x = 0
            oRootPoint.x = 0
        ElseIf eProjDir = eYDir Then 'project Points onto XZ plane passing origin point
            oStartPos.y = 0
            oEndPos.y = 0
            oRootPoint.y = 0
        Else 'project Points onto XY plane passing origin point
            oStartPos.z = 0
            oEndPos.z = 0
            oRootPoint.z = 0
        End If

        Dim oMidPoint As IJDPosition
        Set oMidPoint = GetEdgeMidPoint(oStartPos, oEndPos)
        
        Dim oVecToEdge As IJDVector
        Set oVecToEdge = oMidPoint.Subtract(oRootPoint)
        oVecToEdge.Length = 1

        Dim dDot As Double
        dDot = oVecToEdge.Dot(oBCLDir)
        
        If dDot < 0 Then
            oAftEdges.Add oEdge
        Else
            oForeEdges.Add oEdge
        End If
    Next
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'Get the edges which are more normal to given direction than along it(i.e., we check if they are in template direction or not
Private Function GetEdgesNormalToGivenDirection(ByVal oEdges As IJElements, ByVal eProjDir As enumDirection, ByVal oBCLDir As IJDVector) As IJElements
    Const METHOD = "GetEdgesNormalToGivenDirection"
    On Error GoTo ErrorHandler
    
    Set GetEdgesNormalToGivenDirection = New JObjectCollection
    
    Dim nCount As Long
    For nCount = 1 To oEdges.count
        Dim oEdge As Object
        Set oEdge = oEdges.Item(nCount)
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        GetEndPoints oEdge, oStartPos, oEndPos

        If eProjDir = eXDir Then 'project Points onto YZ plane passing origin point
            oStartPos.x = 0
            oEndPos.x = 0
        ElseIf eProjDir = eYDir Then 'project Points onto XZ plane passing origin point
            oStartPos.y = 0
            oEndPos.y = 0
        Else 'project Points onto XY plane passing origin point
            oStartPos.z = 0
            oEndPos.z = 0
        End If
    
        Dim oVecAlongEdge As IJDVector
        Set oVecAlongEdge = oEndPos.Subtract(oStartPos)
        oVecAlongEdge.Length = 1

        Dim dDot As Double
        dDot = Abs(oVecAlongEdge.Dot(oBCLDir))

        If dDot < COS45 Then
            'Along Template direction(normal to bcl)
            GetEdgesNormalToGivenDirection.Add oEdge
        End If
    Next nCount
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'Among the given edges, find the edge vector to which is most along the given direction
Private Function GetEdgeMostNormalToGivenDir(oEdges As IJElements, eProjDir As enumDirection, oBCLDir As IJDVector, ByVal oRootPoint As IJDPosition) As Object
    Const METHOD = "GetEdgeMostNormalToGivenDir"
    On Error GoTo ErrorHandler
    
    Dim nCount As Long
    Dim dMaxDot As Double
    dMaxDot = -1
    For nCount = 1 To oEdges.count
        Dim oEdge As Object
        Set oEdge = oEdges.Item(nCount)
        
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        GetEndPoints oEdge, oStartPos, oEndPos

        If eProjDir = eXDir Then 'project Points onto YZ plane passing origin point
            oStartPos.x = 0
            oEndPos.x = 0
            oRootPoint.x = 0
        ElseIf eProjDir = eYDir Then 'project Points onto XZ plane passing origin point
            oStartPos.y = 0
            oEndPos.y = 0
            oRootPoint.y = 0
        Else 'project Points onto XY plane passing origin point
            oStartPos.z = 0
            oEndPos.z = 0
            oRootPoint.z = 0
        End If

        Dim oMidPoint As IJDPosition
        Set oMidPoint = GetEdgeMidPoint(oStartPos, oEndPos)
        
        Dim oVecToEdge As IJDVector
        Set oVecToEdge = oRootPoint.Subtract(oMidPoint)
        oVecToEdge.Length = 1

        Dim dDot As Double
        dDot = Abs(oVecToEdge.Dot(oBCLDir))
        
        If dDot > dMaxDot Then
            dMaxDot = dDot
            Set GetEdgeMostNormalToGivenDir = oEdge
        End If
    Next
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub RemoveEdgesAdjacentToEdge(ByRef oEdges As IJElements, ByVal oGivenEdge As Object)
    Const METHOD = "RemoveEdgesAdjacentToEdge"
    On Error GoTo ErrorHandler
    
    Dim nCount As Long
    Dim oAdjacentEdges As IJElements
    Set oAdjacentEdges = New JObjectCollection
    For nCount = 1 To oEdges.count
        Dim oEdge As Object
        Set oEdge = oEdges.Item(nCount)
        
        If AreEdgesAdjacent(oGivenEdge, oEdge) Then
            oAdjacentEdges.Add oEdge
        End If
    Next
    oEdges.RemoveElements oAdjacentEdges
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
