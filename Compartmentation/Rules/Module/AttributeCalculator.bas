Attribute VB_Name = "AttributeCalculator"
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: AttributeCalculator
'
'  History:
'  AppaRao            31st Mar 2005             Creation
'***************************************************************************
Option Explicit

Private Const MODULE = "AttributeCalculator "

Private Function GetResourceMgr() As IUnknown
Const METHOD = "GetResourceMgr"
    On Error GoTo ErrorHandler
    
    Dim JContext            As IJContext
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim strModelDBID        As String
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    'get the model resource manager
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set GetResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set JContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
   
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Private Function GetBoundingFrameNames(oEntity As Object, oAxis As AxisType, bMinimum As Boolean) As String
Const METHOD = "GetBoundingFrameNames"
On Error GoTo ErrorHandler
   
    Dim oFrameMiddleHelper              As SPGMiddleHelper
    Dim oRange                          As IJRangeAlias
    Dim oFrameColl                      As IJElements
    Dim oNamedItem                      As IJNamedItem
    Dim oFirstFrame                     As ISPGNavigate
    Dim oLastFrame                      As ISPGNavigate
    Dim oBoundingFrame                  As IJNamedItem
    Dim oBox                            As GBox
    Dim oIntFrameColl                   As IJElements

    Set oRange = oEntity
    Set oFrameMiddleHelper = New SPGMiddleHelper
    
    Dim oPos1 As IJDPosition
    
    Dim oPos2 As IJDPosition
    
    
    Set oPos1 = New DPosition
    
    Set oPos2 = New DPosition
    
    'get the collection of frames
    oFrameMiddleHelper.EnumPlanesInRange GetResourceMgr, oRange, Nothing, "", oAxis, True, 4, oFrameColl
    
    'if no frames are found give the frame name as blank
    If oFrameColl.Count = 0 Then
        GetBoundingFrameNames = ""
        Exit Function
    End If
    
    OrderPlanesByPosition oFrameColl, oAxis

    If bMinimum Then 'First frame
        'get the first frame
        Set oFirstFrame = oFrameColl.Item(1)
        
        If IsPlaneTouchingRangeBox(oRange, oFirstFrame) = True Then
            Set oBoundingFrame = oFirstFrame
        Else
            oFirstFrame.GetReference Previous, NestingLevelType.Any, oBoundingFrame
            
            'get the other available planes in between the reference plane and th eprevious plane
            oBox = oRange.GetRange
                        
            oPos1.x = -10000
            oPos1.y = -10000
            oPos1.z = -10000
            
            oPos2.x = oBox.m_low.x
            oPos2.y = oBox.m_low.y
            oPos2.z = oBox.m_low.z
           
            Set oIntFrameColl = Nothing
            
            oFrameMiddleHelper.EnumPlanesInRangeEx GetResourceMgr, oPos1, oPos2, oAxis, True, 4, oIntFrameColl
           

            If oIntFrameColl.Count > 0 Then
                OrderPlanesByPosition oIntFrameColl, oAxis
                Set oBoundingFrame = oIntFrameColl.Item(oIntFrameColl.Count)
            End If
            
        End If
        
        If oBoundingFrame Is Nothing Then
            Set oBoundingFrame = oFirstFrame
        End If
        
    Else  'last frame
        'get the last frame
        Set oLastFrame = oFrameColl.Item(oFrameColl.Count)
        
        If IsPlaneTouchingRangeBox(oRange, oLastFrame) = True Then
            Set oBoundingFrame = oLastFrame
        Else
            oLastFrame.GetReference ReferenceType.Next, NestingLevelType.Any, oBoundingFrame
            
            oBox = oRange.GetRange
                        
            oPos1.x = oBox.m_high.x
            oPos1.y = oBox.m_high.y
            oPos1.z = oBox.m_high.z
            
            oPos2.x = 10000
            oPos2.y = 10000
            oPos2.z = 10000
            
            Set oIntFrameColl = Nothing
            
            oFrameMiddleHelper.EnumPlanesInRangeEx GetResourceMgr, oPos1, oPos2, oAxis, True, 4, oIntFrameColl
                        
            If oIntFrameColl.Count > 0 Then
                OrderPlanesByPosition oIntFrameColl, oAxis
                Set oBoundingFrame = oIntFrameColl.Item(1)
            End If

        End If
        
        If oBoundingFrame Is Nothing Then
            Set oBoundingFrame = oLastFrame
        End If
        
    End If
    
    GetBoundingFrameNames = oBoundingFrame.Name
   
    Set oFrameMiddleHelper = Nothing
    Set oRange = Nothing
    Set oFrameColl = Nothing
    Set oNamedItem = Nothing
    Set oFirstFrame = Nothing
    Set oLastFrame = Nothing
    Set oBoundingFrame = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_FRAMES_FAILED_NAMES)
End Function

'Returns True if the plane just touches the range. Returns false if it intersects the range.
Private Function IsPlaneTouchingRangeBox(oRange As IJRangeAlias, oFrame As IJPlaneFacelet) As Boolean
Const METHOD = "IsPlaneTouchingRangeBox"
On Error GoTo ErrorHandler
    
    Dim gRangeBox           As GBox
    Dim oPosition(1 To 8)   As IJDPosition
    Dim i                   As Long
    Dim oRootPt             As IJDPosition
    Dim x                   As Double
    Dim y                   As Double
    Dim z                   As Double
    Dim oNormal             As IJDVector
    Dim oNewVector          As IJDVector
    Dim direction           As Double
    
    gRangeBox = oRange.GetRange
    IsPlaneTouchingRangeBox = True
    
    For i = 1 To 8
        Set oPosition(i) = New DPosition
    Next i
    
    oPosition(1).Set gRangeBox.m_low.x, gRangeBox.m_low.y, gRangeBox.m_low.z
    oPosition(2).Set gRangeBox.m_low.x, gRangeBox.m_low.y, gRangeBox.m_high.z
    oPosition(3).Set gRangeBox.m_low.x, gRangeBox.m_high.y, gRangeBox.m_low.z
    oPosition(4).Set gRangeBox.m_low.x, gRangeBox.m_high.y, gRangeBox.m_high.z
    oPosition(5).Set gRangeBox.m_high.x, gRangeBox.m_low.y, gRangeBox.m_low.z
    oPosition(6).Set gRangeBox.m_high.x, gRangeBox.m_low.y, gRangeBox.m_high.z
    oPosition(7).Set gRangeBox.m_high.x, gRangeBox.m_high.y, gRangeBox.m_low.z
    oPosition(8).Set gRangeBox.m_high.x, gRangeBox.m_high.y, gRangeBox.m_high.z
    
    oFrame.GetRootPoint x, y, z
    
    Set oRootPt = New DPosition
    oRootPt.Set x, y, z
    
    Set oNormal = New DVector
    oFrame.GetNormal x, y, z
    oNormal.Set x, y, z
    
    Set oNewVector = oPosition(1).Subtract(oRootPt)
    direction = oNewVector.Dot(oNormal)
    
    For i = 2 To 8
        Set oNewVector = oPosition(i).Subtract(oRootPt)
        If ((oNewVector.Dot(oNormal)) * direction < 0) Then
            IsPlaneTouchingRangeBox = False
            Exit For
        End If
    Next i
    
    For i = 1 To 8
        Set oPosition(i) = Nothing
    Next i

Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Public Function GetReferencePosition(oRange As IJRangeAlias) As ReferencePosition
Const METHOD = "GetReferencePosition"
On Error GoTo ErrorHandler
    
    Dim gRangeBox           As GBox
    
    gRangeBox = oRange.GetRange
    GetReferencePosition = rpUndefined

    If gRangeBox.m_low.y >= 0 And gRangeBox.m_high.y >= 0 Then
        GetReferencePosition = ReferencePosition.rpPortside
    ElseIf gRangeBox.m_low.y <= 0 And gRangeBox.m_high.y <= 0 Then
        GetReferencePosition = ReferencePosition.rpStarboard
    ElseIf gRangeBox.m_low.x <= 0 And gRangeBox.m_high.x <= 0 Then
        GetReferencePosition = ReferencePosition.rpAft
    ElseIf gRangeBox.m_low.x >= 0 And gRangeBox.m_high.x >= 0 Then
        GetReferencePosition = ReferencePosition.rpFore
    ElseIf gRangeBox.m_low.z <= 0 And gRangeBox.m_high.z <= 0 Then
        GetReferencePosition = ReferencePosition.rpBelow
    ElseIf gRangeBox.m_low.z >= 0 And gRangeBox.m_high.z >= 0 Then
        GetReferencePosition = ReferencePosition.rpAbove
    End If
    

Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Function GetVolumeMoulded(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetVolumeMoulded"
On Error GoTo ErrorHandler

    Dim oPOM                As IJDPOM
    Dim oMiddleCmd          As IJADOMiddleCommand
    Dim vConnName           As Variant
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    Dim strConnectMiddle    As String
    Dim strQuery            As String
    Dim oMnkrEnum           As IJDEnumMoniker
    Dim vItem               As Variant
    Dim oObject             As Object
    
    Set oMiddleCmd = New JMiddleCommand
    oMiddleCmd.Prepared = False
    oMiddleCmd.QueryLanguage = LANGUAGE_SQL
        
    strConnectMiddle = "ConnectMiddle"
    Set oContext = GetJContext()
    
    Set oAccessMiddle = oContext.GetService(strConnectMiddle)
    Set oPOM = oAccessMiddle.GetResourceManagerFromType("Model")
    
    vConnName = oPOM.DatabaseID
    oMiddleCmd.ActiveConnection = vConnName
    
    Dim oRange As IJRangeAlias
    Dim oLx As Double, oLy As Double, oLz As Double
    Dim oHx As Double, oHy As Double, oHz As Double
        
    Set oRange = oCompartEntity
    
    oLx = oRange.GetRange.m_low.x
    oLy = oRange.GetRange.m_low.y
    oLz = oRange.GetRange.m_low.z
    oHx = oRange.GetRange.m_high.x
    oHy = oRange.GetRange.m_high.y
    oHz = oRange.GetRange.m_high.z
    
    strQuery = "select oid from CoreSpatialIndex where xmin >" & oLx & " and ymin >" & oLy & "and zmin >" & oLz & _
                "and xmax <" & oHx & "and ymax <" & oHy & "and zmax <" & oHz & ""
    
    oMiddleCmd.CommandText = strQuery
    Set oMnkrEnum = oMiddleCmd.SelectObjects
    
    If Not oMnkrEnum Is Nothing Then
        For Each vItem In oMnkrEnum
            Set oObject = oPOM.GetObject(vItem)
'            If TypeOf oObject Is IJShpStrPart Then
'
'            End If
        Next vItem
    End If
    GetVolumeMoulded = 100
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_VOLUMEMOULDED)
End Function

Public Function GetWallArea(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetWallArea"
On Error GoTo ErrorHandler

    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetWallArea = oCompartAttribute.SurfaceArea
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
   Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_SURFACEAREA)
End Function

Public Function GetWallLength(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetWallLength"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetWallLength = 100 'oCompartAttribute.SurfaceArea
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_WALLLENGTH)
End Function
Public Function GetDeckHeight(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetDeckHeight"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetDeckHeight = 100 'oCompartAttribute.SurfaceArea
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_DECKHEIGHT)
End Function
Public Function GetSideWallArea(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetSideWallArea"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetSideWallArea = oCompartAttribute.SurfaceArea
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_SURFACEAREA)
End Function
Public Function GetBottomArea(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetBottomArea"
On Error GoTo ErrorHandler
    
        Dim oCompartGeom As Object
        Dim oCompartment As CompartAttributeHelper.Compartment
        Dim oGTTool As IJDTopologyToolBox
        Dim oBodies As IJElements
        Dim i As Long
        Dim dArea As Double
        Dim du As Double, dv As Double, dnx As Double, dny As Double, dnz As Double
        Dim osurface As IJSurface

        Set oCompartment = New CompartAttributeHelper.Compartment

        Set oCompartGeom = oCompartment.CompartGeometry(oCompartEntity)

        Set oBodies = New JObjectCollection

        Set oGTTool = New DGeomOpsToolBox

        If Not oCompartGeom Is Nothing Then
            oGTTool.ExplodeSurfaceBodyByFaces Nothing, oCompartGeom, oBodies
        End If

        For i = 1 To oBodies.Count

            Set osurface = oBodies.Item(i)

            osurface.Normal du, dv, dnx, dny, dnz

            If ((dnx = 0) And (dny = 0) And (dnz = -1)) Then
                osurface.Area dArea, 0.000001
                Exit For
            End If
        Next i

        If dArea = 0 Then
            Dim oCompartAttribute As IJCompartAttributes

            Set oCompartAttribute = oCompartEntity

            dArea = oCompartAttribute.SurfaceArea
        End If

        GetBottomArea = dArea
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_SURFACEAREA)
End Function

Public Function GetBottomCGX(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetBottomCGX"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    GetBottomCGX = oCompartAttribute.CogX
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_COG)
End Function

Public Function GetBottomXGY(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetBottomXGY"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetBottomXGY = oCompartAttribute.CogY
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_COG)
End Function

Public Function GetBottomXGZ(oCompartEntity As IJCompartEntity) As Double
Const METHOD = "GetBottomXGZ"
On Error GoTo ErrorHandler
    Dim oCompartAttribute As IJCompartAttributes
    
    'get CompartAttributes from CompartEntity
    Set oCompartAttribute = oCompartEntity
    
    GetBottomXGZ = oCompartAttribute.CogZ
    
    Set oCompartAttribute = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_STRUCTURE_FAILED_COG)
End Function

Public Function GetLongitudinalMin(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetLongitudinalMin"
On Error GoTo ErrorHandler

    GetLongitudinalMin = GetBoundingFrameNames(oCompartEntity, y, True)
    
Exit Function
ErrorHandler:
   Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Function GetLongitudinalMax(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetLongitudinalMax"
On Error GoTo ErrorHandler
   
    GetLongitudinalMax = GetBoundingFrameNames(oCompartEntity, y, False)
    
Exit Function
ErrorHandler:
   Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Function GetDeckMin(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetDeckMin"
On Error GoTo ErrorHandler

    GetDeckMin = GetBoundingFrameNames(oCompartEntity, z, True)
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Public Function GetDeckMax(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetDeckMax"
On Error GoTo ErrorHandler

    GetDeckMax = GetBoundingFrameNames(oCompartEntity, z, False)
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Function GetTransversalMin(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetTransversalMin"
On Error GoTo ErrorHandler

    GetTransversalMin = GetBoundingFrameNames(oCompartEntity, x, True)
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Function GetTransversalMax(oCompartEntity As IJCompartEntity) As String
Const METHOD = "GetTransversalMax"
        
    GetTransversalMax = GetBoundingFrameNames(oCompartEntity, x, False)
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Public Function GetMinTightness(oCompartEntity As IJCompartEntity) As StructPlateTightness
Const METHOD = "GetMinTightness"
On Error GoTo ErrorHandler
    
    Dim obj                     As Object
    Dim oPlate                  As IJPlate
    Dim lCount                  As Long
    Dim MinTightness            As StructPlateTightness
    Dim bNotProperBounded       As Boolean
    Dim oBoundingCollObj        As Object
    Dim oBoundingColl           As IJDObjectCollection
    
    'Get the Bounding Faces
    Set oBoundingCollObj = oCompartEntity.GetBoundingFaces

    If oBoundingCollObj Is Nothing Then
        GetMinTightness = UnSpecifiedTightness
        Exit Function
    End If
        
    Set oBoundingColl = oBoundingCollObj
    
    lCount = 1
            
   'Iterate through each Plate and get the MinTightness
     For Each obj In oBoundingCollObj
    
        'Set Obj = oBoundingColl.Item(lCount)
        
        If TypeOf obj Is IJPlate Then   'Get the Tightness
        
            Set oPlate = obj
            If Not oPlate Is Nothing Then
                If lCount = 1 Then
                    MinTightness = oPlate.Tightness
                ElseIf MinTightness > oPlate.Tightness Then
                    MinTightness = oPlate.Tightness
                End If
                lCount = lCount + 1
            End If
            
        Else  'If it is not bounded by plate part,
        
            bNotProperBounded = True
            Exit For
            
        End If
        
        Set oPlate = Nothing
        Set obj = Nothing
        
     Next
        
    If bNotProperBounded = False Then
        GetMinTightness = MinTightness  'Return the Minimum tightness
    Else
        GetMinTightness = UnSpecifiedTightness 'Do not Assign any tightness, as these are not connnected with Plates
    End If
       
           
    Set oBoundingColl = Nothing
    Set oBoundingCollObj = Nothing
        
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Public Function FramesExistsinModel() As Boolean
Const METHOD = "FramesExistsinModel"
On Error GoTo ErrorHandler

    FramesExistsinModel = True
    
    Dim oPOM            As IJDPOM
    Dim oMiddleCmd      As IJADOMiddleCommand
    Dim oMnkrEnum       As IJDEnumMoniker
    Dim vItem           As Variant
    Dim vConnName       As Variant
    
    ' get the model POM
    Set oPOM = GetPOM("Model")
    
    Set oMiddleCmd = New JMiddleCommand
    
    oMiddleCmd.Prepared = False
    oMiddleCmd.QueryLanguage = LANGUAGE_SQL
    
    vConnName = oPOM.DatabaseID
    oMiddleCmd.ActiveConnection = vConnName
    
    oMiddleCmd.CommandText = "select oid from GRDSYSSPGCoordinateSystem union select oid from GRDSYSSPGShipCoordinateSystem"

    Set oMnkrEnum = oMiddleCmd.SelectObjects
    
    If oMnkrEnum Is Nothing Then
        FramesExistsinModel = False
    End If
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)

End Function

Private Function GetPOM(strDbType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    Dim strConnectMiddle    As String
    
    strConnectMiddle = "ConnectMiddle"
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService(strConnectMiddle)
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Exit Function
ErrHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_FRAMES_FAILED_NAMES)
End Function

Private Sub OrderPlanesByPosition(oPlanes As IJElements, oAxis As AxisType)
Const METHOD = "OrderPlanesByPosition"
On Error GoTo ErrorHandler
    
    Dim oGridData1      As IJPlaneFacelet
    Dim oGridData2      As IJPlaneFacelet
    Dim i               As Long
    Dim j               As Long
    Dim oTempObject     As Object
    Dim oElemsArray()   As Object
    Dim x1                   As Double
    Dim y1                   As Double
    Dim z1                   As Double
    
    Dim x2                  As Double
    Dim y2                   As Double
    Dim z2                   As Double
    
    ReDim oElemsArray(1 To oPlanes.Count) As Object
    
    For i = 1 To oPlanes.Count
        Set oElemsArray(i) = oPlanes.Item(i)
    Next i
    
    For i = 1 To oPlanes.Count - 1
        For j = 1 To oPlanes.Count - 1
            Set oGridData1 = oElemsArray(j)
            Set oGridData2 = oElemsArray(j + 1)
            oGridData1.GetRootPoint x1, y1, z1
            oGridData2.GetRootPoint x2, y2, z2
            
            If oAxis = x Then
                If x1 > x2 Then
                    Set oTempObject = oElemsArray(j + 1)
                    Set oElemsArray(j + 1) = oElemsArray(j)
                    Set oElemsArray(j) = oTempObject
                End If
            ElseIf oAxis = y Then
                If y1 > y2 Then
                    Set oTempObject = oElemsArray(j + 1)
                    Set oElemsArray(j + 1) = oElemsArray(j)
                    Set oElemsArray(j) = oTempObject
                End If
            ElseIf oAxis = z Then
            
                If z1 > z2 Then
                    Set oTempObject = oElemsArray(j + 1)
                    Set oElemsArray(j + 1) = oElemsArray(j)
                    Set oElemsArray(j) = oTempObject
                End If
            End If
            
            
        Next j
    Next i
    
    oPlanes.Clear
    
    For i = 1 To UBound(oElemsArray)
        oPlanes.Add oElemsArray(i)
    Next i
    
    Set oGridData1 = Nothing
    Set oGridData2 = Nothing
    Set oTempObject = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_FRAMES_FAILED_NAMES)
End Sub
