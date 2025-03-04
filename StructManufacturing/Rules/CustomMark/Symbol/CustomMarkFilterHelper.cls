VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomFilterHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project:
'
'File:
'
'Abstract: implementation of skid mark methods
'
'Author:
'
'Revision:
'*******************************************************************

Option Explicit

'********************************************************************
' Routine: IsConnectedToPlateSystem
'
' Abstract:
'       IsConnectedToPlateSystem = 0 if the Plate/Profile Sys is not connected to Plate System
'       IsConnectedToPlateSystem = 1 if the Plate/Profile Sys is connected to Plate System
'------------------------------------------------------------------------------'
Public Function IsConnectedToPlateSystem(oElem As Object, vArg As Variant) As Long
    
    IsConnectedToPlateSystem = 0
    
    Dim oElements As JCmnShp_CollectionAlias
    Set oElements = vArg
        
    'if no object is sent as an argument then return 0
    If oElements.Count < 1 Then
        Exit Function
    End If
    
     'Determine type of 1st argument object
    If TypeOf oElem Is IJStiffenerSystem Then
    
        ' If the second input is profile system then get the parent plate system
        ' of the input stiffener system and check with Input plate system
        Dim oStiffener As IJStiffener
        Set oStiffener = oElem
        
        If oElements.Item(1) Is oStiffener.PlateSystem Then
            IsConnectedToPlateSystem = 1
        End If
        
    ElseIf TypeOf oElem Is IJPlateSystem Then
    
        ' If the plate system(2nd input) selected is same as first input plate system then exit
        If oElements.Item(1) Is oElem Then
            Exit Function
        End If
        
        ' Determine the type of plate system selected(2nd input), If it is hull then exit
        Dim oPlateSystem As IJPlate
        Set oPlateSystem = oElem
    
        Dim ePlateType As StructPlateType
        ePlateType = oPlateSystem.plateType
        
        If ePlateType = Hull Then
            Exit Function
        End If
        
        ' Get the Root logical connection
        Dim ConnectionsColl As Collection
        Dim ConnectedObjColl As Collection
    
        Dim oStructConnectable As IJStructConnectable
        Set oStructConnectable = oElements.Item(1) 'First input
    
        ' Get all objects connected to plate with a logical connection
        Dim ConnectedObjIEUK As IEnumUnknown
        Dim ConnectionsIEUK As IEnumUnknown
        oStructConnectable.enumConnectedObjects2 ConnectionLogical, ConnectedObjIEUK, ConnectionsIEUK
        
        Dim oConvertUtils As New CONVERTUTILITIESLib.CCollectionConversions
        oConvertUtils.CreateVBCollectionFromIEnumUnknown ConnectedObjIEUK, ConnectedObjColl
        oConvertUtils.CreateVBCollectionFromIEnumUnknown ConnectionsIEUK, ConnectionsColl
        
        If ConnectionsColl.Count = 0 Or ConnectedObjColl.Count = 0 Then
            Exit Function
        End If
            
        Set oConvertUtils = Nothing
        Set ConnectedObjIEUK = Nothing
        Set ConnectionsIEUK = Nothing
    
        Dim oRootLogicalConnection As Object
        Dim i As Long
        For i = 1 To ConnectedObjColl.Count
            If ConnectedObjColl.Item(i) Is oElem Then
                Set oRootLogicalConnection = ConnectionsColl.Item(i)
                Exit For
            End If
        Next
        
        If oRootLogicalConnection Is Nothing Then
            Exit Function
        End If
        
        ' Get one of the PC from Root logical connection
        Dim oPhysicalConn As Object
        GetPhysicalConnection oRootLogicalConnection, oPhysicalConn
        
        If oPhysicalConn Is Nothing Then
            Exit Function
        End If
        
        Dim bIsCrossOfTee As Boolean
        Dim oConnType As ContourConnectionType
        Dim oConnection As IJAppConnection
        
        Set oConnection = oPhysicalConn
        
        Dim oPortElements As IJElements
        oConnection.enumPorts oPortElements
        
        Dim oPort As IJPort
        Set oPort = oPortElements.Item(1)
        
        Dim oPlatePart As IJPlatePart
        Set oPlatePart = oPort.Connectable
        
        Dim oStructDetailHelper As IJStructDetailHelper
        Set oStructDetailHelper = New StructDetailHelper
        
        Dim oRootPlateSys As IJSystem
        oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oRootPlateSys, True
        
        If oRootPlateSys Is oElements.Item(1) Then
            Set oPort = Nothing
            Set oPort = oPortElements.Item(2)
            
            Set oPlatePart = oPort.Connectable
        End If
        
        If oPlatePart Is Nothing Then
            Exit Function
        End If
        
        Dim oSDPartSupport As GSCADSDPartSupport.PartSupport
        Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
        Set oSDPartSupport.Part = oPlatePart
        
        ' Check if the connection type is TEE and bIsCrossOfTee = False
        oSDPartSupport.GetConnectionTypeForContour oConnection, _
                                                   oConnType, _
                                                   bIsCrossOfTee

        If oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsCrossOfTee = False Then
            IsConnectedToPlateSystem = 1
            Exit Function
        End If
        
    End If
    
End Function

'********************************************************************
' Routine: GetPhysicalConnection
'
' Abstract: Gets one of the the PC associated with LC
'********************************************************************
Public Function GetPhysicalConnection(oLCObject As IJDesignParent, ByRef oPCObject As Object)
    On Error GoTo ErrorHandler
    
    Dim lCChildren    As IJDObjectCollection
    oLCObject.GetChildren lCChildren
    
    Dim lcChildObject As Object
    
    For Each lcChildObject In lCChildren
        If TypeName(lcChildObject) = "IJStructPhysicalConnection" Then
            Set oPCObject = lcChildObject
            Exit For
        Else
            Dim oDesignparent As IJDesignParent
            
            On Error Resume Next
            Set oDesignparent = lcChildObject
            
            If Not oDesignparent Is Nothing Then
                GetPhysicalConnection lcChildObject, oPCObject
                If Not oPCObject Is Nothing Then
                    Exit For
                End If
            End If
        End If
        
    Next
    
    Set lcChildObject = Nothing
    Set lCChildren = Nothing

    Exit Function
    
ErrorHandler:
'    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************
' Routine: GetPlatePartFromPlate
'
' Abstract: Gets the Child PlatePart of the Plate system given
'********************************************************************
Public Function GetPartFromSystem(oInputSys As Object) As Object
    Const METHOD = "GetPartFromSystem"
    On Error GoTo ErrorHandler

    Dim colChildren As IJDTargetObjectCol
    Dim oMFSystem As IJSystem
    Dim oPart As Object
    
    'Get the children from the plate system
    Set oMFSystem = oInputSys
    Set colChildren = oMFSystem.GetChildren
    
    'Get the child which supports IJPlatePart/IJStiffenerPart
    Dim i As Integer
    
    For i = 1 To colChildren.Count
        If TypeOf colChildren.Item(i) Is IJPlatePart Or TypeOf colChildren.Item(i) Is IJStiffenerPart Then
            Set oPart = colChildren.Item(i)
            Exit For
        Else
            Dim oLeafSystem As IJSystem
            Set oLeafSystem = colChildren.Item(i)
            If Not oLeafSystem Is Nothing Then
                Dim oPartColl As IJDTargetObjectCol
                Set oPartColl = oLeafSystem.GetChildren
                On Error Resume Next
                Set oPart = oPartColl.Item(1)
                If Not oPart Is Nothing Then
                    If TypeOf oPart Is IJPlatePart Or TypeOf oPart Is IJStiffenerPart Then
                        Exit For
                    End If
                Else
                    On Error GoTo ErrorHandler
                End If
             End If
        End If
    Next i
  
    
    Set GetPartFromSystem = oPart
    
    'Cleanup
    Set colChildren = Nothing
    Set oMFSystem = Nothing
    Set oPart = Nothing
    
    Exit Function
ErrorHandler:
'    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'------------------------------------------------------------------------------'
' IsFrameInstersectingPlate
'       IsFrameInstersectingPlate = 0 if the plane does not intersect the plate
'       IsFrameInstersectingPlate = 1 if the plane intersects the plate
'------------------------------------------------------------------------------'
Public Function IsFrameInstersectingPlate(oElem As Object, vArg As Variant) As Long
    
    'check if the selected object is a frame
    'if frame, check if it intersects the plate passed in vArg.
    IsFrameInstersectingPlate = 0
    
    If Not TypeOf oElem Is IHFrame Then Exit Function
    
    Dim oElements As JCmnShp_CollectionAlias
    Set oElements = vArg
    
    'if no object is sent as an argument then return 0
    If oElements.Count < 1 Then
        Exit Function
    End If
    
    'The first element of oElements should be a plate
    Dim oPlate As IJPlate
    If TypeOf oElements.Item(1) Is IJPlate Then
        Set oPlate = oElements.Item(1)
        'check if oPlate and oelem intersect
        'first object is the frame
        

        
        Dim oMfgGeomHelper As IJMfgGeomHelper
        Set oMfgGeomHelper = New MfgGeomHelper
        Dim oOutputCurve As Object
        Dim oStartPoint As Object
        Dim oEndPoint As Object
        oMfgGeomHelper.IntersectSurfaceWithPlane oPlate, oElem, oOutputCurve, oStartPoint, oEndPoint
        Set oMfgGeomHelper = Nothing
        If Not oOutputCurve Is Nothing Then
            IsFrameInstersectingPlate = 1
        Else
           IsFrameInstersectingPlate = 0
        End If
    End If
    
End Function

'------------------------------------------------------------------------------'
' IsFrameInstersectingPlate
'       IsFromSameFrameSystemAndAxis = 0 if the planes do not belong to the same axis and the same frame system
'       IsFromSameFrameSystemAndAxis = 1 if the planes belong to the same axis and the same frame system
'------------------------------------------------------------------------------'
Public Function IsFromSameFrameSystemAndAxis(oElem As Object, vArg As Variant) As Long

    IsFromSameFrameSystemAndAxis = 0
    
    'check if the selected object is a frame
    If Not TypeOf oElem Is IHFrame Then Exit Function

    Dim oElements As JCmnShp_CollectionAlias
    Set oElements = vArg
    
    'if no object is sent as an argument then return 0
    'The first object of oElements should be a  plate
    'The second object of oElements should be an IHFrame
    If oElements.Count < 2 Then
        Exit Function
    End If
    
    Dim oFrame1 As IHFrame
    If TypeOf oElements.Item(2) Is IHFrame Then
        Set oFrame1 = oElements.Item(2)
        'check if oFrame and oelem are from the same axis and the same frame system
        
        Dim oFrame2 As IHFrame
        Set oFrame2 = oElem
        
        If oFrame1 Is oFrame2 Then
            IsFromSameFrameSystemAndAxis = 0
            Exit Function
        End If
        
        Dim oFrameAxis1 As IHFrameAxis
        Set oFrameAxis1 = oFrame1.FrameAxis
        
        Dim oFrameAxis2 As IHFrameAxis
        Set oFrameAxis2 = oFrame2.FrameAxis
        
        Dim oFrameSystem1 As IHFrameSystem
        Set oFrameSystem1 = oFrameAxis1.FrameSystem
        
        Dim oFrameSystem2 As IHFrameSystem
        Set oFrameSystem2 = oFrameAxis2.FrameSystem
        
        'if the frames are from the same frame system and same axis, return 1
        If oFrameAxis1 Is oFrameAxis2 And oFrameSystem2 Is oFrameSystem1 Then
            IsFromSameFrameSystemAndAxis = 1
        End If
    End If
    
End Function

'------------------------------------------------------------------------------'
' IsLCOrSeamOrFrame
'       IsLCOrSeamOrFrame = 0 if the object does not intersects the plate, and the boundaries.
'       IsLCOrSeamOrFrame = 1 if the object intersects the plate, and the boundaries.
'------------------------------------------------------------------------------'
Public Function IsLCOrSeamOrFrame(oElem As Object, vArg As Variant) As Long

    IsLCOrSeamOrFrame = 0

    Dim lLCOrSeamOrFrame As Long
    Dim lIntersectWithPlate As Long
    Dim lIntersectWithFrames As Long
    lLCOrSeamOrFrame = 0
    lIntersectWithPlate = 0
    lIntersectWithFrames = 0
    
    'check if the object selected is a Logical connection or a seam or a frame
    
    'Check for Logical connection
    If TypeOf oElem Is IJAppConnectionType Then
        Dim oAppConnectionType As IJAppConnectionType
        Set oAppConnectionType = oElem
        If oAppConnectionType.Type = ConnectionLogical Then
        
            Dim oWireBody As IJWireBody
            On Error Resume Next
            Set oWireBody = oAppConnectionType
            If Not oWireBody Is Nothing Then
                lLCOrSeamOrFrame = 1
            End If

        End If
    ElseIf TypeOf oElem Is IJSeam Then
        lLCOrSeamOrFrame = 1
    ElseIf TypeOf oElem Is IHFrame Then
        lLCOrSeamOrFrame = 1
    End If
    
    'if not LC or Frame or Seam
    If lLCOrSeamOrFrame = 0 Then
        Exit Function
    Else
    
        Dim oElements As JCmnShp_CollectionAlias
        Set oElements = vArg
        
        'if objects are not sent as an argument then return 0
        'The first object of oElements should be a  plate
        'The second object of oElements should be an IHFrame(first boundary)
        'The third object of oElements should be an IHFrame(second boundary)

        If oElements.Count < 3 Then
            Exit Function
        End If
        
        'check if intersection exists between the plate and oElem
        Dim oPlate As IJPlate
        If TypeOf oElements.Item(1) Is IJPlate Then
            Set oPlate = oElements.Item(1)
            'check if oPlate and oelem intersect
            Dim oObject As Object
            Set oObject = GetIntersection(oPlate, oElem)
            If Not oObject Is Nothing Then
                If TypeOf oObject Is IJWireBody Then
                    lIntersectWithPlate = 1
                End If
            End If
        End If

        'check if intersection exists between oElem and the boundaries
        Dim oFrame1 As IHFrame
        Dim oFrame2 As IHFrame
        If TypeOf oElements.Item(2) Is IHFrame And TypeOf oElements.Item(3) Is IHFrame Then
            Set oFrame1 = oElements.Item(2)
            Set oFrame2 = oElements.Item(3)
    
            'check if intersection exists between boundaries and oElem
            Dim oObject1 As Object
            Set oObject1 = GetIntersection(oFrame1, oElem)
    
            Dim oObject2 As Object
            Set oObject2 = GetIntersection(oFrame2, oElem)
            
            If Not oObject1 Is Nothing And Not oObject2 Is Nothing Then
                lIntersectWithFrames = 1
            End If
        End If
        
    End If
    
    'return 1 if intersection exists between plate, boundaries and oElem
    If lLCOrSeamOrFrame = 1 And lIntersectWithPlate = 1 And lIntersectWithFrames = 1 Then
        IsLCOrSeamOrFrame = 1
    End If
End Function

Private Function GetIntersection(pIntersectedObject As Object, pIntersectingObject As Object) As Object

    ' Find the intersection.
    Dim oGeometryIntersector    As IMSModelGeomOps.DGeomOpsIntersect
    Set oGeometryIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    Dim oIntersectionUnknown    As IUnknown        ' Resultant intersection.
    On Error Resume Next
    oGeometryIntersector.PlaceIntersectionObject Nothing, pIntersectedObject, pIntersectingObject, Nothing, oIntersectionUnknown
    Set GetIntersection = oIntersectionUnknown

    Set oGeometryIntersector = Nothing
    Set oIntersectionUnknown = Nothing

End Function

'********************************************************************
' Routine: IsConnectedToPinJigSuppPlates
'
' Abstract:
'       IsConnectedToPinJigSuppPlates = 0 if the Plate is not connected to Pin Jig supported plates
'       IsConnectedToPinJigSuppPlates = 1 if the Plate is connected to Pin Jig supported plates
'------------------------------------------------------------------------------'
Public Function IsConnectedToPinJigSuppPlates(oElem As Object, vArg As Variant) As Long
    
    IsConnectedToPinJigSuppPlates = 0
    
    Dim oElements As JCmnShp_CollectionAlias
    Set oElements = vArg
    
    If oElements Is Nothing Then
        Exit Function
    End If
        
    ' If the elements count is less than 3 or third input is not "User Selection" then exit
    If oElements.Count < 3 Then
        Exit Function
    ElseIf Not oElements.Item(3) = "User Selection" Then
        Exit Function
    End If
        
    ' Get the PinJig selected in the first step
    Dim oPinJig As IJPinJig
    Set oPinJig = oElements.Item(1)
    
    If oPinJig Is Nothing Then
        Exit Function
    End If
    
    ' Get the PinJig supported plates
    Dim oPinJigSuppPlates As IJElements
    Set oPinJigSuppPlates = oPinJig.SupportedPlates
    
    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oPartSupport As IJPartSupport
    Dim oConnPlatesColl As IJElements

    Set oPlatePartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oPartSupport = oPlatePartSupport
    
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Dim iCount As Long
    
    For iCount = 1 To oPinJigSuppPlates.Count
        Set oPartSupport.Part = oPinJigSuppPlates.Item(iCount)
    
        Set oConnPlatesColl = New JObjectCollection
        
        ' Get the PC based conncted objects for each PinJig supported plate
        Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
        oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol
        
        Dim jCount As Long
        For jCount = 1 To oConObjCol.Count
            If TypeOf oConObjCol.Item(jCount) Is IJPlatePart Then
                Set oSDPartSupport.Part = oConObjCol.Item(jCount)
                
                ' If the connection type is PARTSUPPORT_CONNTYPE_TEE and if the selected object is in ConnectedObjects then send true
                Dim bIsCrossOfTee As Boolean
                Dim eConnType As ContourConnectionType
                oSDPartSupport.GetConnectionTypeForContour oConnCol.Item(jCount), _
                                                            eConnType, _
                                                            bIsCrossOfTee
                                                   
                'If this is not TEE type conection, continue to the next item
                If eConnType = PARTSUPPORT_CONNTYPE_TEE Then
                    If oElem Is oConObjCol.Item(jCount) Then
                        IsConnectedToPinJigSuppPlates = 1
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
    
End Function
