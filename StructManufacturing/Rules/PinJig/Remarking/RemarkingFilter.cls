VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemarkingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const MODULE As String = "MfgPinJigRule.RemarkingFilter::"

Public Function IsAnXPlane(elem As Object, vArg As Variant) As Integer
    
    IsAnXPlane = 0
    
    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    ' Frame lines => Plane's normal is along X direction
    oDirectionVec.Set 1, 0, 0
    
    Dim oElements As IJElements
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    If Not oPinJig Is Nothing Then
        If TypeOf elem Is ISPGGridPlane Then
            Set oElements = GetReferencePlanesFromAllCSinRange(oPinJig, oDirectionVec)
            Set oDirectionVec = Nothing
            
            If Not oElements Is Nothing Then
                If oElements.Contains(elem) Then
                    IsAnXPlane = 1
                End If
            End If
        End If
    End If

End Function

Public Function IsAnYPlane(elem As Object, vArg As Variant) As Integer
    
    IsAnYPlane = 0
    
    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    ' Frame lines => Plane's normal is along Y direction
    oDirectionVec.Set 0, 1, 0
    
    Dim oElements As IJElements
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    If Not oPinJig Is Nothing Then
        If TypeOf elem Is ISPGGridPlane Then
            Set oElements = GetReferencePlanesFromAllCSinRange(oPinJig, oDirectionVec)
            Set oDirectionVec = Nothing
            
            If Not oElements Is Nothing Then
                If oElements.Contains(elem) Then
                    IsAnYPlane = 1
                End If
            End If
        End If
    End If
    
End Function

Public Function IsAnZPlane(elem As Object, vArg As Variant) As Integer
    
    IsAnZPlane = 0

    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    ' Frame lines => Plane's normal is along Z direction
    oDirectionVec.Set 0, 0, 1
    
    Dim oElements As IJElements
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    If Not oPinJig Is Nothing Then
        If TypeOf elem Is ISPGGridPlane Then
            Set oElements = GetReferencePlanesFromAllCSinRange(oPinJig, oDirectionVec)
            Set oDirectionVec = Nothing
            
            If Not oElements Is Nothing Then
                If oElements.Contains(elem) Then
                    IsAnZPlane = 1
                End If
            End If
        End If
    End If
    
End Function

Public Function IsUserExtendMark(elem As Object, vArg As Variant) As Integer

    IsUserExtendMark = 0
    
    Const OnlyThisPinJig As Boolean = False
    Dim oExtendPinJigMark As IJElements
    
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    If Not oPinJig Is Nothing Then
        ' Return only user marks of type "STRMFG_EXTEND_PINJIG_INTERSECTION"
        Set oExtendPinJigMark = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_EXTEND_PINJIG_INTERSECTION)
        If oExtendPinJigMark.Contains(elem) Then
            IsUserExtendMark = 1
        End If
        
    End If
    Set oExtendPinJigMark = Nothing

End Function


Public Function IsUserRemark(elem As Object, vArg As Variant) As Integer

    IsUserRemark = 0

    Const OnlyThisPinJig As Boolean = False
    
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    If Not oPinJig Is Nothing Then
        If TypeOf elem Is IJMfgMarkingLines_AE Then
            Dim oUserMarkOnThisPinJig As IJElements
            ' Return only user marks of type "STRMFG_PINJIG_MARK_DIAGONAL"
            Set oUserMarkOnThisPinJig = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_PINJIG_DIAGONAL)
        
            Dim oColl2 As IJElements
            ' Return only user marks of type "STRMFG_PINJIG_MARK"
            Set oColl2 = oPinJig.MarkingLinesOnSupportedPlates(OnlyThisPinJig, STRMFG_PINJIG_MARK)
        
            ' Append collection with STRMFG_PINJIG_MARK marks to User mark collection with STRMFG_PINJIG_MARK_DIAGONAL marks
            oUserMarkOnThisPinJig.AddElements oColl2
            Set oColl2 = Nothing
            
            If oUserMarkOnThisPinJig.Contains(elem) Then
                IsUserRemark = 1
            End If
            Set oUserMarkOnThisPinJig = Nothing
        End If
    End If
    
End Function

Public Function IsNavalArchLine(elem As Object, vArg As Variant) As Integer
    
    IsNavalArchLine = 0

    Const OnlyThisPinJig As Boolean = False
    
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
    
    Dim oPinJigNavalArchRule As IJDMfgPinJigRulePerRemarkingType
    Set oPinJigNavalArchRule = New NavalArchLines

    If Not oPinJig Is Nothing Then
        Dim oNavalArchMarks As IJElements
        Set oNavalArchMarks = oPinJigNavalArchRule.GetEntitiesForRemarking(oPinJig)
    
        If Not oNavalArchMarks Is Nothing Then
            If oNavalArchMarks.Contains(elem) Then
                IsNavalArchLine = 1
            End If
        End If
    End If
    
    Set oNavalArchMarks = Nothing
    Set oPinJigNavalArchRule = Nothing
    Set oPinJig = Nothing
    
End Function

Public Function IsConnectedPlatePart(elem As Object, vArg As Variant) As Integer

    Dim oObject As Object
    IsConnectedPlatePart = 0

    Set oObject = IsConnectedObject(elem, vArg)
    If Not oObject Is Nothing Then
        If TypeOf oObject Is IJPlatePart Then 'Or TypeOf oObject Is IJPlate Then
            IsConnectedPlatePart = 1
        End If
    End If
      
End Function

Public Function IsConnectedProfilePart(elem As Object, vArg As Variant) As Integer
    Dim oObject As Object
    IsConnectedProfilePart = 0
    
    Set oObject = IsConnectedObject(elem, vArg)
    If Not oObject Is Nothing Then
        If TypeOf oObject Is IJProfilePart Then 'Or TypeOf oObject Is IJProfile Then
            IsConnectedProfilePart = 1
        End If
    End If
      
End Function

Private Function IsConnectedObject(elem As Object, vArg As Variant) As Object
    Dim oPinJig As IJPinJig
    Set oPinJig = vArg
      
    If Not oPinJig Is Nothing Then

        Dim oSupportedplates As IJElements
        Set oSupportedplates = oPinJig.SupportedPlates
        Dim i As Integer
        
        For i = 1 To oSupportedplates.Count
    
            Dim oPartSupport As IJPartSupport
            Set oPartSupport = New PartSupport
            Set oPartSupport.Part = oSupportedplates.Item(i)
            Dim ConnectedObjColl As Collection
    
            ' Input filters will be applied ConnectionsColl to fill oConnections
            Dim ConnectionsColl As Collection
    
            Dim ThisPortColl As Collection
            Dim OtherPortColl As Collection
    
            oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                             ConnectedObjColl, ConnectionsColl, _
                                             ThisPortColl, OtherPortColl
                                             
            ' Get Plate system for input plate part
            Dim oPlateSystem As IJSystem
            oPartSupport.IsSystemDerivedPart oPlateSystem
            
            Dim oStructConnectable As IJStructConnectable
            Set oStructConnectable = oPlateSystem
            
            ' Get all objects connected to plate with a logical connection
            Dim ConnectedObjIEUK As IEnumUnknown
            Dim ConnectionsIEUK As IEnumUnknown
            oStructConnectable.enumConnectedObjects2 ConnectionLogical, ConnectedObjIEUK, ConnectionsIEUK
            
            Dim oConvertUtils As New CONVERTUTILITIESLib.CCollectionConversions
            Dim LCConnectedObjColl As Collection
            oConvertUtils.CreateVBCollectionFromIEnumUnknown ConnectedObjIEUK, LCConnectedObjColl
            
            Set oConvertUtils = Nothing
            Set ConnectedObjIEUK = Nothing
            Set ConnectionsIEUK = Nothing
            
            Dim j As Integer
            For j = 1 To LCConnectedObjColl.Count
                ConnectedObjColl.Add LCConnectedObjColl.Item(j)
            Next j
                        
            'if the collection contains oObject, set the remarking object type
            For j = 1 To ConnectedObjColl.Count
                If elem Is ConnectedObjColl.Item(j) Then
                   Set IsConnectedObject = ConnectedObjColl.Item(j)
                   Exit For
                End If
            Next j
            
            Set oPartSupport = Nothing
            Set ConnectedObjColl = Nothing
            Set ThisPortColl = Nothing
            Set OtherPortColl = Nothing
            Set ConnectionsColl = Nothing
            
            If Not IsConnectedObject Is Nothing Then Exit For
            
        Next i
    End If
End Function

Public Function IsDesignSeam(elem As Object) As Integer

    Dim oSeamType As IJDSeamType
    
    If TypeOf elem Is IJDSeamType Then
        Set oSeamType = elem
 
        If oSeamType.SeamType = sstDesignSeam Then 'sstIntersectionSeam Then
            IsDesignSeam = 1
        Else
            IsDesignSeam = 0
        End If
    End If

End Function
