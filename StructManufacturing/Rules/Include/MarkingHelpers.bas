Attribute VB_Name = "MarkingHelpers"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    MarkingHelpers.bas
'
'  History:
'       Suma Mallena         May 12. 2008   created
'******************************************************************

Option Explicit
Const MODULE = "MarkingHelpers.bas"

Public Function CheckValidFaceTypeForProfile(ByVal Part As Object, ByVal Upside As Long) As Long 'JXSEC_CODE
Const sMETHOD = "CheckValidFaceTypeForProfile"

    On Error GoTo ErrorHandler
    
    CheckValidFaceTypeForProfile = JXSEC_UNKNOWN
    
    Dim oStructConnectable  As IJStructConnectable
    Dim oEnumPorts          As IJElements
    Dim j                   As Integer
    Dim oPort               As IJPort
    Dim oStructPort         As IJStructPort
    Dim oEntityHelper       As New MfgEntityHelper
    Dim oPortIDsColl        As JCmnShp_CollectionAlias
    Dim i                   As Integer
    
    Set oEnumPorts = New JObjectCollection
    
    Set oStructConnectable = Part
    oEnumPorts.AddElements oStructConnectable.enumAllPorts
    
    Set oPortIDsColl = oEntityHelper.GetSectionIDsOfStructProfilePart(Part)
    
    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
         
        If oStructPort.SectionID = Upside Then
            If oStructPort.SectionID = JXSEC_TOP Then
            
                For i = 1 To oPortIDsColl.Count
                    If oPortIDsColl.Item(i) = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or oPortIDsColl.Item(i) = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
                        CheckValidFaceTypeForProfile = Upside
                        Exit For
                    End If
                Next
                
            ElseIf oStructPort.SectionID = JXSEC_BOTTOM Then
            
                For i = 1 To oPortIDsColl.Count
                    If oPortIDsColl.Item(i) = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or oPortIDsColl.Item(i) = JXSEC_BOTTOM_FLANGE_LEFT_TOP Then
                        CheckValidFaceTypeForProfile = Upside
                        Exit For
                    End If
                Next
                
            Else
       
                CheckValidFaceTypeForProfile = Upside
                Exit For
                
            End If
          
        End If
        
    Next
    
    Set oStructConnectable = Nothing
    Set oEnumPorts = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    Set oPortIDsColl = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
   
End Function

Public Function CheckValidFaceTypeForMember(ByVal Part As Object, ByVal Upside As Long) As Long 'JXSEC_CODE
Const sMETHOD = "CheckValidFaceTypeForMember"
    
    On Error GoTo ErrorHandler

    CheckValidFaceTypeForMember = CLng(JXSEC_UNKNOWN)
    
    Dim oMemberPartSupport      As IJMemberPartSupport
    Dim oPartSupport            As IJPartSupport
    Dim oStructConnectable      As IJStructConnectable
    Dim oEnumPorts              As IJElements
    Dim j                       As Integer
    Dim oPort                   As IJPort
    Dim oStructPort             As IJStructPort
    Dim oEntityHelper           As New MfgEntityHelper
    Dim oPortIDsColl            As JCmnShp_CollectionAlias
    Dim i                       As Integer
    
    Set oMemberPartSupport = New MemberPartSupport
    Set oPartSupport = oMemberPartSupport

    Set oPartSupport.Part = Part
    Set oStructConnectable = oPartSupport
   
    Set oEnumPorts = New JObjectCollection
    oEnumPorts.AddElements oStructConnectable.enumAllPorts
    
    Set oPortIDsColl = oEntityHelper.GetSectionIDsOfStructProfilePart(Part)
    
    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
     
        If oStructPort.SectionID = Upside Then
            If oStructPort.SectionID = JXSEC_TOP Then
            
                For i = 1 To oPortIDsColl.Count
                    If oPortIDsColl.Item(i) = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or oPortIDsColl.Item(i) = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
                        CheckValidFaceTypeForMember = Upside
                        Exit For
                    End If
                Next
                
            ElseIf oStructPort.SectionID = JXSEC_BOTTOM Then
            
                For i = 1 To oPortIDsColl.Count
                    If oPortIDsColl.Item(i) = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or oPortIDsColl.Item(i) = JXSEC_BOTTOM_FLANGE_LEFT_TOP Then
                        CheckValidFaceTypeForMember = Upside
                        Exit For
                    End If
                Next
                
            Else
       
                CheckValidFaceTypeForMember = Upside
                Exit For
                
            End If
        End If
        
    Next
    
    Set oMemberPartSupport = Nothing
    Set oPartSupport = Nothing
    Set oStructConnectable = Nothing
    Set oEnumPorts = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    Set oPortIDsColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
   
End Function

Public Function GetPlateUpsideBasedOnDirection(ByVal Plate As Object, ByVal DirVector As IJDVector) As GSCADMfgRulesDefinitions.enumPlateSide
Const sMETHOD = "GetPlateUpsideBasedOnDirection"

    On Error GoTo ErrorHandler
  
    GetPlateUpsideBasedOnDirection = UnDefinedSide

    Dim oGeomUtils          As IJTopologyLocate
    Dim oStructConnectable  As IJStructConnectable
    Dim oEnumPorts          As IJElements
    Dim j                   As Integer
    Dim oPort               As IJPort
    Dim oStructPort         As IJStructPort
    Dim oCenter             As IJDPosition
    Dim oNormal             As IJDVector
    Dim dAngle              As Double
    
    Set oGeomUtils = New TopologyLocate
    Set oEnumPorts = New JObjectCollection
    
    Set oStructConnectable = Plate
    oEnumPorts.AddElements oStructConnectable.enumAllPorts
    
    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
        
        If oStructPort.ContextID = CTX_BASE Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
            dAngle = DirVector.Dot(oNormal)

            If (dAngle > 0.94) Then
           
                GetPlateUpsideBasedOnDirection = BaseSide
                Exit For
                
            End If
            
         ElseIf oStructPort.ContextID = CTX_OFFSET Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
            dAngle = DirVector.Dot(oNormal)

            If (dAngle > 0.94) Then
            
                GetPlateUpsideBasedOnDirection = OffsetSide
                Exit For
                
            End If
            
        End If
        
    Next j
    
    Set oGeomUtils = Nothing
    Set oEnumPorts = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetProfileUpsideBasedOnDirection(ByVal Profile As Object, ByVal DirVector As IJDVector) As Long
Const sMETHOD = "GetProfileUpsideBasedOnDirection"

    On Error GoTo ErrorHandler
  
    GetProfileUpsideBasedOnDirection = JXSEC_UNKNOWN
    
    Dim oGeomUtils          As IJTopologyLocate
    Dim oStructConnectable  As IJStructConnectable
    Dim oEnumPorts          As IJElements
    Dim j                   As Integer
    Dim oPort               As IJPort
    Dim oStructPort         As IJStructPort
    Dim oCenter             As IJDPosition
    Dim oNormal             As IJDVector
    Dim dAngle              As Double

    
    Set oGeomUtils = New TopologyLocate
    
    Set oStructConnectable = Profile
    
    Set oEnumPorts = New JObjectCollection
    oEnumPorts.AddElements oStructConnectable.enumAllPorts
  
    For j = 1 To oEnumPorts.Count
        Set oPort = oEnumPorts.Item(j)
        Set oStructPort = oPort
        If oStructPort.SectionID = CLng(JXSEC_WEB_LEFT) Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
            dAngle = DirVector.Dot(oNormal)

            If (dAngle > 0.94) Then
           
                GetProfileUpsideBasedOnDirection = JXSEC_WEB_LEFT
                Exit For
                
            End If
            
         ElseIf oStructPort.SectionID = CLng(JXSEC_WEB_RIGHT) Then

            oGeomUtils.FindApproxCenterAndNormal oPort.Geometry, oCenter, oNormal
            
            dAngle = DirVector.Dot(oNormal)

            If (dAngle > 0.94) Then
            
                GetProfileUpsideBasedOnDirection = JXSEC_WEB_RIGHT
                Exit For
                
            End If
        End If
        
    Next j
    
    Set oGeomUtils = Nothing
    Set oEnumPorts = Nothing
    Set oPort = Nothing
    Set oStructPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
 
End Function
