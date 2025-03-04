VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrACDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-16, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrACDefCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/Apr/2011 - Created
'    25/Aug/2011 - pnalugol
'           Added new methods to support Inset brace (B-52232)
'    15/Sep/2011 - pnalugol
'           Modified/Added methods to support Braces on Flange penetrated cases
'    16/Sep/2011 - pnalugol
'           Modified/Added methods to support Flush Braces
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    16/may/2012 - hgunturu TR-212435: Modified the methods CM_CreateGenericWebCut(),CM_ConstructTopFlangeCut()
'                                      CM_ConstructBottomFlangeCut() to select proper Bounding
'    27/Jul/2012 - svsmylav
'      CR-216426: 'CM_IsCornerNeededForShapeAtEdge' method calls
'      a common helper 'IsCornerFeatureNeededForShapeAtEdge (so local implementation is removed).
'    31/Oct/2012 - rmcmillan
'      CR-221523: Added GetGlobalPortAfterChamfer method. Modified CM_CreateGenericWebCut to call it.
'    14/06/2013 - dsmamidi
'       TR- 224879 : End cuts are not getting modified when there is modification in bounding objects.
'                   Modified CM_IsWebCutNeeded() to upddate the Existing Webcut.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'    1/Aug/2013 - svsmylav
'               TR-223552: 'GetSketchPlaneForMSSymbol' method is called with 'False' option for
'                           persistent plane.
'    21/02/2014 - TR-244709 Modified CMNeedToCompute method to update the object based on the modifications on the property page.
'   28/Feb/2013 - skcheeka
'               TR-248958 : Updated CM_IsWebCutNeeded.Validation checks were provided by checking the ports before accessing the object
'   03/Nov/15 -  pkakula: TR-CP-278336 Generic Assebly connection fails when bounding is split
'                         Added CMNeedtoComputeWebCut()
'            DI-282754: Modified CreateAxisWebCut and CreateAxisFlangeCut Methods
'                       to support Standard Generic ACs
'    7/Jan/2016    svsmylav   DI-273986: PC was missing for few on-member cases as bounding port was
'                             filled using 'ReverseMap'. Modified CMConstructSeatedPC' method:
'                             fix uses mapped bounding port id as key to get real port.
'    15/June/2016    knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'*********************************************************************************************
Const MODULE = m_sProjectPath + "MbrACDefCM.cls"

Public Enum eBoundingCondition
    Outside = 1
    Edge = 2
    FaceAndOutside = 3
    EdgeAndOutside = 4
    FaceAndEdge = 5
    OutsideAndOutside = 6
    EdgeAndEdge = 7
    Center = 8
End Enum

Public Enum eEndCutLocation
    EndCutLocationUnknown = 0
    WebEndCut = 1
    TopFlangeEndCut = 2
    BottomFlangeEndCut = 3
End Enum

Public Enum ePenetratedCutPosition
    UnknownPenetratedCut = 0
    CenterCut = 1
    TopCut = 2
    BottomCut = 3
End Enum

' *******************************************************************************************
' Method:
' CM_CreateGenericWebCut
'
' Description:
' Custom method for creating a web cut for a Generic Assembly connection
' *******************************************************************************************
Public Function CM_CreateGenericWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_CreateGenericWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' ------------------------------------------
    ' Get the assembly connection and references
    ' ------------------------------------------
    sMsg = "Creating Web Cut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBoundingObjectColl As IJElements
    Dim oBoundingPort As IJPort
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim i As Long
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oBoundingObj As Object
    Dim oPort As IJPort
    Dim lStatus As Long
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim LongBoxNeeded As Boolean
    
    Set oAppConnection = pMemberDescription.CAO
    
    Dim eAssemblyConnectionType As eACType
    eAssemblyConnectionType = GetMbrAssemblyConnectionType(oAppConnection)
    
    Dim lDispId As Long
    lDispId = pMemberDescription.dispid
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If eAssemblyConnectionType = ACType_Mbr_Generic Then
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        ' -------------------------------------------------------------
        ' Get all the bounding objects from the ports related to the AC
        ' -------------------------------------------------------------
        
        Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        ' ---------------------------------
        ' Get the Bounding  Port
        ' ---------------------------------
          
        Set oSmartOcc = pMemberDescription.CAO
        Set oSmartItem = oSmartOcc.SmartItemObject
        
        If oSmartItem.Name = "Generic_2B_01" Then
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
            Set oBoundingObj = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
            Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObj, oEditJDArgument).Item(1)
        ElseIf oSmartItem.Name = "Generic_LongBoxCuts" Then
            LongBoxNeeded = True
        Else
            Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
        End If
        
    Set oSmartItem = Nothing
    Set oSmartOcc = Nothing
   
    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
        
        Set oConnAttrbs = New SDConnectionUtils
        If LessThan(oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count, 1, 0) Then
         Exit Function
        End If
        
        Set oBoundingObjectColl = New JObjectCollection
        
        For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
            oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
        Next
        
        Set oAssyConn = New AssemblyConn
        Set oAssyConn.object = pMemberDescription.CAO
        
        'Adding the bounding port of the assembly connection to AuxiliaryPorts collection
        oBoundingObjectColl.Add oAssyConn.Port2
        Set oBoundingPort = oBoundingObjectColl.Item(lDispId)
    Else
        sMsg = "Not a vlaid type of assembly connection"
        GoTo ErrorHandler
    End If
    
    ' --------------------------------------------
    ' Throw an error if the DispID is out of range
    ' --------------------------------------------
        If LongBoxNeeded Then
            '
        ElseIf lDispId > oBoundingObjectColl.Count Then
            sMsg = "Creating Web Cut ...lDispId(" & Trim(Str(lDispId)) & ") > Object Count(" & Trim(Str(oBoundingObjectColl.Count)) & ")"
            GoTo ErrorHandler
        End If
        
    ' ---------------------------------
    ' Get the Bounded  Port
    ' ---------------------------------
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' ---------------------
    ' Get edge mapping data
    ' ---------------------
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Dim oSectionAlias As Long
    Dim bPenetratesWeb As Boolean

    Set oEdgeMapColl = New Collection
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    If LongBoxNeeded Then
            '
    Else
    bIsBoundedTube = IsTubularMember(oBoundedData.MemberPart)
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
        bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    End If
    
    Dim bIsTubeByTube As Boolean
    bIsTubeByTube = False
    
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
        If bIsBoundedTube And bIsBoundingTube Then
            bIsTubeByTube = True
        Else
            Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, oSectionAlias, bPenetratesWeb)
        End If
    ElseIf TypeOf oBoundingPort.Connectable Is IJProfile Then
        Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, oSectionAlias, bPenetratesWeb)
    End If
    
    Dim oBndingStiffObj As Object
    Dim bIsMbrBounidng As Boolean
    Dim bIsStiffBounding As Boolean
    
    If eAssemblyConnectionType = ACType_Stiff_Generic Then
      Set oBndingStiffObj = oBoundingObjectColl.Item(lDispId)
      If TypeOf oBndingStiffObj.Connectable Is IJProfile Then
        bIsStiffBounding = True
      ElseIf TypeOf oBndingStiffObj.Connectable Is ISPSMemberPartCommon Then
        bIsMbrBounidng = True
      End If
    End If
    End If
    ' ----------------------------
    ' If this is the first item...
    ' ----------------------------
    Dim oSDO_BoundingMember As New StructDetailObjects.MemberPart
    Dim oSDO_BoundingProfile As New StructDetailObjects.ProfilePart
    Dim eWL As ConnectedEdgeInfo
    Dim eWR As ConnectedEdgeInfo
    Dim eTFL As ConnectedEdgeInfo
    Dim eTFR As ConnectedEdgeInfo
    
    If lDispId = 1 Then
        ' ---------------------------------------------------------------------------
        ' Loop through the bounding objects and see if the web intersects any of them
        ' ---------------------------------------------------------------------------
        Dim oWebleftPort As IJPort
        Dim oWebRightPort As IJPort
        
        Dim bIsWebLeftIntersecting As Boolean
        Dim bIsWebRightIntersecting As Boolean
    
        Dim lIndex As Long
        Dim bWebIntersecting As Boolean
        bWebIntersecting = False
        
        For lIndex = 1 To oBoundingObjectColl.Count
            If bIsBoundedTube Then
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
            Else
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
            End If
            
            If Not oWebleftPort Is Nothing Then
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
            End If
            If Not oWebRightPort Is Nothing Then
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
            End If
    
            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                bWebIntersecting = True
                Exit For

            ' If the web didn't intersect any object, check if it does when extended
            Else
                If Not oWebleftPort Is Nothing Then
                    bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
                End If
                If Not oWebRightPort Is Nothing Then
                    bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
                End If
            
                If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                    bWebIntersecting = True
                    Exit For
                End If
            End If
        Next lIndex
        
        ' ---------------------------------------------------------------------------------------------------------
        ' Create a web cut to the back-side port of the first bounding object if the web did not intersect anything
        ' ---------------------------------------------------------------------------------------------------------
        If bWebIntersecting = False Then
            If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Then
                If bIsBoundingTube Then
                    Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                Else
                    Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                End If
            ElseIf eAssemblyConnectionType = ACType_Mbr_Generic Then
                Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
            End If
            
            If eAssemblyConnectionType = ACType_Stiff_Generic Then
                If bIsMbrBounidng Then
                        If bIsBoundingTube Then
                            Set oBoundingPort = oBndingStiffObj
                        Else
                            Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                        End If
                ElseIf bIsStiffBounding Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                Else
                        Set oBoundingPort = oBndingStiffObj
                End If
            End If
            
        ' ---------------------------------------------
        ' If the web did intersect a bounding object...
        ' ---------------------------------------------
        Else
            ' ----------------------------------
            ' If the bounding object is a member
            ' ----------------------------------
            If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Or _
                bIsMbrBounidng Then
                
                If bIsTubeByTube Or bIsBoundingTube Then
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oBoundingPort = oBndingStiffObj
                    End If
                Else
                    ' ---------------------------
                    ' If the flange is penetrated
                    ' ---------------------------
                    If Not bPenetratesWeb Then
                        ' -------------------------------------------------------------------
                        ' Get the port on the Bounding Member that the Bounded Web Intersects
                        ' -------------------------------------------------------------------
                        GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eTFL, eTFR
                                                                                    
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                        End If
                        
                        ' ---------------------------------------------------------------------------
                        ' If the left web face is outside, use the information for the right web face
                        ' ---------------------------------------------------------------------------
                        If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                        Else
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                        End If
                    ' ----------------------------------------------------------------------------
                    ' If the web is penetrated, use the first selected port on the bounding member
                    ' ----------------------------------------------------------------------------
                    Else
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set oBoundingPort = oBndingStiffObj.Connectable
                        End If
                    End If
                End If
            
            ElseIf TypeOf oBoundingObjectColl.Item(lDispId) Is IJProfile Or _
                    bIsStiffBounding Then
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Profile that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingProfile.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                    End If
                    
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                    ' ----------------------------------------------------------------------------
                    ' If the web is penetrated, use the first selected port on the bounding member
                    ' ----------------------------------------------------------------------------
                    Else
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set oBoundingPort = oBndingStiffObj
                        End If
                End If
            Else
                ' ---------------------------------------------------------------------------------------------
                ' If the bounding object is not a member, assume it is a plate, and use the first selected port
                ' ---------------------------------------------------------------------------------------------
                If eAssemblyConnectionType = ACType_Mbr_Generic Then
                    If LongBoxNeeded Then
                        'The Bounding Port is set at the bottom of this method.
                    Else
                        'Get the Global Port After the Chamfer if it Exists
                        Set oBoundingPort = GetGlobalPortAfterChamfer(oBoundingPort)
                    End If
                ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                    Set oBoundingPort = oBndingStiffObj
                    'Get the Global Port After the Chamfer if it Exists
                    Set oBoundingPort = GetGlobalPortAfterChamfer(oBoundingPort)
                End If
            End If
        End If
    ' ---------------------------
    ' If the second or later item
    ' ---------------------------
    ElseIf LongBoxNeeded Then
            '
    Else
        ' ----------------------------------
        ' If the bounding object is a member
        ' ----------------------------------
        If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Or _
            bIsMbrBounidng Then
            If bIsTubeByTube Or bIsBoundingTube Then
                If eAssemblyConnectionType = ACType_Mbr_Generic Then
                    Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                    Set oBoundingPort = oBndingStiffObj
                End If
            Else
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Member that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                    End If
                    
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                ' ----------------------------------------------------------------------------
                ' If the web is penetrated, use the first selected port on the bounding member
                ' ----------------------------------------------------------------------------
                Else
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oBoundingPort = oBndingStiffObj
                    End If
                End If
            End If
        ElseIf TypeOf oBoundingObjectColl.Item(lDispId) Is IJProfile Or _
               bIsStiffBounding Then
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Profile that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingProfile.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                    End If
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                ' ----------------------------------------------------------------------------
                ' If the web is penetrated, use the first selected port on the bounding member
                ' ----------------------------------------------------------------------------
                Else
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oBoundingPort = oBndingStiffObj
                    End If
                End If
        Else
            ' ---------------------------------------------------------------------------------------------
            ' If the bounding object is not a member, assume it is a plate, and use the first selected port
            ' ---------------------------------------------------------------------------------------------
            If eAssemblyConnectionType = ACType_Mbr_Generic Then
                Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                'Get the Global Port After the Chamfer if it Exists
                Set oBoundingPort = GetGlobalPortAfterChamfer(oBoundingPort)
            ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                Set oBoundingPort = oBndingStiffObj
                'Get the Global Port After the Chamfer if it Exists
                Set oBoundingPort = GetGlobalPortAfterChamfer(oBoundingPort)
            End If
        End If
    End If

    ' ---------------------
    ' Set the parent system
    ' ---------------------
    sMsg = "Retreiving Parent System for WebCut"
    
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' ------------------
    ' Create the Web Cut
    ' ------------------
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    
    Dim sEndCutSel As String
    Dim sACSelection As String
    
    Dim oACSelection As Object
    Set oACSelection = pMemberDescription.CAO
    
    If TypeOf oACSelection Is IJSmartOccurrence Then
        Set oSmartOcc = oACSelection
        Set oSmartItem = oSmartOcc.SmartItemObject
        sACSelection = oSmartItem.Name
    End If
    
    Dim oHelper As New StructDetailObjects.Helper

    If sACSelection = "Generic_LongBox" Then
        sEndCutSel = "MbrLongBoxEndSel"

        Set oBoundingObj = oBoundingObjectColl.Item(1)
        Set oBoundingPort = oBoundingObj.AxisPort(SPSMemberAxisAlong)
    ElseIf LongBoxNeeded Then
        Dim iIndex As Long
        'For member part
        If lDispId = 1 Then
          sEndCutSel = "MbrLongBoxEndSel" ' This gives Center web cut
            For iIndex = 1 To oBoundingObjectColl.Count
              If TypeOf oBoundingObjectColl.Item(iIndex) Is ISPSMemberPartPrismatic Then
                Set oBoundingObj = oBoundingObjectColl.Item(iIndex)
                Set oBoundingPort = oBoundingObj.AxisPort(SPSMemberAxisAlong)
              End If
            Next iIndex
         'For plate part
        Else
          sEndCutSel = "RootGenWebSel"
            If lDispId = 2 Then              'This gives top web cut
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
            ElseIf lDispId = 3 Then          'This gives Bottom web cut
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
            End If
              Set oBoundingObj = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
              Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObj, oEditJDArgument).Item(1)
       End If
    Else
        sEndCutSel = "RootGenWebSel"
    End If
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pObject = oSDO_WebCut.object
    
    ' ---------------------------------------------------------------------------------
    ' In order to place corner features the web cuts must already be attached to the AE
    ' ---------------------------------------------------------------------------------
    Dim oStructProfilePart As IJStructProfilePart
    If TypeOf oSDO_WebCut.Bounded Is IJStructProfilePart Then
        Set oStructProfilePart = oSDO_WebCut.Bounded
        oStructProfilePart.PlaceEndCutFeature oSDO_WebCut.object, oSDO_WebCut.BoundedPort, eEndCutTypes.WebCut
    End If
    

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

' *******************************************************************************************
' Method:
' CM_ConstructTopFlangeCut
'
' Description:
' Custom method for creating a Top Flange cut for a Generic Assembly connection
' *******************************************************************************************
Public Function CM_ConstructTopFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_ConstructTopFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    Set oAppConnection = pMemberDescription.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim oBoundingObjectColl As IJElements
    Dim oBoundingObject As IJDObject
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oPort As IJPort
    Dim iIndex As Long
    
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Set oSmartOcc = pMemberDescription.CAO
    Set oSmartItem = oSmartOcc.SmartItemObject
        
    If lDispId - 5 > oBoundingObjectColl.Count Then
        sMsg = "Creating Top Flange Cut ...lDispId(" & Trim(Str(lDispId - 5)) & ") > Object Count(" & Trim(Str(oBoundingObjectColl.Count)) & ")"
        GoTo ErrorHandler
    ElseIf lDispId = 2 And oBoundingObjectColl.Count = 1 Then
        Set oBoundingObject = oBoundingObjectColl.Item(1)
        ' Set the Bounding Port to the first selected port from the bounding object
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
    ElseIf oSmartItem.Name = "Generic_LongBoxCuts" Then
         '
    ElseIf oSmartItem.Name = "Generic_2B_01" Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_RIGHT)
        Set oBoundingObject = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
        
    Else
        Set oBoundingObject = oBoundingObjectColl.Item(lDispId - 5)
        ' Set the Bounding Port to the first selected port from the bounding object
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
    End If
    
    Set oSmartItem = Nothing
    Set oSmartOcc = Nothing
    
    ' Set the Bounded Port from the AppConnection ports
    Set oBoundedPort = oBoundedData.AxisPort
      
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If

    'Create the Flange Cut
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebCut As IJStructFeature

    'The Web Cut will be a sibling on the same AC
    'If no Web Cuts exist on the AC, then check the other ACs at the same end port of the bounded
    'member for a web cut.
    On Error Resume Next
    Dim bFoundWebCut As Boolean
    bFoundWebCut = False
    For iIndex = 1 To pMemberObjects.Count
        Set oWebCut = pMemberObjects.Item(iIndex)
        If (Not (oWebCut Is Nothing)) Then
            If (oWebCut.get_StructFeatureType = SF_WebCut) Then
                bFoundWebCut = True
                Exit For
            End If
        End If
    Next iIndex
    
    If bFoundWebCut = False Then
        'There was no existing Web cut on the AC, check the other ACs
        Dim oDesignChild As IJDesignChild
        Dim oAssyConn As Object
        Set oDesignChild = pMemberDescription.CAO
        Set oAssyConn = oDesignChild.GetParent
        Set oDesignChild = oAssyConn
        Set oDesignParent = oDesignChild.GetParent
        
        Dim oChildren As IJDObjectCollection
        oDesignParent.GetChildren oChildren
        Dim oAC As Object
        For Each oAC In oChildren
            If TypeOf oAC Is IJStructAssemblyConnection And Not oAC Is oAssyConn Then
                Dim oACBoundedData As MemberConnectionData
                Dim oACBoundingData As MemberConnectionData
                InitMemberConnectionData oAC, oACBoundedData, oACBoundingData, lStatus, sMsg
                If oACBoundedData.ePortId = oBoundedData.ePortId And oACBoundedData.MemberPart = oACBoundingData.MemberPart Then
                        'If oAC allready has a web cut then use it.
                        Dim oACChildren As IJDObjectCollection
                        If TypeOf oAC Is IJDesignParent Then
                            Set oDesignParent = oAC
                            oDesignParent.GetChildren oACChildren
                            Dim oEndCut As Object
                            For Each oEndCut In oACChildren
                                If TypeOf oEndCut Is IJStructFeature Then
                                    Set oWebCut = oEndCut
                                    If oWebCut.get_StructFeatureType = SF_WebCut Then
                                        bFoundWebCut = True
                                        Exit For
                                    End If
                                End If
                            Next oEndCut
                        End If
                        If bFoundWebCut Then Exit For
                End If
            End If
        Next oAC
    End If
    
    On Error GoTo ErrorHandler

    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    Dim sACSelection As String
    
    Dim oACSelection As Object
    Set oACSelection = pMemberDescription.CAO
    
    If TypeOf oACSelection Is IJSmartOccurrence Then
        Set oSmartOcc = oACSelection
        Set oSmartItem = oSmartOcc.SmartItemObject
        sACSelection = oSmartItem.Name
    End If
    
    Dim oHelper As New StructDetailObjects.Helper
    
    If sACSelection = "Generic_LongBox" Or sACSelection = "Generic_LongBoxCuts" Then
        sEndCutSel = "MbrEndFlangeSel"
        Dim oBoundingObj As ISPSMemberPartCommon
        Set oBoundingObj = oBoundingObjectColl.Item(oBoundingObjectColl.Count)   'Passing mmeber part as bounding object
        Set oBoundingPort = oBoundingObj.AxisPort(SPSMemberAxisAlong)
    Else
        sEndCutSel = "RootGenFlangeSel"
    End If
    
    If Not oHelper.GetEquivalentLastPort(oBoundingPort) Is Nothing Then
        Set oBoundingPort = oHelper.GetEquivalentLastPort(oBoundingPort)
    Else
        '
    End If
    
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oWebCut, _
                       sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pObject = oSDO_FlangeCut.object
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' Method:
' CM_ConstructBottomFlangeCut
'
' Description:
' Custom method for creating a Bottom Flange cut for a Generic Assembly connection
' *******************************************************************************************
Public Function CM_ConstructBottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM.CM_ConstructBottomFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim sEndCutSel As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    Set oAppConnection = pMemberDescription.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim oBoundingObjectColl As IJElements
    Dim oBoundingObject As IJDObject
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oPort As IJPort
    Dim iIndex As Long
            
    Set oSmartOcc = pMemberDescription.CAO
    Set oSmartItem = oSmartOcc.SmartItemObject
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If lDispId - 10 > oBoundingObjectColl.Count Then
        sMsg = "Creating Bottom Flange Cut ...lDispId(" & Trim(Str(lDispId - 10)) & ") > Object Count(" & Trim(Str(oBoundingObjectColl.Count)) & ")"
        GoTo ErrorHandler
    ElseIf lDispId = 3 And oBoundingObjectColl.Count = 1 Then
        Set oBoundingObject = oBoundingObjectColl.Item(1)
        ' Set the Bounding Port to the first selected port from the bounding object
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
    ElseIf oSmartItem.Name = "Generic_LongBoxCuts" Then
          '
    ElseIf oSmartItem.Name = "Generic_2B_01" Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_RIGHT)
        Set oBoundingObject = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
    Else
        Set oBoundingObject = oBoundingObjectColl.Item(lDispId - 10)
        Set oBoundingPort = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
    End If
    
    ' Set the Bounded Port from the AppConnection ports

    Set oBoundedPort = oBoundedData.AxisPort
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If

    'Create the Flange Cut
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebCut As IJStructFeature

    'The Web Cut will be a sibling on the same AC
    'If no Web Cuts exist on the AC, then check the other ACs at the same end port of the bounded
    'member for a web cut.
    On Error Resume Next
    Dim bFoundWebCut As Boolean
    bFoundWebCut = False
    For iIndex = 1 To pMemberObjects.Count
        Set oWebCut = pMemberObjects.Item(iIndex)
        If (Not (oWebCut Is Nothing)) Then
            If (oWebCut.get_StructFeatureType = SF_WebCut) Then
                bFoundWebCut = True
                Exit For
            End If
        End If
    Next iIndex
    
    If bFoundWebCut = False Then
        'There was no existing Web cut on the AC, check the other ACs
        Dim oDesignChild As IJDesignChild
        Dim oAssyConn As Object
        Set oDesignChild = pMemberDescription.CAO
        Set oAssyConn = oDesignChild.GetParent
        Set oDesignChild = oAssyConn
        Set oDesignParent = oDesignChild.GetParent
        
        Dim oChildren As IJDObjectCollection
        oDesignParent.GetChildren oChildren
        Dim oAC As Object
        For Each oAC In oChildren
            If TypeOf oAC Is IJStructAssemblyConnection And Not oAC Is oAssyConn Then
                Dim oACBoundedData As MemberConnectionData
                Dim oACBoundingData As MemberConnectionData
                InitMemberConnectionData oAC, oACBoundedData, oACBoundingData, lStatus, sMsg
                If oACBoundedData.ePortId = oBoundedData.ePortId And oACBoundedData.MemberPart = oACBoundingData.MemberPart Then
                    'If oAC allready has a web cut then use it.
                    Dim oACChildren As IJDObjectCollection
                    If TypeOf oAC Is IJDesignParent Then
                        Set oDesignParent = oAC
                        oDesignParent.GetChildren oACChildren
                        Dim oEndCut As Object
                        For Each oEndCut In oACChildren
                            If TypeOf oEndCut Is IJStructFeature Then
                                Set oWebCut = oEndCut
                                If oWebCut.get_StructFeatureType = SF_WebCut Then
                                    bFoundWebCut = True
                                    Exit For
                                End If
                            End If
                        Next oEndCut
                    End If
                    If bFoundWebCut Then Exit For
                End If
            End If
        Next oAC
    End If
    On Error GoTo ErrorHandler

    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    Dim sACSelection As String
    
    Dim oACSelection As Object
    Set oACSelection = pMemberDescription.CAO
    
    If TypeOf oACSelection Is IJSmartOccurrence Then
        Set oSmartOcc = oACSelection
        Set oSmartItem = oSmartOcc.SmartItemObject
        sACSelection = oSmartItem.Name
    End If
    
    Dim oHelper As New StructDetailObjects.Helper
    
    If sACSelection = "Generic_LongBox" Or sACSelection = "Generic_LongBoxCuts" Then
        sEndCutSel = "MbrEndFlangeSel"
        Dim oBoundingObj As ISPSMemberPartCommon
        Set oBoundingObj = oBoundingObjectColl.Item(oBoundingObjectColl.Count)
        Set oBoundingPort = oBoundingObj.AxisPort(SPSMemberAxisAlong)
    Else
        sEndCutSel = "RootGenFlangeSel"
    End If
    
    If Not oHelper.GetEquivalentLastPort(oBoundingPort) Is Nothing Then
        Set oBoundingPort = oHelper.GetEquivalentLastPort(oBoundingPort)
    Else
        '
    End If
    
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oWebCut, _
                       sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pObject = oSDO_FlangeCut.object
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' Method:
' CM_IsWebCutNeeded
'
' Description: Conditional method to determine if a Web cut is needed
' *******************************************************************************************
Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    
    Dim oWebleftPort As IJPort
    Dim oWebRightPort As IJPort
    
    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean
    
    Dim bIsBoundedTube As Boolean
    bIsBoundedTube = IsTubularMember(oBoundedData.MemberPart)
    
    'If the Web Doesn't intersect any of the Bounding Objects then we will create a web cut to the
    'back port of the first bounding object
    If lDispId = 1 Then
        Dim lIndex As Long
        Dim bWebIntersecting As Boolean
        bWebIntersecting = False
        'For each bounding object, check if the web intersects the bounding object. If it doesn't intersect
        'any then create the web cut on the first bounding object.
        For lIndex = 1 To oBoundingObjectColl.Count
            If bIsBoundedTube Then
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
            Else
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
            End If
            
            If Not oWebleftPort Is Nothing Then
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
            End If
            If Not oWebRightPort Is Nothing Then
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
            End If
    
            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                bWebIntersecting = True
                If lIndex = 1 Then bIsNeeded = True
                Exit For
            Else
                'Check if the extended ports intersect the bounding object
                If Not oWebleftPort Is Nothing Then
                    bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
                End If
                If Not oWebRightPort Is Nothing Then
                    bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
                End If
                If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                    bWebIntersecting = True
                    If lIndex = 1 Then bIsNeeded = True
                    Exit For
                End If
            End If
        Next lIndex
        
        If bWebIntersecting = False Then
            'Web doesn't intersect any objects. However, we need a web cut as input to the flange cuts.
            'Create the Web Cut If there isn't allready a web cut created by a sibling AC
            Dim oDesignParent As IJDesignParent
            Dim oDesignChild As IJDesignChild
            Dim oAC As Object
            Set oAC = pMemberDescription.CAO
            Dim bWebCutExists As Boolean
            bWebCutExists = False
            
            If TypeOf oAC Is IJDesignChild Then
                Set oDesignChild = oAC
                Set oDesignParent = oDesignChild.GetParent
                
                Dim oChildren As IJDObjectCollection
                oDesignParent.GetChildren oChildren

                For Each oAC In oChildren
                    'Check each sibling Assembly Connection on the Bounded Member to see if it is for the same
                    'end port on the bounded member
                    If TypeOf oAC Is IJStructAssemblyConnection And Not oAC Is pMemberDescription.CAO Then
                        Dim oACBoundedData As MemberConnectionData
                        Dim oACBoundingData As MemberConnectionData
                        InitMemberConnectionData oAC, oACBoundedData, oACBoundingData, lStatus, sMsg
                        If oACBoundedData.ePortId = oBoundedData.ePortId And oACBoundedData.MemberPart = oACBoundingData.MemberPart Then
                            'AC is for the same end port and it is a generic AC
                            'check to see if any Web Cuts exist on the AC
                            Dim oACChildren As IJDObjectCollection
                            If TypeOf oAC Is IJDesignParent Then
                                Set oDesignParent = oAC
                                oDesignParent.GetChildren oACChildren
                                Dim oEndCut As Object
                                For Each oEndCut In oACChildren
                                    If TypeOf oEndCut Is IJStructFeature Then
                                        Dim oStructFeature As IJStructFeature
                                        Set oStructFeature = oEndCut
                                        If oStructFeature.get_StructFeatureType = SF_WebCut Then
                                            'Struct Feature is a Web Cut.  Therefore, we don't need to
                                            'create another web cut.
                                            bWebCutExists = True
                                            Exit For
                                        End If
                                    End If
                                Next oEndCut
                            End If
                            If bWebCutExists Then Exit For
                        End If
                    End If
                Next oAC
                If bWebCutExists Then
                    bIsNeeded = False
                Else
                    bIsNeeded = True
                End If
            End If
        End If
    End If
    
    If lDispId > 1 And lDispId <= oBoundingObjectColl.Count Then
        'A web cut should be created if the web of the bounded object intersects
        'the bounding object.
        If bIsBoundedTube Then
            Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
        Else
            Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
            Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
        End If
        If Not oWebleftPort Is Nothing Then
            bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId))
        End If
        If Not oWebRightPort Is Nothing Then
            bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId))
        End If
        
        If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
            bIsNeeded = True
        Else
            'Check if the extended ports intersect the bounding object and create the cut if the
            'bounding object is a member.
            bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId))
            bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId))
        
            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then bIsNeeded = True
        End If
    End If
    
    ' When the bounding objects ports are modified, the correspondig webcut needs to be updated...
    If (bIsNeeded) And (Not oAppConnection Is Nothing) Then
        
        Dim oMemObj As IJDMemberObjects
        Set oMemObj = oAppConnection
        
        On Error Resume Next
        Dim oItemMember As Object
        Set oItemMember = oMemObj.ItemByDispid(lDispId)
        Set oMemObj = Nothing
        On Error GoTo ErrorHandler
        
        If Not oItemMember Is Nothing Then
            If TypeOf oItemMember Is IJStructFeature Then
                Dim oFeatureSmartOcc As IJSmartOccurrence
                Set oFeatureSmartOcc = oItemMember
                'Getting the item member of webcut
                Dim oACObj As Object
                Set oACObj = pMemberDescription.CAO
                Dim oBoundingObject As Object
                'MODIFY THE BOUNDING OBJECT
                Dim oSDO_WebCut As New StructDetailObjects.WebCut
                Set oSDO_WebCut.object = oFeatureSmartOcc
                Dim featureUtils As New GSCADSDCreateModifyUtilities.SDFeatureUtils
                Dim pResourceMgr As IUnknown
                Dim oBdgPortfromGWCInput As IJPort
                ' GenericWebCutInput method is the refactored code of CM_CreateGenericWebCut to get the update port
                Set oBdgPortfromGWCInput = GenericWebCutInput(pMemberDescription, pResourceMgr, oSDO_WebCut.object)
                'Replacing the bounding Object with new Port
                Set oBoundingObject = oBoundingObjectColl.Item(lDispId)
                Dim OldBdgPort As GSCADSDPartSupport.IJStructPort
                Dim NewBdgPort As GSCADSDPartSupport.IJStructPort
                Dim OldBdgPortID As Long
                Dim NewBdgPortID As Long
                Dim ParentSystem As IJSystem
                Dim OldPort As IJPort
                Dim NewPort As IJPort
                Set pResourceMgr = GetResourceManagerFromObject(oACObj)
                If TypeOf oBdgPortfromGWCInput Is IJStructPort And TypeOf oSDO_WebCut.BoundingPort Is IJStructPort Then
                    'Getting the Port IDs from Old and new Ports
                    Set OldBdgPort = oSDO_WebCut.BoundingPort
                    Set NewBdgPort = oBdgPortfromGWCInput
                    Set NewPort = oBdgPortfromGWCInput
                    Set OldPort = oSDO_WebCut.BoundingPort
                    OldBdgPortID = OldBdgPort.OperatorID
                    NewBdgPortID = NewBdgPort.OperatorID
                
                    'Checking whether the port IDs are matching for both new and old bounding ports
                    If Not OldPort Is Nothing And Not NewPort Is Nothing Then
                        If OldBdgPortID = NewBdgPortID Then
                            If NewPort.Connectable Is OldPort.Connectable Then
                                oFeatureSmartOcc.Update
                            Else
                                featureUtils.CreateWebCut pResourceMgr, oBdgPortfromGWCInput, oSDO_WebCut.BoundedPort, "", ParentSystem, oSDO_WebCut.object
                            End If
                        Else
                            featureUtils.CreateWebCut pResourceMgr, oBdgPortfromGWCInput, oSDO_WebCut.BoundedPort, "", ParentSystem, oSDO_WebCut.object
                        End If
                    End If
                End If
                'This will create/Update the EndCuts accordingly
                Set oFeatureSmartOcc = Nothing
            End If
            Set oItemMember = Nothing
        End If
    End If
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' *******************************************************************************************
' Method:
' CM_IsTopFlangeCutNeeded
'
' Description: Conditional method to determine if a TopFlange cut is needed
' *******************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    
    Dim oWebleftPort As IJPort
    Dim oWebRightPort As IJPort
    
    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean
    
    Dim oTopFlangeRightTop As IJPort
    Dim oTopFlangeLeftTop As IJPort
    Dim oTopFlangeRightBottom As IJPort
    Dim oTopFlangeLeftBottom As IJPort

    'Booleans to specify if the ports are intersecting the object
    Dim bTopFlangeRightTop As Boolean
    Dim bTopFlangeLeftTop As Boolean
    Dim bTopFlangeRightBottom As Boolean
    Dim bTopFlangeLeftBottom As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight

    'A flange cut should be created if the Top Flange is intersecting the bounding object and the web is not intersecting
    'the bounding object. If the web is intersecting the bounding object, it will be responsible for creating the flange cut.
    If bTopFlangeLeft Or bTopFlangeRight Then
        If (lDispId - 5) <= oBoundingObjectColl.Count Then
            If bTopFlangeRight Then
                Set oTopFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_RIGHT_TOP)
                Set oTopFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
                If oTopFlangeRightTop Is Nothing Then Set oTopFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
                
                bTopFlangeRightTop = IsPortIntersectingObj(oTopFlangeRightTop, oBoundingObjectColl.Item(lDispId - 5))
                bTopFlangeRightBottom = IsPortIntersectingObj(oTopFlangeRightBottom, oBoundingObjectColl.Item(lDispId - 5))
            Else
                bTopFlangeRightTop = False
                bTopFlangeRightBottom = False
            End If
            
            If bTopFlangeLeft Then
                Set oTopFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_LEFT_TOP)
                Set oTopFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
                If oTopFlangeLeftTop Is Nothing Then Set oTopFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
            
                bTopFlangeLeftTop = IsPortIntersectingObj(oTopFlangeLeftTop, oBoundingObjectColl.Item(lDispId - 5))
                bTopFlangeLeftBottom = IsPortIntersectingObj(oTopFlangeLeftBottom, oBoundingObjectColl.Item(lDispId - 5))
            Else
                bTopFlangeLeftTop = False
                bTopFlangeLeftBottom = False
            End If
            
            'Top Flange is intersecting if any of the top flange ports are intersecting
            If bTopFlangeRightTop Or bTopFlangeLeftTop Or bTopFlangeRightBottom Or bTopFlangeLeftBottom Then
                'Create the flange cut if the web is not intersecting
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
                
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId - 5))
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 5))
                
                If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then
                    If TypeOf oBoundingObjectColl.Item(lDispId - 5) Is ISPSMemberPartPrismatic Then
                        bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId - 5))
                        bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 5))
                    
                        If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then bIsNeeded = True
                    Else
                        bIsNeeded = True
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' *******************************************************************************************
' Method:
' CM_IsBottomFlangeCutNeeded
'
' Description: Conditional method to determine if a Bottom Flange cut is needed
' *******************************************************************************************
Public Sub CM_IsBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsBottomFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long
    Dim lStatus As Long
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    
    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean
    
    Dim oWebleftPort As IJPort
    Dim oWebRightPort As IJPort
    
    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean
    
    Dim oBottomFlangeRightBottom As IJPort
    Dim oBottomFlangeLeftBottom As IJPort
    Dim oBottomFlangeRightTop As IJPort
    Dim oBottomFlangeLeftTop As IJPort

    'Booleans to specify if the ports are intersecting the object
    Dim bBottomFlangeRightBottom As Boolean
    Dim bBottomFlangeLeftBottom As Boolean
    Dim bBottomFlangeRightTop As Boolean
    Dim bBottomFlangeLeftTop As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    'A flange cut should be created if the Bottom Flange is intersecting the bounding object and the web is not intersecting
    'the bounding object. If the web is intersecting the bounding object, it will be responsible for creating the flange cut.
    If bBottomFlangeLeft Or bBottomFlangeRight Then
        If (lDispId - 10) <= oBoundingObjectColl.Count Then
            If bBottomFlangeRight Then
                Set oBottomFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM)
                Set oBottomFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
                If oBottomFlangeRightBottom Is Nothing Then Set oBottomFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
                
                bBottomFlangeRightBottom = IsPortIntersectingObj(oBottomFlangeRightBottom, oBoundingObjectColl.Item(lDispId - 10))
                bBottomFlangeRightTop = IsPortIntersectingObj(oBottomFlangeRightTop, oBoundingObjectColl.Item(lDispId - 10))
            Else
                bBottomFlangeRightBottom = False
                bBottomFlangeRightTop = False
            End If
            
            If bBottomFlangeLeft Then
                Set oBottomFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM)
                Set oBottomFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_LEFT_TOP)
                If oBottomFlangeLeftBottom Is Nothing Then Set oBottomFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
            
                bBottomFlangeLeftBottom = IsPortIntersectingObj(oBottomFlangeLeftBottom, oBoundingObjectColl.Item(lDispId - 10))
                bBottomFlangeLeftTop = IsPortIntersectingObj(oBottomFlangeLeftTop, oBoundingObjectColl.Item(lDispId - 10))
            Else
                bBottomFlangeLeftBottom = False
                bBottomFlangeLeftTop = False
            End If
            
            'Bottom Flange is intersecting if any of the bottom flange ports are intersecting
            If bBottomFlangeRightBottom Or bBottomFlangeLeftBottom Or bBottomFlangeRightTop Or bBottomFlangeLeftTop Then
                'Create the flange cut if the web is not intersecting
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
                
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId - 10))
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 10))
                
                If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then
                    If TypeOf oBoundingObjectColl.Item(lDispId - 10) Is ISPSMemberPartPrismatic Then
                        bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lDispId - 10))
                        bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 10))
                    
                        If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then bIsNeeded = True
                    Else
                        bIsNeeded = True
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' *******************************************************************************************
' Method:
' CMConstructWebCut
'
' Description:
' Custom method for creating a web cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM::CMConstructWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDispId As Long
    
    lDispId = pMemberDescription.dispid
    
    Dim sEndCutSel As String
      
    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
   
    ' Create Web Cut based on the Root Selection Rule
    sMsg = "Create WebCut based on Root Selection Rule"
    If lDispId = 2 Then
       'if Dispid = 2 should create Top Web Cut
       sEndCutSel = "TopWebSel"
       CreateAxisWebCut pMemberDescription, pResourceManager, sEndCutSel, pObject
    ElseIf lDispId = 3 Then
      'if Dispid = 3 should create Btm Web Cut
      sEndCutSel = "BottomWebSel"
      CreateAxisWebCut pMemberDescription, pResourceManager, sEndCutSel, pObject
    Else
      'if the code is here we expect(not necessarily) the dispid should be = 1
      'dispid = 1(Center Cut or Normal Cut)
      sEndCutSel = "MbrAxisWebSel"
      CreateAxisWebCut pMemberDescription, pResourceManager, sEndCutSel, pObject
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' Method:
' CMConstructFlangeCut
'
' Description:
' Custom method for creating a Flange cut for a Axis Assembly connection
' *******************************************************************************************
Public Function CMConstructFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                         ByVal pResourceManager As IUnknown, _
                                         ByRef pObject As Object)
    Const METHOD = "MbrACDefCM::CMConstructFlangeCut"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    Dim lDispId As Long
    
    lDispId = pMemberDescription.dispid
    
    Dim sEndCutSel As String
    
    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)

    sEndCutSel = "MbrAxisFlangeSel"
   
    ' Create Web Cut based on the Root Selection Rule
    sMsg = "Create WebCut based on Root Selection Rule"
    If lDispId = 4 Or _
       lDispId = 5 Or _
       lDispId = 18 Then
       'if Dispid = 4, 5 or 18 should create Top Flange Cut(Center Cut or Normal Cut)
       sEndCutSel = "MbrAxisFlangeSel"
       CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, False
    ElseIf lDispId = 6 Then
      'if Dispid = 6 should create Top Flange Cut for Top Flange
      sEndCutSel = "TopFlangeSel"
      CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, False
    ElseIf lDispId = 7 Then
      'if Dispid = 7 should create Btm Flange Cut for Top Flange
      sEndCutSel = "BottomFlangeSel"
      CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, False
    ElseIf lDispId = 8 Or _
           lDispId = 9 Or _
           lDispId = 19 Then
      'if Dispid = 8, 9, or 19 should create Btm Flange Cut(Center Cut or Normal Cut)
      sEndCutSel = "MbrAxisFlangeSel"
      CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, True
    ElseIf lDispId = 10 Then
      'if Dispid = 10 should create Top Flange Cut for btm Flange
      sEndCutSel = "TopFlangeSel"
      CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, True
    ElseIf lDispId = 11 Then
      'if Dispid = 11 should create Btm Flange Cut for btm Flange
      sEndCutSel = "BottomFlangeSel"
      CreateAxisFlangeCut pMemberDescription, pResourceManager, sEndCutSel, pObject, True
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function
' *******************************************************************************************
' Method:
' CreateAxisWebCut
'
' Description:
' Helper Method for creating a Web cut for a Axis Assembly connection
' *******************************************************************************************
Public Sub CreateAxisWebCut(pMemberDescription As IJDMemberDescription, _
                         pResourceManager As IUnknown, _
                         sEndCutSelRule As String, _
                         pEndCutObject As Object, _
                         Optional pBoundingPort As IJPort = Nothing, _
                         Optional pBoundedPort As IJPort = Nothing)
    Const METHOD = "MbrACDefCM::CreateAxisWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oSDO_BoundingMember As New StructDetailObjects.MemberPart
    Dim oSDO_WebCut As StructDetailObjects.WebCut
        
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
    
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    
    Set oBoundedPart = oBoundedPort.Connectable
    Set oBoundingPart = oBoundingPort.Connectable
                                                                                          
    If Not pBoundedPort Is Nothing Then
      Set oBoundedPort = pBoundedPort
    End If
    
    Dim bIsBoundingTube As Boolean
    bIsBoundingTube = False
    
    bIsBoundingTube = IsTubularMember(oBoundingPart)
    
    If Not bIsBoundingTube Then
        If pBoundingPort Is Nothing Then
            'Edge Mapping
            Dim oSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
              
            Set oSDO_BoundingMember.object = oBoundingPart
            
            Dim eWL As ConnectedEdgeInfo
            Dim eWR As ConnectedEdgeInfo
            Dim eTFL As ConnectedEdgeInfo
            Dim eTFR As ConnectedEdgeInfo
            
            Set oEdgeMapColl = New Collection
            bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
                        
            If Not bPenetratesWeb Then
              GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eTFL, eTFR
              Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, , bPenetratesWeb)
              If eWL.IntersectingEdge <> eWR.IntersectingEdge Then
                If eWL.IntersectingEdge = eBounding_Edge.Above Or _
                   eWL.IntersectingEdge = eBounding_Edge.Below Then
                   If eWR.IntersectingEdge = eBounding_Edge.Web_Right Or eWR.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Top_Flange_Right)))
                    ElseIf eWR.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Then
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                        'Set what Bounded WR intersects to
                    Else
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    End If
                ElseIf eWL.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or _
                       eWR.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                   'Set Bounding TFR mapped port as Bounding Port
                   Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Top_Flange_Right)))
                ElseIf eWL.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                       eWR.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Then
                   'Set Bounding BFR mapped Port As Bounding Port
                   Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                ElseIf eWL.IntersectingEdge = eBounding_Edge.Web_Right Then
                    If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl) Then
                        If eWR.IntersectingEdge = eBounding_Edge.Web_Right Then
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                        Else
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                        End If
                    Else
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                    End If
                Else
                   'Set Bounding WR mapped Port As Bounding Port
                   Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                End If
              Else
                If eWL.IntersectingEdge = eBounding_Edge.Above Or _
                   eWL.IntersectingEdge = eBounding_Edge.Below Then
                   'Set what Bounding WL maps to
                   Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Left)))
                Else
                   'Set what Bounded WL intersects to
                   Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                End If
              End If
            End If
        Else
            Set oBoundingPort = pBoundingPort
        End If
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Web Cut
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                       sEndCutSelRule, oSystemParent

    sMsg = "Return the created Web Cut"
    Set pEndCutObject = oSDO_WebCut.object
    
    Set oAppConnection = Nothing
    Set oSDO_BoundingMember = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub


' *******************************************************************************************
' Method:
' CreateAxisFlangeCut
'
' Description:
' Helper Method for creating a Flange cut for a Axis Assembly connection
' *******************************************************************************************
Public Sub CreateAxisFlangeCut(pMemberDescription As IJDMemberDescription, _
                           pResourceManager As IUnknown, _
                           sEndCutSelRule As String, _
                           pEndCutObject As Object, _
                           bIsBottomFlange As Boolean, _
                           Optional pBoundingPort As IJPort = Nothing, _
                           Optional pBoundedPort As IJPort = Nothing)
    Const METHOD = "MbrACDefCM::CreateAxisFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    Dim iIndex As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oSDO_BoundingMember As New StructDetailObjects.MemberPart
    
    Dim oMemberObjects As IJDMemberObjects
    Dim oWebCut As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
        
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
                         
    If Not pBoundedPort Is Nothing Then
      Set oBoundedPort = pBoundedPort
    End If

    Dim bIsBoundingTube As Boolean
    Dim bIsBoundedTube As Boolean
    bIsBoundingTube = False
    bIsBoundedTube = False
    
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    
    Set oBoundedPart = oBoundedPort.Connectable
    Set oBoundingPart = oBoundingPort.Connectable
    
    bIsBoundingTube = IsTubularMember(oBoundingPart)
    bIsBoundedTube = IsTubularMember(oBoundedPart)
    
    If Not bIsBoundingTube Then
        If pBoundingPort Is Nothing Then
          'Edge Mapping
          Dim oEndCutMappingRule As IJEndCutEdgeMappingRule
          Dim oSectionAlias As Long
          Dim bPenetratesWeb As Boolean
          Dim oEdgeMapColl As JCmnShp_CollectionAlias
              
          Set oSDO_BoundingMember.object = oBoundingPart
            
          Dim eTop As ConnectedEdgeInfo
          Dim eBtm As ConnectedEdgeInfo
          Dim eTFlangeInside As ConnectedEdgeInfo
          Dim eBFlangeInside As ConnectedEdgeInfo
                  
          Set oEndCutMappingRule = CreateEdgeMappingRuleSymbolInstance
          Set oEdgeMapColl = New Collection
          bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
          
          If bPenetratesWeb Then
             GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eTop, eBtm, eTFlangeInside, eBFlangeInside
             Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, , bPenetratesWeb)
             If bIsBottomFlange Then
                If eBtm.IntersectingEdge <> eBFlangeInside.IntersectingEdge Then
                  If eBtm.IntersectingEdge = eBounding_Edge.Above Or _
                     eBtm.IntersectingEdge = eBounding_Edge.Below Then
                    If eBFlangeInside.IntersectingEdge = eBounding_Edge.Web_Right And _
                        ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl) Then
                        'Bottom flange right exists, so set its right edge as port
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                    Else
                        'set BoundingPort to what Inside Btm Flange intersects
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBFlangeInside.IntersectingEdge)))
                    End If
                    
                  ElseIf eBtm.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                         eBtm.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                     'set BoundingPort to Btm Flange Intersecting Edge
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBtm.IntersectingEdge)))
                  ElseIf eBFlangeInside.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                         eBFlangeInside.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                     'set BoundingPort to Btm Flange Inside Intersecting Edge
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBFlangeInside.IntersectingEdge)))
                  ElseIf eBtm.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom And (eBFlangeInside.IntersectingEdge = Above Or Below) Then
                  
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Top_Flange_Right)))
                  ElseIf eBtm.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top And (eBFlangeInside.IntersectingEdge = Below Or Above) Then
                     
                      Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                     
                  Else
                     'set BoundingPort to mapped(BoundingObject's mapped) WR port
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                  End If
                Else
                  If eBtm.IntersectingEdge = eBounding_Edge.Above Or _
                     eBtm.IntersectingEdge = eBounding_Edge.Below Then
                      'set BoundingPort to mapped(BoundingObject's mapped) WL port
                      Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Left)))
                  Else
                      'Set Bounding Port to what Bounded Inside Btm Flange intersects
                      Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBFlangeInside.IntersectingEdge)))
                  End If
                End If
             Else
                If eTop.IntersectingEdge <> eTFlangeInside.IntersectingEdge Then
                  If eTop.IntersectingEdge = eBounding_Edge.Above Or _
                     eTop.IntersectingEdge = eBounding_Edge.Below Then
                    If eTFlangeInside.IntersectingEdge = eBounding_Edge.Web_Right And _
                        ItemExists(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl) Then
                        'Top flange right exists, so set its right edge as port
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Top_Flange_Right)))
                    Else
                        'set BoundingPort to what Inside Top Flange intersects
                        Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eTFlangeInside.IntersectingEdge)))
                    End If
                  ElseIf eTop.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                         eTop.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                     'set BoundingPort to Top Flange Intersecting Edge
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eTop.IntersectingEdge)))
                  ElseIf eTFlangeInside.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                         eTFlangeInside.IntersectingEdge = eBounding_Edge.Top_Flange_Right Then
                     'set BoundingPort to Top Flange Inside Intersecting Edge
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eTFlangeInside.IntersectingEdge)))
                     ElseIf eTop.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Bottom And (eTFlangeInside.IntersectingEdge = Below Or Above) Then
                     
                       Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Bottom_Flange_Right)))
                     ElseIf eTop.IntersectingEdge = eBounding_Edge.Top_Flange_Right_Bottom And (eTFlangeInside.IntersectingEdge = Above Or Below) Then
                     
                       Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Top_Flange_Right)))
                   
                     
                  Else
                     'set BoundingPort to mapped(BoundingObject's mapped) WR port
                     Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                  End If
                Else
                  If eTop.IntersectingEdge = eBounding_Edge.Above Or _
                     eTop.IntersectingEdge = eBounding_Edge.Below Then
                      'set BoundingPort to mapped(BoundingObject's mapped) WL port
                      Set oBoundingPort = oSDO_BoundingMember.SubPort(oEdgeMapColl.Item(CStr(eBounding_Edge.Web_Right)))
                  Else
                      'Set Bounding Port to what Bounded Inside Btm Flange intersects
                      Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eTFlangeInside.IntersectingEdge)))
                  End If
                End If
             End If
          End If
        Else
          Set oBoundingPort = pBoundingPort
        End If
    End If
    
    ' Need to get the IJSystem Interface from this CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
     
    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    Set oMemberObjects = oAppConnection
    
    ' -------------------------------------
    'Get all the web cuts in the connection
    ' -------------------------------------
    Dim oWebCutCollection As New Collection
    Dim oCtrWebCut As Object
    Dim oTopWebCut As Object
    Dim oBtmWebCut As Object
    
    Set oCtrWebCut = oMemberObjects.ItemByDispid(1)
    Set oTopWebCut = oMemberObjects.ItemByDispid(2)
    Set oBtmWebCut = oMemberObjects.ItemByDispid(3)
    
    If Not oCtrWebCut Is Nothing Then
        oWebCutCollection.Add oCtrWebCut
    End If
    
    If Not oTopWebCut Is Nothing Then
        Dim oItem2 As IJStructFeature
        Set oItem2 = oTopWebCut
        If oItem2.get_StructFeatureType = SF_WebCut Then
            oWebCutCollection.Add oTopWebCut
        End If
    End If
    
    If Not oBtmWebCut Is Nothing Then
        Dim oItem3 As IJStructFeature
        Set oItem3 = oBtmWebCut
        If oItem3.get_StructFeatureType = SF_WebCut Then
            oWebCutCollection.Add oBtmWebCut
        End If
    End If
    
    ' Possible cases:
    ' 1) only a single web cut - The web cut is the input.  Could be center, top, or bottom.
    ' 2) center and bottom web cuts - Center web cut is input for top flange cut, bottom web cut is input for bottom flange cut
    ' 3) center and top cut - Center web cut is input for bottom flange cut, top web cut is input for top flange cut
    ' 4) three web cuts - top web cut is input for top flange cut, bottom web cut is input for bottom flange cut
    If oWebCutCollection.Count = 1 Then
        Set oWebCut = oWebCutCollection.Item(1)
    ElseIf oWebCutCollection.Count = 2 Then
        If Not oCtrWebCut Is Nothing And Not oBtmWebCut Is Nothing Then
           If bIsBottomFlange Then
                Set oWebCut = oBtmWebCut
            Else
                Set oWebCut = oCtrWebCut
                End If
        ElseIf Not oCtrWebCut Is Nothing And Not oTopWebCut Is Nothing Then
            If bIsBottomFlange Then
                Set oWebCut = oCtrWebCut
           Else
                Set oWebCut = oTopWebCut
                End If
           End If
    ElseIf oWebCutCollection.Count = 3 Then
        If bIsBottomFlange Then
            Set oWebCut = oBtmWebCut
        Else
            Set oWebCut = oTopWebCut
       End If
    Else
      'WebCut is needed to create the Flange Cut
      'Go to Error Handler
       sMsg = ": Could not set(find) WebCut as input for the creation of Flange Cut"
       GoTo ErrorHandler
    End If
            
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oWebCut, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pEndCutObject = oSDO_FlangeCut.object
    
    Set oAppConnection = Nothing
    Set oSDO_BoundingMember = Nothing
    Set oEndCutMappingRule = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

' *******************************************************************************************
' Method:
' CM_CopyQuestionsToTopFlangeCut
'
' Description: copy questions to end cut rules to set BottomFlange to
'   - Yes; if the cut is to be applied on bottom flange
'   - No; if the cut is to be applied on Top Flange
' *******************************************************************************************
Public Sub CM_CopyQuestionsToTopFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsToTopFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Set Top flange"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange", "No"
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

' *******************************************************************************************
' Method:
' CM_CopyQuestionsToBottomFlangeCut
'
' Description: copy questions to end cut rules to set BottomFlange to
'   - Yes; if the cut is to be applied on bottom flange
'   - No; if the cut is to be applied on Top Flange
' *******************************************************************************************
Public Sub CM_CopyQuestionsToBottomFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsToBottomFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange", "Yes"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' *******************************************************************************************
' For Axis member end cuts
' Method:
' CMCopyQuestionsToTopFlangeCut
'
' Description: copy questions to end cut rules to set BottomFlange to
'   - Yes; if the cut is to be applied on bottom flange
'   - No; if the cut is to be applied on Top Flange
' *******************************************************************************************
Public Sub CMCopyQuestionsToTopFlangeCut(pMemberDescription As IJDMemberDescription)
Const METHOD = "MbrACDefCM::CMCopyQuestionsToTopFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Set Top flange"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    If pMemberDescription.dispid = 4 Or _
       pMemberDescription.dispid = 5 Or _
       pMemberDescription.dispid = 18 Then
       pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrAxisFlangeSel", "BottomFlange", "No"
    ElseIf pMemberDescription.dispid = 6 Then
       pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.TopFlangeSel", "BottomFlange", "No"
    ElseIf pMemberDescription.dispid = 7 Then
       pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.BottomFlangeSel", "BottomFlange", "No"
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' *******************************************************************************************
' For Axis member end cuts
' Method:
' CMCopyQuestionsToBottomFlangeCut
'
' Description: copy questions to end cut rules to set BottomFlange to
'   - Yes; if the cut is to be applied on bottom flange
'   - No; if the cut is to be applied on Top Flange
' *******************************************************************************************
Public Sub CMCopyQuestionsToBottomFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CMCopyQuestionsToBottomFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    If pMemberDescription.dispid = 8 Or _
       pMemberDescription.dispid = 9 Or _
       pMemberDescription.dispid = 19 Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrAxisFlangeSel", "BottomFlange", "Yes"
    ElseIf pMemberDescription.dispid = 10 Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.TopFlangeSel", "BottomFlange", "Yes"
    ElseIf pMemberDescription.dispid = 11 Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.BottomFlangeSel", "BottomFlange", "Yes"
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the respective smart occurences
' *******************************************************************************************
Public Sub CMNeedToCompute(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "MbrACDefCM::CMNeedToCompute"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a smart occurence"
    
       If Not pObject Is Nothing Then
          If TypeOf pObject Is IJStructFeature Then
            Dim oSD_WC As New StructDetailObjects.WebCut
            Set oSD_WC.object = pObject
            Dim oSelectionRuleInfo As IJSelectionRuleInfo
            Set oSelectionRuleInfo = pPropertyDescriptions.CAO
            Dim eSelectionModificationType As SelectionRuleModificationType
            eSelectionModificationType = oSelectionRuleInfo.IsSelectionRuleModified
            If eSelectionModificationType = SelectionRule_CurrentRuleModified Then
                ' selection rule is modified from previous, hence update the object.
                oSD_WC.ForceUpdateForFlangeCuts
            End If
          End If
       End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_CreateEndToEndWebCut
' Description :
'       Create WebCut given :
'           the Given the MemberDescription and Root Selection Rule
'
'*********************************************************************************************
Public Sub CM_CreateEndToEndWebCut(pMemberDescription As IJDMemberDescription, _
                         pResourceManager As IUnknown, _
                         sEndCutSelRule As String, _
                         bUseBoundingEndPort As Boolean, _
                         bUseBoundedAlongPort As Boolean, _
                         lEndCutSet As Long, _
                         pEndCutObject As Object)
    Const METHOD = "MbrAssemblyUtilities::CM_CreateEndToEndWebCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oPort As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_WebCut As StructDetailObjects.WebCut

    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    If bUseBoundedAlongPort Then
        ' By default, the Bounded Port is the AxisStart or AxisEnd Port
        ' For ShortBox cases, need the Bounded AxisLong Port
        Set oBoundedPort = oBoundedData.MemberPart.AxisPort(SPSMemberAxisAlong)
    End If
    
    Set oBoundingPort = oBoundingData.AxisPort
    If bUseBoundingEndPort Then
        ' By default, the Bounding Port is the AlongAxis Port
        ' For End to End cases, need the Bounding End Port
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    
    If lEndCutSet = 2 Then
        ' For the Second set of End Cuts
        ' Switch the Bounding and Bounded Ports used to create the EndCut
        Set oPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oPort
        Set oPort = Nothing
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Web Cut
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pEndCutObject = oSDO_WebCut.object
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*************************************************************************
'Function
'CM_CreateEndToEndFlangeCut
'
'Abstract
'   Create FlangeCut given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'   lEndCutSet
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub CM_CreateEndToEndFlangeCut(pMemberDescription As IJDMemberDescription, _
                            pResourceManager As IUnknown, _
                            sEndCutSelRule As String, _
                            bUseBoundingEndPort As Boolean, _
                            bUseBoundedAlongPort As Boolean, _
                            lEndCutSet As Long, _
                            pEndCutObject As Object)
    Const METHOD = "MbrAssemblyUtilities::CM_CreateEndToEndFlangeCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    Dim nWebCuts As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oPort As Object
    Dim oWebCut As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    If bUseBoundedAlongPort Then
        ' By default, the Bounded Port is the AxisStart or AxisEnd Port
        ' For ShortBox cases, need the Bounded AxisLong Port
        Set oBoundedPort = oBoundedData.MemberPart.AxisPort(SPSMemberAxisAlong)
    End If
    
    Set oBoundingPort = oBoundingData.AxisPort
    If bUseBoundingEndPort Then
        ' By default, the Bounding Port is the AlongAxis Port
        ' For End to End cases, need the Bounding End Port
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    

    If lEndCutSet = 2 Then
        ' For the Second set of End Cuts
        ' Switch the Bounding and Bounded Ports used to create the EndCut
        Set oPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oPort
        Set oPort = Nothing
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    nWebCuts = 0
    Set oMemberObjects = oAppConnection
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(iIndex)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    nWebCuts = nWebCuts + 1
                    If nWebCuts = lEndCutSet Then
                        Set oWebCut = oStructFeature
                        Exit For
                    End If
                End If
            End If
        End If
    Next iIndex
        
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oWebCut, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pEndCutObject = oSDO_FlangeCut.object
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
' *******************************************************************************************
' Method:
' CMConstructBraceAC1
'
' Description: This method creates AC between the Child inset brace and Bounding of the Parent AC.
' AC is created between the Start port of the Brace and Axis-Along port of the Bounding(of the Parent AC).
' *******************************************************************************************
Public Function CMConstructBraceAC1(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
                                   
    Const METHOD = "MbrACDefCM::CMConstructBraceAC1"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDispId As Long
    Dim lBracedispID As Long
    
    lDispId = pMemberDescription.dispid
    
    'incase if the ACSelection = MbrAxisByCenter, two braces (Top/Bottom) are possible.
    'In such caseTop brace is created with dispid 12, Bottom brace is created with dispid 15.
    If lDispId = 13 Or lDispId = 14 Then
        lBracedispID = 12
    ElseIf lDispId = 16 Or lDispId = 17 Then
        lBracedispID = 15
    End If
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    Dim oBoundedData As MemberConnectionData, oBoundingData As MemberConnectionData
    
    Dim oChildBrace As Object
    Set oChildBrace = oMemObjs.ItemByDispid(lBracedispID)
    
    If TypeOf oChildBrace Is ISPSMemberPartPrismatic Then
        Dim oBraceMember As ISPSMemberPartPrismatic
        Set oBraceMember = oChildBrace
    End If
    
    Set oBoundedPort = oBraceMember.AxisPort(SPSMemberAxisStart)
    
    Dim lStatus As Long
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    
    Set oBoundingPort = oBoundingData.MemberPart.AxisPort(SPSMemberAxisAlong)
    
    Dim oStructConn As IJStructAssemblyConnection
    
    'create struct assembly connection
    CreateChildACForBrace pResourceManager, oBoundedPort, oBoundingPort, oStructConn

    'Set naming rule to name the nested AC
    SetAssemblyConnectionNameRule oStructConn
    
    'make the AC as a child of the child inset brace created
    Dim oDesignParent As IJDesignParent
    Set oDesignParent = oChildBrace
    
    oDesignParent.AddChild oStructConn

    Set pObject = oStructConn
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' Method:
' CreateMemberBrace
'
' Description: creates member brace given start and end positions, cross section,bIsTop as optional
' bIsTop controls the angle of rotation of brace
' sets all the attributes of the member part
' *******************************************************************************************
Public Sub CreateMemberBrace(pMemberDescription As IJDMemberDescription, _
                           pResourceManager As IUnknown, oStartPos As IJDPosition, oEndPos As IJDPosition, oCrossSection As IJCrossSection, _
                           ByRef pMemberBrace As Object, Optional bIsTop As Boolean, Optional bIsPenetratesWeb As Boolean)
                           
Const METHOD = "CreateMemberBrace"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String, lStatus As Long
    
    Dim oSPSMemberFactory As New SPSMemberFactory
    Dim oMemberPart As ISPSMemberPartPrismatic
    
    'create SPS member part
    Set oMemberPart = oSPSMemberFactory.CreateMemberPartPrismatic(3, pResourceManager)
    
    Dim oMaterial As Object
    Dim sSecName As String
    Dim pAttr As IJDAttributes
    
    Dim oStructQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Set oStructQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    
    If oCrossSection Is Nothing Then
        Set oCrossSection = oStructQuery.GetCrossSection("AISC-LRFD-3.1", "W14X22")
    Else
        Set pAttr = oCrossSection
        sSecName = pAttr.CollectionOfAttributes("IStructCrossSection").Item("SectionName").Value
        Set oCrossSection = oStructQuery.GetCrossSection("AISC-LRFD-3.1", sSecName)
    End If

    'set material, cross section, material defintion
    Set oMaterial = oStructQuery.GetMaterialByGrade("Air", "A")
    Set oStructQuery = Nothing
             
    
    Set oMemberPart.CrossSection.definition = oCrossSection
    Set oCrossSection = Nothing
    
    Set oMemberPart.MaterialDefinition = oMaterial
    Set oMaterial = Nothing
    
    'set cardinal point, member rotation angle, type category, type
    oMemberPart.CrossSection.CardinalPoint = 2
    
    'Get the rotation angle of bounded
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oBoundedData As MemberConnectionData, oBoundingData As MemberConnectionData
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If bIsTop Then
        oMemberPart.Rotation.BetaAngle = (oBoundedData.MemberPart.Rotation.BetaAngle) + GetPI
    Else
        oMemberPart.Rotation.BetaAngle = oBoundedData.MemberPart.Rotation.BetaAngle
    End If
   
    oMemberPart.MemberType.TypeCategory = 3
    oMemberPart.MemberType.Type = 3
    
    'set start and end points of the brace
    Dim oLine3dPart As Line3d
    Set oLine3dPart = oMemberPart.Axis
    
    oLine3dPart.SetStartPoint oStartPos.x, oStartPos.y, oStartPos.z
    oLine3dPart.SetEndPoint oEndPos.x, oEndPos.y, oEndPos.z
    
    'Set child inset brace as a child of the AC between the Parent Bounded and Parent Bounding.
    Dim oParent As IJDesignParent
    Set oParent = pMemberDescription.CAO
    oParent.AddChild oMemberPart
    
    'Set naming rule
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim vAttrib As String
    
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    oNameRuleHlpr.GetEntityNamingRulesGivenProgID "SPSMembers.SPSMemberPartPrismatic", NamingRules
    
    If NamingRules.Count > 0 Then
        Set oNameRuleHolder = NamingRules.Item(1)
        vAttrib = oNameRuleHolder.Name
        Set oNameRuleHolder = Nothing
    End If
    NamingRules.Clear
    Set oNameRuleHlpr = Nothing
    
    NameMemberObject oMemberPart, vAttrib
    
    Set pMemberBrace = oMemberPart
    
    Set oMemberPart = Nothing
    Set oSPSMemberFactory = Nothing
    Set oParent = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
' *******************************************************************************************
' Method:
' CMConstructBraceAC2
'
' Description: This method creates AC between the Child inset brace and Bounded of the Parent AC.
' AC is created between the End port of the Brace and Axis-Along port of the Bounded(of the Parent AC).
' *******************************************************************************************
Public Function CMConstructBraceAC2(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "MbrACDefCM::CMConstructBraceAC2"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    Dim lBracedispID As Long
    
    lDispId = pMemberDescription.dispid

    'incase if the ACSelection = MbrAxisByCenter, two braces (Top/Bottom) are possible.
    'In such caseTop brace is created with dispid 12, Bottom brace is created with dispid 15.
    If lDispId = 13 Or lDispId = 14 Then
        lBracedispID = 12
    ElseIf lDispId = 16 Or lDispId = 17 Then
        lBracedispID = 15
    End If
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    Dim oBoundedData As MemberConnectionData, oBoundingData As MemberConnectionData
    
    Dim oChildBrace As Object
    Set oChildBrace = oMemObjs.ItemByDispid(lBracedispID)
    
    If TypeOf oChildBrace Is ISPSMemberPartPrismatic Then
        Dim oBraceMember As ISPSMemberPartPrismatic
        Set oBraceMember = oChildBrace
    End If
    
    Set oBoundedPort = oBraceMember.AxisPort(SPSMemberAxisEnd)
    
    Dim lStatus As Long
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    
    Set oBoundingPort = oBoundedData.MemberPart.AxisPort(SPSMemberAxisAlong)
    
    Dim oStructConn As IJStructAssemblyConnection
    
    'create struct assembly connection
    CreateChildACForBrace pResourceManager, oBoundedPort, oBoundingPort, oStructConn
    
    'Set assembly connection name rule
    SetAssemblyConnectionNameRule oStructConn
    
    'set AC as child of the InsetBrace
    Dim oDesignParent As IJDesignParent
    Set oDesignParent = oChildBrace
    
    oDesignParent.AddChild oStructConn

    Set pObject = oStructConn
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*********************************************************************************************
' Method      : CreateChildACForBrace
' Description : Construct methods CMConstructBraceAC1 & CMConstructBraceAC2 call this method to create the StructAssembly connection
' Create a member Axis AC
'********************************************************************************************
Public Sub CreateChildACForBrace(ByVal pResourceManager As IUnknown, oBoundedPort As IJPort, oBoundingPort As IJPort, ByRef oStructConn As IJStructAssemblyConnection)
    Const METHOD = "MbrACDefCM::CreateChildACForBrace"
    On Error GoTo ErrorHandler
    
    'Create StructAssemblyConnection
    Dim oConnections As New CMNSTRStructConnections.StructConnectionsFactory
    Set oStructConn = oConnections.CreateStructAssemblyConnection(pResourceManager)
    
    Set oConnections = Nothing
     
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = oStructConn
    
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oRefColl As IJDReferencesCollection
    
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefColl = oSymFactory.CreateEntity(referencesCollection, pResourceManager)
    Set oSymFactory = Nothing
    
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollectionHelper As IMSRelation.DCollectionHelper
    
    Set oRelationHelper = oSmartOcc
    Set oCollectionHelper = oRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    Set oRelationHelper = Nothing
    
    ' Remove existing one
    If oCollectionHelper.Count = 1 Then
        Dim oOldRefColl As IJDReferencesCollection
    
        Set oOldRefColl = oCollectionHelper.Item("RC")
        oOldRefColl.IJDEditJDArgument.RemoveAll
        Set oOldRefColl = Nothing
        oCollectionHelper.Remove ("RC")
    End If
    
    ' Add new reference collection
    Dim oRelationshipHelper As DRelationshipHelper
    Dim oRevision As IJRevision
    
    oCollectionHelper.Add oRefColl, "RC", oRelationshipHelper
    Set oCollectionHelper = Nothing
    
    Set oRevision = New JRevision
    oRevision.AddRelationship oRelationshipHelper
    Set oRevision = Nothing
    
    'Add ports to the connection
    Dim mAppConnection As IJAppConnection
    Set mAppConnection = oStructConn
    
    mAppConnection.addPort oBoundedPort
    mAppConnection.addPort oBoundingPort
    
    oSmartOcc.RootSelection = "MbrAxisSel" 'MbrAxisSel is the smart class required to create Axis ACs
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
' Method      : CM_IsTopBraceNeeded
' Description : This is a conditional which determines if a top brace needs to be created. Takes the answer to the question on
' the property page and creates the brace appropriately.
'********************************************************************************************
Public Sub CM_IsTopBraceNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = "::CM_IsTopBraceNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Final Construct of Brace"
    
    Dim sInsetBraceAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "TopBraceType", sInsetBraceAnswer
    If sInsetBraceAnswer = "InsetMember" Or sInsetBraceAnswer = "FlushMember" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomBraceNeeded
' Description : This is a conditional which determines if a bottom brace needs to be created. Takes the answer to the question on
' the property page and creates the brace appropriately.
'********************************************************************************************
Public Sub CM_IsBottomBraceNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = "::CM_IsBottomBraceNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Final Construct of Brace"
    
    Dim sInsetBraceAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "BottomBraceType", sInsetBraceAnswer
    If sInsetBraceAnswer = "InsetMember" Or sInsetBraceAnswer = "FlushMember" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'*********************************************************************************************
' Method      : CM_IsTopBraceAC1Needed
' Description : Conditional method to determine if an AC between brace and Bounded(of the Parent AC) is needed
' Creates AC when the brace exists
'********************************************************************************************
Public Sub CM_IsTopBraceAC1Needed(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "::CM_IsTopBraceAC1Needed"
    On Error GoTo ErrorHandler
    
    Dim lBracedispID As Long
    lBracedispID = 12 'Dispid of top child brace
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    bIsNeeded = False
    
    If Not oMemObjs.ItemByDispid(lBracedispID) Is Nothing Then
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'*********************************************************************************************
' Method      : CM_IsTopBraceAC2Needed
' Description : Conditional method to determine if an AC between brace and Bounding(of the Parent AC) is needed
' Creates AC when the brace exists
'********************************************************************************************
Public Sub CM_IsTopBraceAC2Needed(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "::CM_IsTopBraceAC1Needed"
    On Error GoTo ErrorHandler
    
    Dim lBracedispID As Long
    lBracedispID = 12 'Dispid of top child brace
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    bIsNeeded = False
    
    If Not oMemObjs.ItemByDispid(lBracedispID) Is Nothing Then
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomBraceAC1Needed
' Description : Conditional method to determine if an AC between brace and Bounded(of the Parent AC) is needed
' Creates AC when the brace exists
'********************************************************************************************
Public Sub CM_IsBottomBraceAC1Needed(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "::CM_IsTopBraceAC1Needed"
    On Error GoTo ErrorHandler
    
    Dim lBracedispID As Long
    lBracedispID = 15 'Dispid of bottom child brace
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    bIsNeeded = False
    
    If Not oMemObjs.ItemByDispid(lBracedispID) Is Nothing Then
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub


'*********************************************************************************************
' Method      : CM_IsBottomBraceAC2Needed
' Description : Conditional method to determine if an AC between brace and Bounding(of the Parent AC) is needed
' Creates AC when the brace exists
'********************************************************************************************
Public Sub CM_IsBottomBraceAC2Needed(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "::CM_IsTopBraceAC1Needed"
    On Error GoTo ErrorHandler
    
    Dim lBracedispID As Long
    lBracedispID = 15   'Dispid of bottom child brace
    
    Dim oMemObjs As IJDMemberObjects
    Set oMemObjs = pMemberDescription.CAO
    
    bIsNeeded = False
    
    If Not oMemObjs.ItemByDispid(lBracedispID) Is Nothing Then
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
' Method      : CMConstructTopBrace
' Description : Construct method to create Top brace
'********************************************************************************************
Public Function CMConstructTopBrace(ByVal pMemberDescription As IJDMemberDescription, _
                                         ByVal pResourceManager As IUnknown, _
                                         ByRef pObject As Object)
Const METHOD = "CMConstructTopBrace"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    
    lDispId = pMemberDescription.dispid
        
    ConstructMemberBrace pMemberDescription, pResourceManager, pObject, True
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*********************************************************************************************
' Method      : CMConstructBottomBrace
' Description : Construct method to create Bottom brace
'********************************************************************************************
Public Function CMConstructBottomBrace(ByVal pMemberDescription As IJDMemberDescription, _
                                         ByVal pResourceManager As IUnknown, _
                                         ByRef pObject As Object)
Const METHOD = "CMConstructBottomBrace"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    
    lDispId = pMemberDescription.dispid
    
    ConstructMemberBrace pMemberDescription, pResourceManager, pObject, False
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*********************************************************************************************
'*********************************************************************************************
' Method      : ConstructMemberBrace
' Description : This method gets all the info needed to create the InsetBrace
' Gets the Top/Bottom(bounded) Intersecting edges of Bounding
' Determines the Brace Length : If its a Top brace, get the height between Top of the Bounded & Top(alias) of Bounding
'       If its a Bottom brace, get the height between Bottom of the Bounded & Bottom(alias) of Bounding
'       BraceSlope determines the length of the brace. The X component of brace should be equal to (BraceSlope * Height determined above)
' Determine the Brace X component: same as Axis direction of the Bounded(of the Parent AC)
' Determine the Connect points of the Brace (Start and end points of the brace)
' Usually InsetBrace is cut from the same cross section that of the Bounded of the parent AC. But, if the InsetBrace doesnt fill the gap
'       between Bounded and Bounding, then the cross section is queried from the catalog based on some requirements.
'       Minimum brace height is caluclated and the cross section is queried for.
' Using the Connect points, Cross section, Create a member brace.
' bIsTopBrace is true if the brace is connected to the Top/FL(Flange penetrated) of the Bounded
Public Sub ConstructMemberBrace(ByVal pMemberDescription As IJDMemberDescription, ByRef pResourceManager As IUnknown, ByRef pObject As Object, bIsTopBrace As Boolean)
Const METHOD = "MbrAxisToCenterDef::ConstructMemberBrace"
    
    Dim sMsg As String
    
    Dim sBoundedEndType As String, sBoundingEndType As String, sBtmBraceType As String, sTopBraceType As String
    Dim dBoundedEndValue As Double, dBoundingEndValue As Double
    'Get the connect points of the brace
    'if web penetrated
    If bIsTopBrace Then
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceBoundedEndType", sBoundedEndType
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceBoundedEndValue", dBoundedEndValue
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceBoundingEndType", sBoundingEndType
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceBoundingEndValue", dBoundingEndValue
        GetSelectorAnswer pMemberDescription.CAO, "TopBraceType", sTopBraceType
    Else
        GetSelectorAnswer pMemberDescription.CAO, "BottomBraceBoundedEndType", sBoundedEndType
        GetSelectorAnswer pMemberDescription.CAO, "BottomBraceBoundedEndValue", dBoundedEndValue
        GetSelectorAnswer pMemberDescription.CAO, "BottomBraceBoundingEndType", sBoundingEndType
        GetSelectorAnswer pMemberDescription.CAO, "BottomBraceBoundingEndValue", dBoundingEndValue
        GetSelectorAnswer pMemberDescription.CAO, "BottomBraceType", sBtmBraceType
    End If
           
    'we have to determine whether a top/bottom brace should be created
    'Get the Edge Mapping and Section Alias
    Dim bPenetratesWeb As Boolean
    Dim bIsTopFlange As Boolean
    Dim oAppConnection As IJAppConnection
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Dim oSectionAlias As Long
    
    Set oEdgeMapColl = New Collection
    
    bIsTopFlange = True
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    
    ' Set the Bounded and Bounding Port
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
                         
    Set oBoundedPort = oBoundedData.AxisPort
    Set oBoundingPort = oBoundingData.AxisPort
    
    'Get the Position along Bounding
    Dim oSDO_Bounding As New StructDetailObjects.MemberPart, oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounding.object = oBoundingPort.Connectable: Set oSDO_Bounded.object = oBoundedPort.Connectable

    Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, oSectionAlias, bPenetratesWeb)
    
    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
    Set oMeasurements = New Collection

    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements

    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR

    'oEdge1 and oEdge2 should be Top/Bottom incase of Web-Penetrated and WL/WR incase of Flange-penetarted inorder to get the intersecting edges
    Dim oEdge1 As ConnectedEdgeInfo
    Dim oEdge2 As ConnectedEdgeInfo

    oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
    oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
  
    ' Gets the Top/Bottom(bounded) Intersecting edges of Bounding
    Dim oTopOrWLIntersectingEdge As JXSEC_CODE, oBtmOrWRIntersectingEdge As JXSEC_CODE
    GetIntersectingEdgeID oEdge1, oTopOrWLIntersectingEdge
    GetIntersectingEdgeID oEdge2, oBtmOrWRIntersectingEdge
    
    'Intersect the edgeID with the Top/Bottom based on Top/Bottom brace
    
    Dim oTop_Bounded As IJPort, oBtm_Bounded As IJPort
    Set oTop_Bounded = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_TOP)
    Set oBtm_Bounded = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_BOTTOM)
    
    Dim oTop_Extended_Bounded As Object, oBtm_Extended_Bounded As Object
    Set oTop_Extended_Bounded = GetExtendedPort(oTop_Bounded)
    Set oBtm_Extended_Bounded = GetExtendedPort(oBtm_Bounded)
    
    Dim oTopOrWLIntersectingEdgePort As IJPort, oBtmOrWRIntersectingEdgePort As IJPort
    
    Dim pAgtorUnk As IUnknown, oIntersectedUnknown As IUnknown, oIntersectionWire As IJWireBody
    Dim oGeomOps As New IMSModelGeomOps.DGeomOpsIntersect
    Dim oWUtility As New GSCADShipGeomOps.SGOWireBodyUtilities
    
    Dim dMinBraceHeight As Double, dYComp_Brace As Double
    
    '*********************************************************************************************
    '*********************************************************************************************
    'Get the intersection of Top/Bottom of Bounded with the BoundingIntersecting Edge
    
    'If Web penetrated
    If bPenetratesWeb Then
        If bIsTopBrace Then
             'get the Top.Intersecting edge
             Set oTopOrWLIntersectingEdgePort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(oTopOrWLIntersectingEdge, oEdgeMapColl))
             oGeomOps.PlaceIntersectionObject Nothing, oTopOrWLIntersectingEdgePort, oTop_Extended_Bounded, pAgtorUnk, oIntersectedUnknown
         Else    'bottom brace
            'get the btm.Intersecting edge
            Set oBtmOrWRIntersectingEdgePort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(oBtmOrWRIntersectingEdge, oEdgeMapColl))
            oGeomOps.PlaceIntersectionObject Nothing, oBtmOrWRIntersectingEdgePort, oBtm_Extended_Bounded, pAgtorUnk, oIntersectedUnknown
        End If
    Else 'For Flange penetrated, always take WL inetrsecting edge and intersect with Top/Bottom based on Top/Bottom brace
    
        Set oTopOrWLIntersectingEdgePort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(oTopOrWLIntersectingEdge, oEdgeMapColl))
        
        If bIsTopBrace Then
             oGeomOps.PlaceIntersectionObject Nothing, oTopOrWLIntersectingEdgePort, oTop_Extended_Bounded, pAgtorUnk, oIntersectedUnknown
        Else 'bottom brace
            oGeomOps.PlaceIntersectionObject Nothing, oTopOrWLIntersectingEdgePort, oBtm_Extended_Bounded, pAgtorUnk, oIntersectedUnknown
        End If
    End If
    
    'Get the vector for Bounded U axis away from the bounding
        
    'Get the Position along Bounded
    'Get the vector along Bounded away from AC position
    Dim oBraceXVector As IJDVector
    Set oBraceXVector = New dVector

    Dim oBoundedPart As ISPSMemberPartPrismatic
    Set oBoundedPart = oBoundedPort.Connectable

    'get the bounded aixs
    Dim oBoundedCurve As IJCurve
    Set oBoundedCurve = oBoundedPart.Axis

    Dim startParam As Double, endParam As Double
    Dim posX As Double, posY As Double, posZ As Double
    Dim tanX As Double, tanY As Double, tanZ As Double
    Dim curX As Double, curY As Double, curZ As Double

    oBoundedCurve.ParamRange startParam, endParam

    Dim SplitAxisPort As ISPSSplitAxisPort
    Set SplitAxisPort = oBoundedPort

    Dim MemberAxisPortIndex As SPSMemberAxisPortIndex
    MemberAxisPortIndex = SplitAxisPort.PortIndex
        
    If MemberAxisPortIndex = SPSMemberAxisStart Then
        ' Negate the u-vector if the port is at the start, so that it points away from the member
        oBoundedCurve.Evaluate startParam, posX, posY, posZ, tanX, tanY, tanZ, curX, curY, curZ
        tanX = -tanX
        tanY = -tanY
        tanZ = -tanZ
    ElseIf MemberAxisPortIndex = SPSMemberAxisEnd Then
        oBoundedCurve.Evaluate endParam, posX, posY, posZ, tanX, tanY, tanZ, curX, curY, curZ
    End If

    'get the vector that represents the direction of bounded from AC position towards the member other end
    oBraceXVector.Set -tanX, -tanY, -tanZ
    
    'Get the Skecth plane to intersect the oIntersectionWire to get the intersection position of Top/Bottom of Bounded with the Bounding intersecting edge
    Dim oIntersectionPos As IJDPosition: Set oIntersectionPos = New DPosition
    
    Dim oSketchPlane As IJPlane
    Set oSketchPlane = New Plane3d
    
    If bPenetratesWeb Then
         'Get the web sketching plane
        GetSketchPlaneForMSSymbol oBoundedPort, oBoundingPort, bPenetratesWeb, bIsTopFlange, oSketchPlane, False
    Else
        'Get the normal of Web left of Bounded
        Dim oWL_Bounded As IJPort, oWL_Extended_Bounded As Object, oWLGeom As IJSurfaceBody
        Dim oWLNormal As IJDVector: Set oWLNormal = New dVector
        Dim SketchPosOffset As IJDVector, oSketchPosition As IJDPosition
        Set SketchPosOffset = New dVector: Set oSketchPosition = New DPosition
        
'        Dim oStructGeomUtil As New GSCADStructGeomUtilities.PartInfo
    
        Set oWL_Bounded = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_WEB_LEFT)
        Set oWL_Extended_Bounded = GetExtendedPort(oWL_Bounded)
        Set oWLGeom = oWL_Bounded.Geometry

'        Set oWLNormal = oStructGeomUtil.GetPortNormal(oWL_Bounded, True)
        
        'Get the position on WebLeft
        Dim oBoundedPos As IJDPosition, pDist As Double, oWL_Pos As IJDPosition
        Dim px As Double, py As Double, pz As Double
                
        Dim oModelBodyUtility As IJSGOModelBodyUtilities
        Set oModelBodyUtility = New SGOModelBodyUtilities
        
        Set oBoundedPos = New DPosition: Set oWL_Pos = New DPosition
        oBoundedPart.PointAtEnd(MemberAxisPortIndex).GetPoint px, py, pz
        
        oBoundedPos.Set px, py, pz
        
        oModelBodyUtility.GetClosestPointOnBody oWL_Bounded.Geometry, oBoundedPos, oWL_Pos, pDist
        oWLGeom.GetNormalFromPosition oWL_Pos, oWLNormal
        
        SketchPosOffset.Set -oWLNormal.x, -oWLNormal.y, -oWLNormal.z
        SketchPosOffset.Length = oSDO_Bounded.webThickness / 2
        Set oSketchPosition = oWL_Pos.Offset(SketchPosOffset)
        
        
        'set the sketch plane
        oSketchPlane.SetNormal oWLNormal.x, oWLNormal.y, oWLNormal.z
        oSketchPlane.SetRootPoint oSketchPosition.x, oSketchPosition.y, oSketchPosition.z
    End If
    'Get the intersection
    If TypeOf oIntersectedUnknown Is IJWireBody Then
        Set oIntersectionWire = oIntersectedUnknown
        Set oIntersectionPos = oWUtility.FindSinglePointIntersection(oIntersectionWire, oSketchPlane)
    End If
    
    '*********************************************************************************************
    '*********************************************************************************************
    Dim oFinalPos_Bounded As IJDPosition, oFinalPos_Bounding As IJDPosition
    Set oFinalPos_Bounded = New DPosition: Set oFinalPos_Bounding = New DPosition

    Dim dDistTopBoundedToTopOfBounding As Double, dDistBtmToBtmOfBounding As Double
    If bPenetratesWeb Then
        dDistTopBoundedToTopOfBounding = oMeasurements.Item("DimPt15ToTop")
        dDistBtmToBtmOfBounding = oMeasurements.Item("DimPt23ToBottom")
    End If
    
    
'    If InStr(1, sBoundedEndType, "Structure", vbTextCompare) > 0 Then
'        oBraceXVector.Length = dBoundedEndValue
    If InStr(1, sBoundedEndType, "Offset", vbTextCompare) > 0 Then
        oBraceXVector.Length = dBoundedEndValue
    ElseIf InStr(1, sBoundedEndType, "Slope", vbTextCompare) > 0 Then
        'get the Bounding offset /Structure and calculate slope
        If bPenetratesWeb Then
            If bIsTopBrace Then
                If GreaterThan(dBoundingEndValue, dDistTopBoundedToTopOfBounding) Then
                    dBoundingEndValue = dDistTopBoundedToTopOfBounding
                End If
            Else
                If GreaterThan(dBoundingEndValue, dDistBtmToBtmOfBounding) Then
                    dBoundingEndValue = dDistBtmToBtmOfBounding
                End If
            End If
            
            If InStr(1, sBoundingEndType, "Offset", vbTextCompare) > 0 Then
                oBraceXVector.Length = dBoundingEndValue * dBoundedEndValue
            End If
        Else    'assuming Bounding has Offset/Structure (As Structure is not implemented for FP cases now
            If dBoundingEndValue > (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) Then
                dBoundingEndValue = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1)
            End If
            oBraceXVector.Length = dBoundingEndValue * dBoundedEndValue
        End If
    End If
    
    Set oFinalPos_Bounded = oIntersectionPos.Offset(oBraceXVector)
    
    '***********************************************************************************
    '***********************************************************************************
    Dim dFT_Bounded As Double, dFT_Bounding As Double
    dFT_Bounded = oSDO_Bounded.flangeThickness
    dFT_Bounding = oSDO_Bounding.flangeThickness
    
    Dim oTFRCornerEdge As IJWireBody, oTFRBCornerEdge As IJWireBody, oTFRB_WR_CornerEdge As IJWireBody, oTop_WR_CornerEdge As IJWireBody
    Dim oPos_TFRT_Corner As IJDPosition, oPos_TFRB_Corner As IJDPosition, oPos_TFRB_WR_Corner As IJDPosition, oPos_Top_WR_Corner As IJDPosition
    Set oPos_TFRB_Corner = New DPosition: Set oPos_TFRT_Corner = New DPosition: Set oPos_TFRB_WR_Corner = New DPosition: Set oPos_Top_WR_Corner = New DPosition
    Dim oTopProjVector As IJDVector
    Set oTopProjVector = New dVector
    
    Dim oBFRCornerEdge As IJWireBody, oBFRTCornerEdge As IJWireBody, oBFRT_WR_CornerEdge As IJWireBody, oBtm_WR_CornerEdge As IJWireBody
    Dim oPos_BFRB_Corner As IJDPosition, oPos_BFRT_Corner As IJDPosition, oPos_BFRT_WR_Corner As IJDPosition, oPos_Btm_WR_Corner As IJDPosition
    Set oPos_BFRB_Corner = New DPosition: Set oPos_BFRT_Corner = New DPosition: Set oPos_BFRT_WR_Corner = New DPosition: Set oPos_Btm_WR_Corner = New DPosition
    Dim oBtmProjVector As IJDVector
    Set oBtmProjVector = New dVector
    
    Dim oBounding_UpDir As IJDVector: Set oBounding_UpDir = New dVector
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    oBounding_UpDir.Set oBoundingData.Matrix.IndexValue(8), oBoundingData.Matrix.IndexValue(9), _
        oBoundingData.Matrix.IndexValue(10)
        
    Dim oBounded_UpDir As IJDVector: Set oBounded_UpDir = New dVector
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    oBounded_UpDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), _
        oBoundedData.Matrix.IndexValue(10)
        
    Dim oBounding_AxisDir As IJDVector: Set oBounding_AxisDir = New dVector
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    oBounding_AxisDir.Set oBoundingData.Matrix.IndexValue(0), oBoundingData.Matrix.IndexValue(1), _
            oBoundingData.Matrix.IndexValue(2)
        
    If bIsTopBrace Then
        If oBounded_UpDir.Dot(oBounding_UpDir) > 0 Then
             oTopProjVector.Set oBounding_UpDir.x, oBounding_UpDir.y, oBounding_UpDir.z
        Else
             oTopProjVector.Set -oBounding_UpDir.x, -oBounding_UpDir.y, -oBounding_UpDir.z
        End If
    Else
        If oBounded_UpDir.Dot(oBounding_UpDir) > 0 Then
             oBtmProjVector.Set -oBounding_UpDir.x, -oBounding_UpDir.y, -oBounding_UpDir.z
        Else
             oBtmProjVector.Set oBounding_UpDir.x, oBounding_UpDir.y, oBounding_UpDir.z
        End If
    End If
    
    Dim dBoundingOffset As Double
    
    If InStr(1, sBoundingEndType, "Slope", vbTextCompare) > 0 Then
        dBoundingOffset = dBoundingEndValue * dBoundedEndValue
    ElseIf InStr(1, sBoundingEndType, "Offset", vbTextCompare) > 0 Then
        dBoundingOffset = dBoundingEndValue
    End If
    
    If bPenetratesWeb Then
        Dim oTop_Bounding As IJPort, oTop_TFR_Bounding As IJPort, oTop_TFRB_Bounding As IJPort, oWebRight_Bounding As IJPort
        Dim oBottom_Bounding As IJPort, oBtm_BFR_Bounding As IJPort, oBtm_BFRT_Bounding As IJPort
        
        Set oTop_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_TOP, oEdgeMapColl))
        Set oTop_TFR_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl))
        Set oTop_TFRB_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_TOP_FLANGE_RIGHT_BOTTOM, oEdgeMapColl))
        Set oWebRight_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_WEB_RIGHT, oEdgeMapColl))
        Set oBottom_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_BOTTOM, oEdgeMapColl))
        Set oBtm_BFR_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl))
        Set oBtm_BFRT_Bounding = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_BOTTOM_FLANGE_RIGHT_TOP, oEdgeMapColl))
        
        Dim pTop_TFR_Intersect As IUnknown, pTFR_TFRB_Intersect As IUnknown, pTFRB_WR_Intersect As IUnknown, pTop_WR_Intersect As IUnknown, pTop_WLIntersect_Intersect As IUnknown, pBtm_WLIntersect_Intersect As IUnknown
        Dim pBtm_BFR_Intersect As IUnknown, pBFR_BFRT_Intersect As IUnknown, pBFRT_WR_Intersect As IUnknown, pBtm_WR_Intersect As IUnknown
            
        'Get the Connect point of the brace on Bounding of the Parent AC
        If bIsTopBrace Then 'if top brace
                    
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl) Then
                oGeomOps.PlaceIntersectionObject Nothing, oTop_Bounding, oTop_TFR_Bounding, pAgtorUnk, pTop_TFR_Intersect
                
                If TypeOf pTop_TFR_Intersect Is IJWireBody Then
                    Set oTFRCornerEdge = pTop_TFR_Intersect
                    Set oPos_TFRT_Corner = oWUtility.FindSinglePointIntersection(oTFRCornerEdge, oSketchPlane)
                End If
                
                If GreaterThanOrEqualTo(dBoundingOffset, dDistTopBoundedToTopOfBounding) Then
                    Set oFinalPos_Bounding = oPos_TFRT_Corner
                    dYComp_Brace = dBoundingOffset
                ElseIf LessThan(dBoundingOffset, dDistTopBoundedToTopOfBounding) And LessThan(dBoundingOffset, (dDistTopBoundedToTopOfBounding - oSDO_Bounding.flangeThickness)) Then
                    
                    oTopProjVector.Length = dBoundingOffset
                    Set oFinalPos_Bounding = oIntersectionPos.Offset(oTopProjVector)
                    
                    dYComp_Brace = dBoundingOffset
                ElseIf LessThan(dBoundingOffset, dDistTopBoundedToTopOfBounding) And GreaterThan(dBoundingOffset, (dDistTopBoundedToTopOfBounding - oSDO_Bounding.flangeThickness)) Then
                     
                    oGeomOps.PlaceIntersectionObject Nothing, oTop_TFRB_Bounding, oTop_TFR_Bounding, pAgtorUnk, pTFR_TFRB_Intersect
                    
                    If TypeOf pTFR_TFRB_Intersect Is IJWireBody Then
                        Set oTFRBCornerEdge = pTFR_TFRB_Intersect
                        Set oPos_TFRB_Corner = oWUtility.FindSinglePointIntersection(oTFRBCornerEdge, oSketchPlane)
                    End If
                    
                    oTopProjVector.Length = dBoundingOffset - (dDistTopBoundedToTopOfBounding - oSDO_Bounding.flangeThickness)
                    dYComp_Brace = oTopProjVector.Length
                    Set oFinalPos_Bounding = oPos_TFRB_Corner.Offset(oTopProjVector)
                End If
                     
            Else    'Alias doesnt have a Top Flange
                'Proj vector???
                oGeomOps.PlaceIntersectionObject Nothing, oTop_Bounding, oWebRight_Bounding, pAgtorUnk, pTop_WR_Intersect
                If TypeOf pTop_WR_Intersect Is IJWireBody Then
                    Set oTop_WR_CornerEdge = pTop_WR_Intersect
                    Set oPos_Top_WR_Corner = oWUtility.FindSinglePointIntersection(oTop_WR_CornerEdge, oSketchPlane)
                End If
                
                If GreaterThanOrEqualTo(dBoundingOffset, dDistTopBoundedToTopOfBounding) Then
                    Set oFinalPos_Bounding = oPos_Top_WR_Corner
                    dYComp_Brace = dBoundingOffset
                Else
                    oTopProjVector.Length = dBoundingOffset
                    dYComp_Brace = dBoundingOffset
                    Set oFinalPos_Bounding = oIntersectionPos.Offset(oTopProjVector)
                End If
            End If
        Else 'if bottom brace
            
            If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl) Then
            
                oGeomOps.PlaceIntersectionObject Nothing, oBottom_Bounding, oBtm_BFR_Bounding, pAgtorUnk, pBtm_BFR_Intersect
                
                If TypeOf pBtm_BFR_Intersect Is IJWireBody Then
                    Set oBFRCornerEdge = pBtm_BFR_Intersect
                    Set oPos_BFRB_Corner = oWUtility.FindSinglePointIntersection(oBFRCornerEdge, oSketchPlane)
                End If
                
                If GreaterThanOrEqualTo(dBoundingOffset, dDistBtmToBtmOfBounding) Then
                    dYComp_Brace = dBoundingOffset
                    Set oFinalPos_Bounding = oPos_BFRB_Corner
                ElseIf LessThan(dBoundingOffset, dDistBtmToBtmOfBounding) And LessThan(dBoundingOffset, (dDistBtmToBtmOfBounding - oSDO_Bounding.flangeThickness)) Then
                        
                    oBtmProjVector.Length = dBoundingOffset
                    dYComp_Brace = dBoundingOffset
                    Set oFinalPos_Bounding = oIntersectionPos.Offset(oBtmProjVector)
                ElseIf LessThan(dBoundingOffset, dDistBtmToBtmOfBounding) And GreaterThan(dBoundingOffset, (dDistBtmToBtmOfBounding - oSDO_Bounding.flangeThickness)) Then
                         
                    oGeomOps.PlaceIntersectionObject Nothing, oBtm_BFR_Bounding, oBtm_BFRT_Bounding, pAgtorUnk, pBFR_BFRT_Intersect
                    
                    If TypeOf pBFR_BFRT_Intersect Is IJWireBody Then
                        Set oBFRTCornerEdge = pBFR_BFRT_Intersect
                        Set oPos_BFRT_Corner = oWUtility.FindSinglePointIntersection(oBFRTCornerEdge, oSketchPlane)
                    End If
                    
                    oBtmProjVector.Length = dBoundingOffset - (dDistBtmToBtmOfBounding - oSDO_Bounding.flangeThickness)
                    dYComp_Brace = oBtmProjVector.Length
                    Set oFinalPos_Bounding = oPos_BFRT_Corner.Offset(oBtmProjVector)
                End If
                
            Else    'Alias doesnt have a Btm Flange
                'Proj vector???
                oGeomOps.PlaceIntersectionObject Nothing, oBottom_Bounding, oWebRight_Bounding, pAgtorUnk, pBtm_WR_Intersect
                If TypeOf pBtm_WR_Intersect Is IJWireBody Then
                    Set oBtm_WR_CornerEdge = pBtm_WR_Intersect
                    Set oPos_Btm_WR_Corner = oWUtility.FindSinglePointIntersection(oBtm_WR_CornerEdge, oSketchPlane)
                End If
                
                If GreaterThanOrEqualTo(dBoundingOffset, dDistBtmToBtmOfBounding) Then
                    dYComp_Brace = dBoundingOffset
                    Set oFinalPos_Bounding = oPos_Btm_WR_Corner
                Else
                    oBtmProjVector.Length = dBoundingOffset
                    dYComp_Brace = dBoundingOffset
                    Set oFinalPos_Bounding = oIntersectionPos.Offset(oBtmProjVector)
                End If
            End If
        End If
        
    Else    'Flange penetrated (Both Bounded/bounding connect points are different
    
        Dim oVectorAlongBounding As IJDVector
        Set oVectorAlongBounding = New dVector
    
        If bIsTopBrace Then
            If oBounded_UpDir.Dot(oBounding_AxisDir) > 0 Then
                oVectorAlongBounding.Set oBounding_AxisDir.x, oBounding_AxisDir.y, oBounding_AxisDir.z
            Else
                oVectorAlongBounding.Set -oBounding_AxisDir.x, -oBounding_AxisDir.y, -oBounding_AxisDir.z
            End If
            
        Else
            If oBounded_UpDir.Dot(oBounding_AxisDir) < 0 Then
                oVectorAlongBounding.Set oBounding_AxisDir.x, oBounding_AxisDir.y, oBounding_AxisDir.z
            Else
                oVectorAlongBounding.Set -oBounding_AxisDir.x, -oBounding_AxisDir.y, -oBounding_AxisDir.z
            End If
        End If
    
        If InStr(1, sBoundingEndType, "Offset", vbTextCompare) > 0 Then
            
            If dBoundingEndValue > (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) Then
                dBoundingEndValue = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1)
            End If
            oVectorAlongBounding.Length = dBoundingEndValue
            
            dYComp_Brace = oVectorAlongBounding.Length
            Set oFinalPos_Bounding = oIntersectionPos.Offset(oVectorAlongBounding)
            
        ElseIf InStr(1, sBoundingEndType, "Slope", vbTextCompare) > 0 Then
            
            oVectorAlongBounding.Length = dBoundingEndValue * dBoundedEndValue
            
            If GreaterThan(oVectorAlongBounding.Length, (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1)) Then
                oVectorAlongBounding.Length = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1)
            End If
            
            dYComp_Brace = oVectorAlongBounding.Length
            Set oFinalPos_Bounding = oIntersectionPos.Offset(oVectorAlongBounding)
            
        End If
    
    End If
    
    Dim oBraceSlope As Double
    
    'use brace end points and set it to bounding axis to find out the angle
    Dim oBraceDir As IJDVector
    Set oBraceDir = New dVector
    
    Set oBraceDir = oFinalPos_Bounded.Subtract(oFinalPos_Bounding)
    
    Dim oBounded_WebDir As IJDVector, oBounded_AxisDir As IJDVector
    Set oBounded_WebDir = New dVector: Set oBounded_AxisDir = New dVector
    
    oBounded_AxisDir.Set oBoundedData.Matrix.IndexValue(0), oBoundedData.Matrix.IndexValue(1), oBoundedData.Matrix.IndexValue(2)
    oBounded_WebDir.Set oBoundedData.Matrix.IndexValue(4), oBoundedData.Matrix.IndexValue(5), oBoundedData.Matrix.IndexValue(6)
    
    Dim dAngle As Double
    ' Adjust to 0-180
    Dim dPI As Double
    dPI = Atn(1#) * 4
    
    dAngle = oBounded_AxisDir.Angle(oBraceDir, oBounded_WebDir)
    If 0 < dAngle < dPI / 2 Then
    ElseIf (dPI / 2) < dAngle < dPI Then
        dAngle = dPI - dAngle
    ElseIf dPI < dAngle < (3 * dPI / 2) Then
        dAngle = dAngle - dPI
    ElseIf (3 * dPI / 2) < dAngle < 2 * dPI Then
        dAngle = (2 * dPI) - dAngle
    End If
    
    If bIsTopBrace Then
        If sTopBraceType = "InsetMember" Then
            dMinBraceHeight = (dYComp_Brace + 0.1) * Cos(dAngle)
        ElseIf sTopBraceType = "FlushMember" Then
            dMinBraceHeight = (dYComp_Brace) * Cos(dAngle)
        End If
        
        If bPenetratesWeb Then
            If ItemExists(JXSEC_TOP_FLANGE_RIGHT, oEdgeMapColl) Then
                If dYComp_Brace < (dDistTopBoundedToTopOfBounding - oSDO_Bounding.flangeThickness) Then
                    'do nothing
                Else
                    dMinBraceHeight = dMinBraceHeight + (((oSDO_Bounding.FlangeLength - oSDO_Bounding.webThickness) / 2) * Sin(dAngle))
                End If
            End If
        End If
                
    Else
        If sBtmBraceType = "InsetMember" Then
            dMinBraceHeight = (dYComp_Brace + 0.01) * Cos(dAngle)
        ElseIf sBtmBraceType = "FlushMember" Then
            dMinBraceHeight = (dYComp_Brace) * Cos(dAngle)
        End If
        
        If bPenetratesWeb Then
            If ItemExists(JXSEC_BOTTOM_FLANGE_RIGHT, oEdgeMapColl) Then
                If dYComp_Brace < (dDistBtmToBtmOfBounding - oSDO_Bounding.flangeThickness) Then
                    'do nothing
                Else
                    dMinBraceHeight = dMinBraceHeight + (((oSDO_Bounding.FlangeLength - oSDO_Bounding.webThickness) / 2) * Sin(dAngle))
                End If
            End If
        End If
    End If
    
    'Get the cross section
    Dim oCrossSection As IJCrossSection
    GetBraceCrossSection oSDO_Bounded, oSDO_Bounding, oCrossSection, dMinBraceHeight
    
    'Create a stand alone member part
    CreateMemberBrace pMemberDescription, pResourceManager, oFinalPos_Bounding, oFinalPos_Bounded, oCrossSection, pObject, bIsTopBrace, bPenetratesWeb
        
     Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Sub

Public Function CMConstructSeatedPC(ByVal pMD As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    ByRef pObject As Object)
                                 
    Const METHOD = "MbrACDefCM::CMConstructSeatedPC"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating seated PC...pMD.index = " & Str(pMD.Index)
    
    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Set oAppConnection = pMD.CAO
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
      
    ' --------------------------
    ' Get the relative positions
    ' --------------------------
    Dim cWLOrTop As ConnectedEdgeInfo
    Dim cWROrBtm As ConnectedEdgeInfo
    Dim cTFIOrFL As ConnectedEdgeInfo
    Dim cBFIOrFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, cWLOrTop, cWROrBtm, cTFIOrFL, cBFIOrFR

    ' -----------------
    ' If web penetrated
    ' -----------------
    Dim boundingConnPortXid As JXSEC_CODE
    Dim boundedConnPortXid As JXSEC_CODE
    
    boundingConnPortXid = JXSEC_UNKNOWN
    boundedConnPortXid = JXSEC_UNKNOWN
    
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim sWeldSel As String

    If bPenetratesWeb Then
        ' -------------------------------------------------------------------------
        ' Either the top touches bottom or the bottom touches top
        ' Weld type depends on whether or not there is a flange, and if rectangular
        ' -------------------------------------------------------------------------
        If cWLOrTop.CoplanarEdge = eBounding_Edge.Bottom Then
            boundedConnPortXid = JXSEC_TOP
            boundingConnPortXid = JXSEC_BOTTOM
            
            If HasTopFlange(oBoundedData.MemberPart) Then
                sWeldSel = "LapWeld"
            Else
                sWeldSel = "TeeWeld"
            End If
            
        ElseIf cWROrBtm.CoplanarEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_BOTTOM
            boundingConnPortXid = JXSEC_TOP
            
            If HasBottomFlange(oBoundedData.MemberPart) Then
                sWeldSel = "LapWeld"
            Else
                sWeldSel = "TeeWeld"
            End If
            
        End If
    ' --------------------
    ' If flange penetrated
    ' --------------------
    Else
        ' -----------------------------------------------------------
        ' If this is the first PC and the web left touches the bottom
        ' -----------------------------------------------------------
        If pMD.dispid = 20 And cWLOrTop.CoplanarEdge = eBounding_Edge.Bottom Then
            boundedConnPortXid = JXSEC_WEB_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
            sWeldSel = "LapWeld"
        ' ---------------------------------------------------------
        ' If this is the first PC and the web right touches the top
        ' ---------------------------------------------------------
        ElseIf pMD.dispid = 20 And cWROrBtm.CoplanarEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_WEB_RIGHT
            boundingConnPortXid = JXSEC_TOP
            sWeldSel = "LapWeld"
        ' --------------------------------------------------------------
        ' If this is the first PC and the flange left touches the bottom
        ' --------------------------------------------------------------
        ElseIf pMD.dispid = 20 And cTFIOrFL.CoplanarEdge = eBounding_Edge.Bottom Then
            boundedConnPortXid = JXSEC_TOP_FLANGE_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
            sWeldSel = "TeeWeld"
        ' ------------------------------------------------------------
        ' If this is the first PC and the flange right touches the top
        ' ------------------------------------------------------------
        ElseIf pMD.dispid = 20 And cBFIOrFR.CoplanarEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_TOP_FLANGE_RIGHT
            boundingConnPortXid = JXSEC_TOP
            sWeldSel = "TeeWeld"
        ' ---------------------------------------------------------------
        ' If this is the second PC and the flange left touches the bottom
        ' ---------------------------------------------------------------
        ElseIf pMD.dispid = 21 And cTFIOrFL.CoplanarEdge = eBounding_Edge.Bottom Then
            boundedConnPortXid = JXSEC_BOTTOM_FLANGE_LEFT
            boundingConnPortXid = JXSEC_BOTTOM
            sWeldSel = "TeeWeld"
        ' -------------------------------------------------------------
        ' If this is the second PC and the flange right touches the top
        ' -------------------------------------------------------------
        ElseIf pMD.dispid = 21 And cBFIOrFR.CoplanarEdge = eBounding_Edge.Top Then
            boundedConnPortXid = JXSEC_BOTTOM_FLANGE_RIGHT
            boundingConnPortXid = JXSEC_TOP
            sWeldSel = "TeeWeld"
        End If
    End If
    
    ' --------------------
    ' Get the edge mapping
    ' --------------------
    'Get the Edge Mapping and Section Alias
    Dim pSectionAlias As Long
    Dim oEdgeMap As JCmnShp_CollectionAlias
    Set oEdgeMap = New Collection

    Set oEdgeMap = GetEdgeMap(oAppConnection, oBoundingData.AxisPort, oBoundedData.AxisPort, pSectionAlias, bPenetratesWeb)
    ' ------------------------------------------
    ' Get connectable ports on the last geometry
    ' ------------------------------------------
    If boundedConnPortXid = JXSEC_UNKNOWN Or boundingConnPortXid = JXSEC_UNKNOWN Then
        sMsg = "Could not identify ports for connection"
        GoTo ErrorHandler
    End If
    
    Dim oBoundedConnPort As IJPort
    Dim oBoundingConnPort As IJPort
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim oLateBindingPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oStructSymbolTools As IJStructSymbolTools
    Set oStructSymbolTools = New StructSymbolTools
    
    If pHelper.ObjectType(oBoundedData.MemberPart) = SDOBJECT_STIFFENER Then
        Set oProfilePart.object = oBoundedData.MemberPart
        Set oBoundedConnPort = oProfilePart.SubPort(boundedConnPortXid)
    ElseIf pHelper.ObjectType(oBoundedData.MemberPart) = SDOBJECT_MEMBER Then
        Set oBoundedConnPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, boundedConnPortXid)
        Set oStructPort = oBoundedConnPort
        oStructSymbolTools.GetLateBindingPort oBoundedData.MemberPart, oStructPort.ProxyType, oStructPort.OperationID, oStructPort.OperatorID, oStructPort.ContextID, oStructPort.OperatorID, oLateBindingPort
        Set oBoundedConnPort = oLateBindingPort
    End If

    Dim RealEdgeID As JXSEC_CODE
    If ItemExists(boundingConnPortXid, oEdgeMap, RealEdgeID) Then
        boundingConnPortXid = RealEdgeID
        If pHelper.ObjectType(oBoundingData.MemberPart) = SDOBJECT_STIFFENER Then
            Set oProfilePart.object = oBoundingData.MemberPart
            Set oBoundingConnPort = oProfilePart.SubPort(boundingConnPortXid)
        ElseIf pHelper.ObjectType(oBoundingData.MemberPart) = SDOBJECT_MEMBER Then
            Set oBoundingConnPort = GetLateralSubPortBeforeTrim(oBoundingData.MemberPart, boundingConnPortXid)
            Set oBoundingConnPort = pHelper.GetEquivalentLastPort(oBoundingConnPort)
        End If
    End If
            
    If oBoundedConnPort Is Nothing Or oBoundingConnPort Is Nothing Then
        sMsg = "Could not identify ports for connection"
        GoTo ErrorHandler
    End If
    
    ' --------------------------
    ' Create physical connection
    ' --------------------------
    Dim oPhysicalConnection As New PhysicalConn
    sMsg = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oBoundedConnPort, _
                               oBoundingConnPort, _
                               sWeldSel, _
                               oAppConnection, _
                               ConnectionStandard
    
    If Not oPhysicalConnection Is Nothing Then
        Set pObject = oPhysicalConnection.object
    Else
        sMsg = "Creation of physical connection failed"
        GoTo ErrorHandler
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' ========================================================================================================
' Method: IsOutsideCutNeeded
' Description: Determine whether an outside cut is needed
' Arguments
'     oAssyConn     - The member assembly connection
'     isBottomCut   - Flag for whether a top or bottom cut is being considered
'     eCond1        - Bounding condition
'     sFaceQuestion - If there is only one face question, set it here
'                     If there are two face questions, use the one corresponding to the isBottomCut flag
'                     Otherwise use vbNullString
'     sEdgeQuestion - If there is only one edge question, set it here
'                     If there are two edge questions, use the one corresponding to the isBottomCut flag
'                     Otherwise use vbNullString
' ========================================================================================================
'eBounded_Edge has all the possible edges on the bounded member that are used
'by the measurement symbol.

Public Function IsOutsideCutNeeded(oAssyConn As IJAppConnection, _
                                   eCutLocation As eEndCutLocation, _
                                   isBottomCut As Boolean, _
                                   eCond As eBoundingCondition, _
                                   sFaceQuestion As String, _
                                   sEdgeQuestion As String) As Boolean
    
    Const METHOD = "MbrACDefCM::IsOutsideCutNeeded"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    IsOutsideCutNeeded = False
    
    Dim bIsBottomFlange As Boolean
    bIsBottomFlange = False
    
    If eCutLocation = BottomFlangeEndCut Then
        bIsBottomFlange = True
    End If
    
    ' ---------------------------------
    ' Get the Bounded and Bounding Port
    ' ---------------------------------
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    InitMemberConnectionData oAssyConn, oBoundedData, oBoundingData, lStatus, sMsg

    ' ------------------------------------------------
    ' Outside web cut not needed if not web-penetrated
    ' ------------------------------------------------
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
    
    If Not bPenetratesWeb And eCutLocation = WebEndCut Then
        Exit Function
    End If
    
    ' -----------------------------------------------------------------------------
    ' Top flange cut not needed if not flange-penetrated or has no top flange
    ' Bottom flange cut not needed if not flange-penetrated or has no bottom flange
    ' -----------------------------------------------------------------------------
    If eCutLocation = TopFlangeEndCut And (bPenetratesWeb Or Not HasTopFlange(oBoundedData.MemberPart)) Then
        Exit Function
    ElseIf eCutLocation = BottomFlangeEndCut And (bPenetratesWeb Or Not HasBottomFlange(oBoundedData.MemberPart)) Then
        Exit Function
    End If
    
    ' ----------------------------------------------------------------------------------------
    ' If both bounded top and bottom are outside the bounding object, an outside cut is needed
    ' ----------------------------------------------------------------------------------------
    ' Need to add consideration of tube members here !!!!!!!!!!!
    If eCond = eBoundingCondition.OutsideAndOutside Then
        IsOutsideCutNeeded = True
        Exit Function
    End If
    
    ' --------------------------------------------------
    ' If there is material above, a top cut is needed
    ' If there is material below, a bottom cut is needed
    ' --------------------------------------------------
    Dim eTopBtmFlag As eCutFlag
    If isBottomCut Then
        eTopBtmFlag = eCutFlag.BottomCut
    Else
        eTopBtmFlag = eCutFlag.TopCut
    End If
    
    If BoundedHasOutsideMaterial(oAssyConn, eTopBtmFlag) Then
        IsOutsideCutNeeded = True
        Exit Function
    End If
    
    ' -----------------------------------------------------------------------------
    ' Get the top and bottom shape, and choose which to process based on input flag
    ' -----------------------------------------------------------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String
    Dim sTopAnswerCol As String
    Dim sTopShape As String
    
    GetMemberACTopAndBottomShape oAssyConn, sBottomAnswerCol, sBottomShape, sTopAnswerCol, sTopShape
     
    Dim sShapeCol As String
    Dim sShape As String
    
    If isBottomCut Then
        sShapeCol = sBottomAnswerCol
        sShape = sBottomShape
    Else
        sShapeCol = sTopAnswerCol
        sShape = sTopShape
    End If
            
    ' ------------------------------
    ' If based on ShapeAtFace answer
    ' ------------------------------
    If sShapeCol = gsShapeAtFaceCol Then
        
        Select Case LCase(sShape)
            ' --------------------------------------------------------
            ' Copes and snipes cut to the outside, and thus need a cut
            ' --------------------------------------------------------
            Case LCase(gsCope), LCase(gsSnipe)
    
                IsOutsideCutNeeded = True
                Exit Function
                
            ' --------------------------------------------------
            ' If the shape intends to cut between web and flange
            ' --------------------------------------------------
            Case LCase(gsInside)
            
                Dim bTFL As Boolean
                Dim bBFL As Boolean
                Dim bTFR As Boolean
                Dim bBFR As Boolean
    
                CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
            
                ' --------------------------------------------------------------------------
                ' If web-penetrated, there must be a flange at the cut location (top/bottom)
                ' --------------------------------------------------------------------------
                If bPenetratesWeb Then
                    If isBottomCut Then
                        If Not bBFL And Not bBFR Then
                            IsOutsideCutNeeded = True
                            Exit Function
                        End If
                    Else
                        If Not bTFL And Not bTFR Then
                            IsOutsideCutNeeded = True
                            Exit Function
                        End If
                    End If
                ' --------------------------------------------------------------------------------------------
                ' If flange-penetrated, the anwer only applies to the top, and only if there is no left flange
                ' --------------------------------------------------------------------------------------------
                Else
                    If Not isBottomCut And Not bTFL And Not bBFL Then
                        IsOutsideCutNeeded = True
                    End If
                End If
        End Select
    ' -----------------------
    ' If based on ShapeAtEdge
    ' -----------------------
    ElseIf sShapeCol = gsShapeAtEdgeCol Or sShapeCol = gsShapeAtEdgeOutsideCol Or sShapeCol = gsShapeAtEdgeOverlapCol Then
    
        ' -----------------------------------------------------
        ' See if the corner feature will take care of the shape
        ' -----------------------------------------------------
        ' We've established above if a cut is needed for outside material
        Dim isCornerAtBottomEdge As Boolean
        isCornerAtBottomEdge = isBottomCut Xor (eCond = Edge)
        
        Dim bIsCornerFeatureNeeded As Boolean
        CM_IsCornerNeededForShapeAtEdge oAssyConn, isCornerAtBottomEdge, bIsBottomFlange, bIsCornerFeatureNeeded
        
        If bIsCornerFeatureNeeded Then
            Exit Function
        End If
        
        ' ---------------------------------------------------------------------------------------------------------
        ' If not needed for outside material and not resolved by corner feature (see above), then base it on answer
        ' ---------------------------------------------------------------------------------------------------------
        ' Check if minimum overlap met
        Dim dInsideOverlap As Double
        Dim dOutsideOverlap As Double
        Dim bIsEdgeToEdge As Boolean
    
        GetEdgeOverlapAndClearance oAssyConn, isCornerAtBottomEdge, bIsBottomFlange, dInsideOverlap, dOutsideOverlap, , , , bIsEdgeToEdge
        
        Dim dOverlap As Double
        
        dOverlap = dInsideOverlap
        
        If eCond = Edge Then
            dOverlap = dOutsideOverlap
        End If
        
        Select Case LCase(sShape)
            Case LCase(gsNone)
                Exit Function
            
            ' -------------------------------------------------------------------------------------------------------------
            ' The following cases need a web cut if not resolved by a corner feature, but only if not met by a bounded edge
            ' -------------------------------------------------------------------------------------------------------------
            Case LCase(gsCornerToFlange)
                
                If Not bIsEdgeToEdge Then
                    IsOutsideCutNeeded = True
                End If
                
            ' ---------------------------------------------------------------
            ' The following cases need a cut only if a minimum overlap is met
            ' ---------------------------------------------------------------
            Case LCase(gsEdgeToFlange), _
                 LCase(gsEdgeToOutside)
                 
                If sShape = gsEdgeToFlange Or sShape = gsEdgeToOutside Then
                    If dOverlap >= GetMinOverlapToEdge() Then
                        IsOutsideCutNeeded = True
                    End If
                End If
                
            ' --------------------------------------
            ' All cases not handled above need a cut
            ' --------------------------------------
            Case Else
                IsOutsideCutNeeded = True
                Exit Function
        End Select
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function


' isBottom is for future use
Public Function IsOutsideWebCutNeededForChildBrace(oAssyConn As IJAppConnection, isBottom As Boolean, sBraceQuestion As String) As Boolean
  
    Const METHOD = "MbrACDefCM::IsOutsideWebCutNeededForChildBrace"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    IsOutsideWebCutNeededForChildBrace = False
    
    ' ---------------------------------
    ' Get the Bounded and Bounding Port
    ' ---------------------------------
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    InitMemberConnectionData oAssyConn, oBoundedData, oBoundingData, lStatus, sMsg

    ' -------------------------------------
    ' Determine if web or flange penetrated
    ' -------------------------------------
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
    
    ' -----------------
    ' If web penetrated
    ' -----------------
    If Not bPenetratesWeb Then
        Dim sBraceAnswer As String
        GetSelectorAnswer oAssyConn, sBraceQuestion, sBraceAnswer
        If sBraceAnswer = "InsetMember" Then
            IsOutsideWebCutNeededForChildBrace = True
        End If
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*********************************************************************************************
' Method      : CM_IsAxisEdgeCornerRequested
' Description :
'
'*********************************************************************************************
Public Sub CM_IsCornerNeededForShapeAtEdge(oConnection As IJAppConnection, bIsBottomEdge As Boolean, IsBottomFlange As Boolean, bIsNeeded As Boolean)
    
    Const METHOD = "MbrACDefCM::CM_IsCornerNeededForShapeAtEdge"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    'Get values from a common method and return results
    IsCornerFeatureNeededForShapeAtEdge oConnection, bIsBottomEdge, IsBottomFlange, bIsNeeded
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : GetGlobalPortAfterChamfer
' Description :
'   Gets the global port after the chamfer if it exists.  If there is no chamfer, it returns the
'   original port.
'*********************************************************************************************
Public Function GetGlobalPortAfterChamfer(oPort As IJPort) As IJPort
    
    Const METHOD = "GetGlobalPortAfterChamfer"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    'Get the Global Port After the Chamfer if it Exists
    Dim oStructPort As IJStructPort
    Dim oStructConnectable As IJStructConnectable
    Dim oConnectablePorts As IJElements
    
    Set GetGlobalPortAfterChamfer = oPort
    Set oStructPort = oPort
    Set oStructConnectable = oPort.Connectable
    oStructConnectable.enumConnectablePortsByOperation oConnectablePorts, "ChamferCutAE.ChamferCutPlate.1", PortFace, False, True
    
    If oConnectablePorts Is Nothing Then
        'No Connectable Ports from The ChamferCutAE.ChamferCutPlate.1 - Return the Original Port
        Exit Function
    End If
    
    Dim nPortCount As Integer
    nPortCount = oConnectablePorts.Count

    Dim nOrigCtx As Integer
    nOrigCtx = oStructPort.ContextID And (CTX_BASE Or CTX_OFFSET)
    
    Dim nPortIndex As Integer
    For nPortIndex = 1 To nPortCount
        Set oStructPort = oConnectablePorts.Item(nPortIndex)

        If oStructPort.OperationID = 0 And oStructPort.OperatorID = 0 And (oStructPort.ContextID And nOrigCtx) Then
            Set GetGlobalPortAfterChamfer = oStructPort
            Exit For
        End If
    Next nPortIndex
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' Method : GenericWebCutInput
'  This method is a copy of construct method of webcut

Private Function GenericWebCutInput(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object) As IJPort
    Const METHOD = "MbrACDefCM.GenericWebCutInput"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' ------------------------------------------
    ' Get the assembly connection and references
    ' ------------------------------------------
    sMsg = "Creating Web Cut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBoundingObjectColl As IJElements
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim i As Long
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oBoundingObj As Object
    Dim oPort As IJPort
    Dim lStatus As Long
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedPort As IJPort
    Dim LongBoxNeeded As Boolean
    
    Set oAppConnection = pMemberDescription.CAO
    
    Dim eAssemblyConnectionType As eACType
    eAssemblyConnectionType = GetMbrAssemblyConnectionType(oAppConnection)
    
    Dim lDispId As Long
    lDispId = pMemberDescription.dispid
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If eAssemblyConnectionType = ACType_Mbr_Generic Then
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        ' -------------------------------------------------------------
        ' Get all the bounding objects from the ports related to the AC
        ' -------------------------------------------------------------
        
        Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        ' ---------------------------------
        ' Get the Bounding  Port
        ' ---------------------------------
          
        Set oSmartOcc = pMemberDescription.CAO
        Set oSmartItem = oSmartOcc.SmartItemObject
        
        If oSmartItem.Name = "Generic_2B_01" Then
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
            Set oBoundingObj = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
            Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObj, oEditJDArgument).Item(1)
        ElseIf oSmartItem.Name = "Generic_LongBoxCuts" Then
            LongBoxNeeded = True
        Else
            Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
        End If
        
    Set oSmartItem = Nothing
    Set oSmartOcc = Nothing
   
    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
        
        Set oConnAttrbs = New SDConnectionUtils
        If LessThan(oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count, 1, 0) Then
         Exit Function
        End If
        
        Set oBoundingObjectColl = New JObjectCollection
        
        For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
            oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
        Next
        
        Set oAssyConn = New AssemblyConn
        Set oAssyConn.object = pMemberDescription.CAO
        
        'Adding the bounding port of the assembly connection to AuxiliaryPorts collection
        oBoundingObjectColl.Add oAssyConn.Port2
        Set GenericWebCutInput = oBoundingObjectColl.Item(lDispId)
    Else
        sMsg = "Not a vlaid type of assembly connection"
        GoTo ErrorHandler
    End If
    
    ' --------------------------------------------
    ' Throw an error if the DispID is out of range
    ' --------------------------------------------
        If LongBoxNeeded Then
            '
        ElseIf lDispId > oBoundingObjectColl.Count Then
            sMsg = "Creating Web Cut ...lDispId(" & Trim(Str(lDispId)) & ") > Object Count(" & Trim(Str(oBoundingObjectColl.Count)) & ")"
            GoTo ErrorHandler
        End If
        
    ' ---------------------------------
    ' Get the Bounded  Port
    ' ---------------------------------
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' ---------------------
    ' Get edge mapping data
    ' ---------------------
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Dim oSectionAlias As Long
    Dim bPenetratesWeb As Boolean

    Set oEdgeMapColl = New Collection
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    If LongBoxNeeded Then
            '
    Else
    bIsBoundedTube = IsTubularMember(oBoundedData.MemberPart)
    If TypeOf GenericWebCutInput.Connectable Is ISPSMemberPartPrismatic Then
        bIsBoundingTube = IsTubularMember(GenericWebCutInput.Connectable)
    End If
    
    Dim bIsTubeByTube As Boolean
    bIsTubeByTube = False
    
    If TypeOf GenericWebCutInput.Connectable Is ISPSMemberPartPrismatic Then
        If bIsBoundedTube And bIsBoundingTube Then
            bIsTubeByTube = True
        Else
            Set oEdgeMapColl = GetEdgeMap(oAppConnection, GenericWebCutInput, oBoundedPort, oSectionAlias, bPenetratesWeb)
        End If
    ElseIf TypeOf GenericWebCutInput.Connectable Is IJProfile Then
        Set oEdgeMapColl = GetEdgeMap(oAppConnection, GenericWebCutInput, oBoundedPort, oSectionAlias, bPenetratesWeb)
    End If
    
    Dim oBndingStiffObj As Object
    Dim bIsMbrBounidng As Boolean
    Dim bIsStiffBounding As Boolean
    
    If eAssemblyConnectionType = ACType_Stiff_Generic Then
      Set oBndingStiffObj = oBoundingObjectColl.Item(lDispId)
      If TypeOf oBndingStiffObj.Connectable Is IJProfile Then
        bIsStiffBounding = True
      ElseIf TypeOf oBndingStiffObj.Connectable Is ISPSMemberPartCommon Then
        bIsMbrBounidng = True
      End If
    End If
    End If
    ' ----------------------------
    ' If this is the first item...
    ' ----------------------------
    Dim oSDO_BoundingMember As New StructDetailObjects.MemberPart
    Dim oSDO_BoundingProfile As New StructDetailObjects.ProfilePart
    Dim eWL As ConnectedEdgeInfo
    Dim eWR As ConnectedEdgeInfo
    Dim eTFL As ConnectedEdgeInfo
    Dim eTFR As ConnectedEdgeInfo
    
    If lDispId = 1 Then
        ' ---------------------------------------------------------------------------
        ' Loop through the bounding objects and see if the web intersects any of them
        ' ---------------------------------------------------------------------------
        Dim oWebleftPort As IJPort
        Dim oWebRightPort As IJPort
        
        Dim bIsWebLeftIntersecting As Boolean
        Dim bIsWebRightIntersecting As Boolean
    
        Dim lIndex As Long
        Dim bWebIntersecting As Boolean
        bWebIntersecting = False
        
        For lIndex = 1 To oBoundingObjectColl.Count
            If bIsBoundedTube Then
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
            Else
                Set oWebleftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
            End If
            
            If Not oWebleftPort Is Nothing Then
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
            End If
            If Not oWebRightPort Is Nothing Then
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
            End If
    
            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                bWebIntersecting = True
                Exit For

            ' If the web didn't intersect any object, check if it does when extended
            Else
                If Not oWebleftPort Is Nothing Then
                    bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebleftPort, oBoundingObjectColl.Item(lIndex))
                End If
                If Not oWebRightPort Is Nothing Then
                    bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
                End If
            
                If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                    bWebIntersecting = True
                    Exit For
                End If
            End If
        Next lIndex
        
        ' ---------------------------------------------------------------------------------------------------------
        ' Create a web cut to the back-side port of the first bounding object if the web did not intersect anything
        ' ---------------------------------------------------------------------------------------------------------
        If bWebIntersecting = False Then
            If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Then
                If bIsBoundingTube Then
                    Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                Else
                    Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                    Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                End If
            ElseIf eAssemblyConnectionType = ACType_Mbr_Generic Then
                Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
            End If
            
            If eAssemblyConnectionType = ACType_Stiff_Generic Then
                If bIsMbrBounidng Then
                        If bIsBoundingTube Then
                            Set GenericWebCutInput = oBndingStiffObj
                        Else
                            Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                            Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                        End If
                ElseIf bIsStiffBounding Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(JXSEC_WEB_LEFT)))
                Else
                        Set GenericWebCutInput = oBndingStiffObj
                End If
            End If
            
        ' ---------------------------------------------
        ' If the web did intersect a bounding object...
        ' ---------------------------------------------
        Else
            ' ----------------------------------
            ' If the bounding object is a member
            ' ----------------------------------
            If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Or _
                bIsMbrBounidng Then
                
                If bIsTubeByTube Or bIsBoundingTube Then
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set GenericWebCutInput = oBndingStiffObj
                    End If
                Else
                    ' ---------------------------
                    ' If the flange is penetrated
                    ' ---------------------------
                    If Not bPenetratesWeb Then
                        ' -------------------------------------------------------------------
                        ' Get the port on the Bounding Member that the Bounded Web Intersects
                        ' -------------------------------------------------------------------
                        GetConnectedEdgeInfo oAppConnection, oBoundedPort, GenericWebCutInput, eWL, eWR, eTFL, eTFR
                                                                                    
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                        End If
                        
                        ' ---------------------------------------------------------------------------
                        ' If the left web face is outside, use the information for the right web face
                        ' ---------------------------------------------------------------------------
                        If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                            Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                        Else
                            Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                        End If
                    ' ----------------------------------------------------------------------------
                    ' If the web is penetrated, use the first selected port on the bounding member
                    ' ----------------------------------------------------------------------------
                    Else
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set GenericWebCutInput = oBndingStiffObj.Connectable
                        End If
                    End If
                End If
            
            ElseIf TypeOf oBoundingObjectColl.Item(lDispId) Is IJProfile Or _
                    bIsStiffBounding Then
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Profile that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, GenericWebCutInput, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingProfile.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                    End If
                    
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                    ' ----------------------------------------------------------------------------
                    ' If the web is penetrated, use the first selected port on the bounding member
                    ' ----------------------------------------------------------------------------
                    Else
                        If eAssemblyConnectionType = ACType_Mbr_Generic Then
                            Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                        ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                            Set GenericWebCutInput = oBndingStiffObj
                        End If
                End If
            Else
                ' ---------------------------------------------------------------------------------------------
                ' If the bounding object is not a member, assume it is a plate, and use the first selected port
                ' ---------------------------------------------------------------------------------------------
                If eAssemblyConnectionType = ACType_Mbr_Generic Then
                    If LongBoxNeeded Then
                        'The Bounding Port is set at the bottom of this method.
                    Else
                        'Get the Global Port After the Chamfer if it Exists
                        Set GenericWebCutInput = GetGlobalPortAfterChamfer(GenericWebCutInput)
                    End If
                ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                    Set GenericWebCutInput = oBndingStiffObj
                    'Get the Global Port After the Chamfer if it Exists
                    Set GenericWebCutInput = GetGlobalPortAfterChamfer(GenericWebCutInput)
                End If
            End If
        End If
    ' ---------------------------
    ' If the second or later item
    ' ---------------------------
    ElseIf LongBoxNeeded Then
            '
    Else
        ' ----------------------------------
        ' If the bounding object is a member
        ' ----------------------------------
        If TypeOf oBoundingObjectColl.Item(lDispId) Is ISPSMemberPartCommon Or _
            bIsMbrBounidng Then
            If bIsTubeByTube Or bIsBoundingTube Then
                If eAssemblyConnectionType = ACType_Mbr_Generic Then
                    Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                    Set GenericWebCutInput = oBndingStiffObj
                End If
            Else
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Member that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, GenericWebCutInput, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingMember.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingMember.object = oBndingStiffObj.Connectable
                    End If
                    
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingMember.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                ' ----------------------------------------------------------------------------
                ' If the web is penetrated, use the first selected port on the bounding member
                ' ----------------------------------------------------------------------------
                Else
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set GenericWebCutInput = oBndingStiffObj
                    End If
                End If
            End If
        ElseIf TypeOf oBoundingObjectColl.Item(lDispId) Is IJProfile Or _
               bIsStiffBounding Then
                ' ---------------------------
                ' If the flange is penetrated
                ' ---------------------------
                If Not bPenetratesWeb Then
                    ' -------------------------------------------------------------------
                    ' Get the port on the Bounding Profile that the Bounded Web Intersects
                    ' -------------------------------------------------------------------
                    GetConnectedEdgeInfo oAppConnection, oBoundedPort, GenericWebCutInput, eWL, eWR, eTFL, eTFR
                    
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set oSDO_BoundingProfile.object = oBoundingObjectColl.Item(lDispId)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set oSDO_BoundingProfile.object = oBndingStiffObj.Connectable
                    End If
                    ' ---------------------------------------------------------------------------
                    ' If the left web face is outside, use the information for the right web face
                    ' ---------------------------------------------------------------------------
                    If eWL.IntersectingEdge = eBounding_Edge.Above Or eWL.IntersectingEdge = eBounding_Edge.Below Then
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWR.IntersectingEdge)))
                    Else
                        Set GenericWebCutInput = GetLateralSubPortBeforeTrim(oSDO_BoundingProfile.object, oEdgeMapColl.Item(CStr(eWL.IntersectingEdge)))
                    End If
                ' ----------------------------------------------------------------------------
                ' If the web is penetrated, use the first selected port on the bounding member
                ' ----------------------------------------------------------------------------
                Else
                    If eAssemblyConnectionType = ACType_Mbr_Generic Then
                        Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                    ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                        Set GenericWebCutInput = oBndingStiffObj
                    End If
                End If
        Else
            ' ---------------------------------------------------------------------------------------------
            ' If the bounding object is not a member, assume it is a plate, and use the first selected port
            ' ---------------------------------------------------------------------------------------------
            If eAssemblyConnectionType = ACType_Mbr_Generic Then
                Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObjectColl.Item(lDispId), oEditJDArgument).Item(1)
                'Get the Global Port After the Chamfer if it Exists
                Set GenericWebCutInput = GetGlobalPortAfterChamfer(GenericWebCutInput)
            ElseIf eAssemblyConnectionType = ACType_Stiff_Generic Then
                Set GenericWebCutInput = oBndingStiffObj
                'Get the Global Port After the Chamfer if it Exists
                Set GenericWebCutInput = GetGlobalPortAfterChamfer(GenericWebCutInput)
            End If
        End If
    End If

    ' ---------------------
    ' Set the parent system
    ' ---------------------
    sMsg = "Retreiving Parent System for WebCut"
    
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' ------------------
    ' Create the Web Cut
    ' ------------------
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    
    Dim sEndCutSel As String
    Dim sACSelection As String
    
    Dim oACSelection As Object
    Set oACSelection = pMemberDescription.CAO
    
    If TypeOf oACSelection Is IJSmartOccurrence Then
        Set oSmartOcc = oACSelection
        Set oSmartItem = oSmartOcc.SmartItemObject
        sACSelection = oSmartItem.Name
    End If
    
    Dim oHelper As New StructDetailObjects.Helper

    If sACSelection = "Generic_LongBox" Then
        sEndCutSel = "MbrLongBoxEndSel"

        Set oBoundingObj = oBoundingObjectColl.Item(1)
        Set GenericWebCutInput = oBoundingObj.AxisPort(SPSMemberAxisAlong)
    ElseIf LongBoxNeeded Then
        Dim iIndex As Long
        'For member part
        If lDispId = 1 Then
          sEndCutSel = "MbrLongBoxEndSel" ' This gives Center web cut
            For iIndex = 1 To oBoundingObjectColl.Count
              If TypeOf oBoundingObjectColl.Item(iIndex) Is ISPSMemberPartPrismatic Then
                Set oBoundingObj = oBoundingObjectColl.Item(iIndex)
                Set GenericWebCutInput = oBoundingObj.AxisPort(SPSMemberAxisAlong)
              End If
            Next iIndex
         'For plate part
        Else
          sEndCutSel = "RootGenWebSel"
            If lDispId = 2 Then              'This gives top web cut
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
            ElseIf lDispId = 3 Then          'This gives Bottom web cut
              Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)
            End If
              Set oBoundingObj = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
              Set GenericWebCutInput = GetPortsFromBoundingObject(oBoundingObj, oEditJDArgument).Item(1)
       End If
    Else
        sEndCutSel = "RootGenWebSel"
    End If
    

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

'******************************************************************************************
' Method:
' CMNeedtoComputeWebCut
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************

Public Sub CMNeedtoComputeWebCut(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "EndCutDefCM::CMNeedtoComputeWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the FlangeCut "
    
    If TypeOf pObject Is IJStructFeature Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        If Not oSmartItem Is Nothing Then
            Dim pMemberDescription As IJDMemberDescription
            Set pMemberDescription = GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions, pObject)
            Dim oEndCutObject As Object
            Set oEndCutObject = pMemberDescription.CAO
            
            Dim oSDO_WebCut As New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = pObject
            Dim oFeatureUtils As New GSCADSDCreateModifyUtilities.SDFeatureUtils
            Dim pResourceMgr As IUnknown
            Dim oNewBoundingPort As IJPort
            
            Set pResourceMgr = GetResourceManagerFromObject(pObject)
            ' GenericWebCutInput method is the refactored code of CM_CreateGenericWebCut to get the update port
            Set oNewBoundingPort = GenericWebCutInput(pMemberDescription, pResourceMgr, oSDO_WebCut.object)
            
            Dim oParentSystem As IJSystem
            Dim oDesignParent As IJDesignParent
            ' Need to get the IJSystem Interface
            sMsg = "Retreiving Parent System for WebCut"
            If TypeOf oEndCutObject Is IJDesignParent Then
                Set oDesignParent = oEndCutObject
                If TypeOf oDesignParent Is IJSystem Then
                    Set oParentSystem = oDesignParent
                End If
            End If
            Dim oOldBdgPort As IJPort
            Dim oNewBdgPort As IJPort
            
            Set oOldBdgPort = oSDO_WebCut.BoundingPort
            Set oNewBdgPort = oNewBoundingPort
            'Checking whether the new and old bounding ports are same or not
            If Not oOldBdgPort Is Nothing And Not oNewBdgPort Is Nothing Then
                If Not oOldBdgPort.Connectable Is oNewBdgPort.Connectable Then
                   oFeatureUtils.CreateWebCut pResourceMgr, oNewBoundingPort, oSDO_WebCut.BoundedPort, "", oParentSystem, oSDO_WebCut.object
                End If
            End If
        End If
    End If
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

