Attribute VB_Name = "MbrEndCutUtilities"
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-2015, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrEndCutUtilities.bas
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    30/Nov/2011 - svsmylav TR-205302: Removed commented code '.Subport'.
'    13/Oct/2014  -hgunturu TR-240217: Modified Set_FlangeCuttingBehavior() Method
'    18/Dec/2014  -hgunturu TR-265939: Changed the flange cutting behaviour for Miters cases
'    27/Dec/2015  -hgunturu DI-282754: Modified Set_FlangeCuttingBehavior() Method to improve performance
'*********************************************************************************************
Private Const MODULE = "StructDetail\Data\SmartOccurrence\SMMbrEndCut\MbrEndCutUtilities"
Public Const m_sProjectName As String = CUSTOMERID + "MbrEndCut"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\SmartOccurrence\" + m_sProjectName + "\"

'Constants with GUIDs
Public Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Public Const CA_FLANGECUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Public Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

'*************************************************************************
'Function: GetBoundingAndBounded
'
'Description:
'   Given any end cut (web/flange) Smart Occurent Object,
'    return the bounding and bounded ports.
'    Also, optionally return bounding and bounded objects
'
'Input
'   oSmartOccurence
'
'Return
'   oBoundingPort
'   oBoundedPort
'   Optionally oBoundingObject
'   Optionally oBoundedObject
'Exceptions
'
'***************************************************************************
Public Sub GetBoundingAndBounded(ByVal oSmartOccurence As Object, _
    oBoundingPort As IJPort, oBoundedPort As IJPort, _
    Optional oBoundingObject As Object, Optional oBoundedObject As Object, _
    Optional eEndCutType As eEndCutTypes = WebCut)
    
    Const METHOD = "MbrEndCutUtilities::GetBoundingAndBounded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    Dim oStructFeature As GSCADSDCreateModifyUtilities.IJStructFeature
    Set oStructFeature = oSmartOccurence
    
    Dim oBounded As Object
    Dim oBounding As Object
       
    If oStructFeature.get_StructFeatureType = SF_FlangeCut Then
          Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
          Set oSDO_FlangeCut.object = oSmartOccurence
          Set oBoundedPort = oSDO_FlangeCut.BoundedPort
          Set oBoundingPort = oSDO_FlangeCut.BoundingPort
          Set oBounded = oSDO_FlangeCut.Bounded
          Set oBounding = oSDO_FlangeCut.Bounding
          If oSDO_FlangeCut.IsTopFlange Then
            eEndCutType = eEndCutTypes.FlangeCutTop
          Else
            eEndCutType = eEndCutTypes.FlangeCutBottom
          End If
    Else
          Set oSDO_WebCut = New StructDetailObjects.WebCut
          Set oSDO_WebCut.object = oSmartOccurence
          Set oBoundedPort = oSDO_WebCut.BoundedPort
          Set oBoundingPort = oSDO_WebCut.BoundingPort
          Set oBounded = oSDO_WebCut.Bounded
          Set oBounding = oSDO_WebCut.Bounding
          eEndCutType = WebCut
    End If
    Set oBoundingObject = oBounding
    Set oBoundedObject = oBounded
    
    Set oSDO_FlangeCut = Nothing
    Set oSDO_WebCut = Nothing
    Set oStructFeature = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "MbrEndCutUtilities::GetBoundingAndBounded", sMsg).Number

End Sub

'*************************************************************************
'Function: GetAppropriateCutDepth
'
'Description:
'  This method is implemented to get the Correct cut depth needed for the
'     falnge cut so that it can trigger the cutting tool needed(along with web notch)
'
'  Implemented for the TR-181872\TR-189574
'
'***************************************************************************
Public Function GetAppropriateCutDepth(pPRL As IJDParameterLogic, bIsBottomFlange As Boolean) As Double
    
    Const METHOD = "MbrEndCutUtilities::GetAppropriateCutDepth"
    
    Dim oBoundedPort    As IJPort
    Dim oMember         As New StructDetailObjects.MemberPart
    Dim oWRPort         As IJPort
    Dim oTopOrBtmPort   As IJPort
    Dim sSectionType    As String
    Dim oWebRight       As IJDModelBody
    Dim oTopOrBtm       As IJDModelBody
    Dim dtFlange        As Double
    Dim oSectionAttrbs  As IJDAttributes
    Dim oMindist        As Double
    Dim oPoint1         As IJDPosition
    Dim oPOint2         As IJDPosition
    Dim sMsg            As String
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    
    
    GetAppropriateCutDepth = 0.1
    
    sMsg = "Getting the CUt depth for the S type members "
    
    Set oBoundedPort = pPRL.InputObject(INPUT_BOUNDED)
    
  If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        Set oMember.object = oBoundedPort.Connectable
        sSectionType = oMember.sectionType
        Select Case UCase(sSectionType)

        Case "HSSC", "HSSR", "PIPE", "CS", "RS"
          GetAppropriateCutDepth = EndCut_GetCutDepth(pPRL)
        Case Else
            Set oWRPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_WEB_RIGHT)
            If bIsBottomFlange Then
              Set oTopOrBtmPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_BOTTOM)
            Else
              Set oTopOrBtmPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_TOP)
            End If
            If TypeOf oWRPort.Geometry Is IJDModelBody And TypeOf oTopOrBtmPort.Geometry Is IJDModelBody Then
              Set oWebRight = oWRPort.Geometry
              Set oTopOrBtm = oTopOrBtmPort.Geometry
                oWebRight.GetMinimumDistance oTopOrBtm, oPoint1, oPOint2, oMindist
                GetAppropriateCutDepth = oMindist
            Else
              GetAppropriateCutDepth = EndCut_GetCutDepth(pPRL)
            End If
        End Select
    Else
        GetAppropriateCutDepth = EndCut_GetCutDepth(pPRL)
    End If

    
    Dim vValue As Variant
    
    Dim oStructFeature As IJStructFeature
    Set oStructFeature = pPRL.SmartOccurrence
    
    If oStructFeature.get_StructFeatureType = SF_FlangeCut Then
        If Has_Attribute(pPRL.SmartOccurrence, "CuttingBehavior") Then
            GetAppropriateCutDepth = 2 * GetAppropriateCutDepth
        End If
    End If
          
    Set oBoundedPort = Nothing
    Set oMember = Nothing
    Set oWRPort = Nothing
    Set oTopOrBtmPort = Nothing
    Set oWebRight = Nothing
    Set oTopOrBtm = Nothing
    Set oSectionAttrbs = Nothing
    Set oPoint1 = Nothing
    Set oPOint2 = Nothing
    Set oMemberPartPrismatic = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
    End Function
    
'*************************************************************************
' Function:
' GetNumberOfWebCutsOnAC
'
' Abstract
'
'***************************************************************************
Public Function GetNumberOfWebCutsOnAC(oSO As IJSmartOccurrence) As Long

    GetNumberOfWebCutsOnAC = 0

    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSO, sACItemName, oACObject
    
    Dim oMemberCut As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oACObject
    
    Dim nWebCuts As Long
    Dim oStructFeature As IJStructFeature
    Dim i As Long

    nWebCuts = 0
    For i = 1 To oMemberObjects.Count
        If Not oMemberObjects.ItemByDispid(i) Is Nothing Then
            If TypeOf oMemberObjects.ItemByDispid(i) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(i)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    GetNumberOfWebCutsOnAC = GetNumberOfWebCutsOnAC + 1
                End If
            End If
        End If
    Next i

End Function

'*************************************************************************
' Function:
' GetFirstPenetrated
'
' Abstract
'
'***************************************************************************
Public Function GetFirstPenetrated(bBottomFlange As Boolean, oSO As IJSmartOccurrence) As Object
    Const METHOD = "MbrEndCutUtilities::GetFirstPenetrated"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Set GetFirstPenetrated = Nothing

    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Set oSDOWebCut.object = oSO

    ' ----------
    ' Get the AC
    ' ----------
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSO, sACItemName, oACObject
    
    Dim oMemberCut As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oACObject
    
    ' ------------------------------------
    ' Get the end port geometry and normal
    ' ------------------------------------
    Dim oBoundedPort As Object
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    Dim oBoundingObject As Object
    Dim eEndCutType As eEndCutTypes
    
    EndCut_InputData oSO, oBoundedPort, oBoundedPart, oBoundingObject, oBoundingPart, eEndCutType
    
    Dim oEndPort As IJPort
    Set oEndPort = Member_GetSolidPort(oBoundedPort, True)
    
    Dim oEndSurface As IJSurfaceBody
    Set oEndSurface = oEndPort.Geometry
    
    Dim oSGOModelUtil As IJSGOModelBodyUtilities
    Dim oEndPortNorm As IJDVector
    Dim oPointOnEnd As IJDPosition
    Dim dist As Double
    
    Set oSGOModelUtil = New SGOModelBodyUtilities
    
    oSGOModelUtil.GetClosestPointOnBody oEndSurface, oSDOWebCut.BoundedLocation, oPointOnEnd, dist
    
    oEndSurface.GetNormalFromPosition oPointOnEnd, oEndPortNorm
    
    ' ----------------------------
    ' Get the flange port geometry
    ' ----------------------------
    Dim oPort As IJPort
    If bBottomFlange Then
        Set oPort = GetLateralSubPortBeforeTrim(oSDOWebCut.Bounded, JXSEC_BOTTOM)
    Else
        Set oPort = GetLateralSubPortBeforeTrim(oSDOWebCut.Bounded, JXSEC_TOP)
    End If

    Dim oExtendedPort As IJSurfaceBody
    Set oExtendedPort = GetExtendedPort(oPort)
    
    ' ------------------------------------------------------------------
    ' Loop through all web cuts on AC (one per bounding object expected)
    ' ------------------------------------------------------------------
    Dim oIntersect As IJDTopologyIntersect
    Dim oCommonGeom As Object
    Dim oPointOnCommon As IJDPosition
    Dim maxDist As Double
    
    Set oIntersect = New DGeomOpsIntersect
    maxDist = -1000000#
    
    Dim oStructFeature As IJStructFeature
    Dim i As Long

    Dim oDirVector As IJDVector

    For i = 1 To oMemberObjects.Count
        If Not oMemberObjects.ItemByDispid(i) Is Nothing Then
            If TypeOf oMemberObjects.ItemByDispid(i) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(i)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    Set oSDOWebCut.object = oStructFeature

                    ' -----------------------------------
                    ' Intersect flange with bounding port
                    ' -----------------------------------
                    Dim oBoundingPort As IJPort
                    Dim oBoundingPortGeom As Object
                    ' -----------------------------------
                    ' Check if it is a profile or member if it is use the global lateral port
                    ' If it is a plate use the web cut port
                    ' -----------------------------------
                    If TypeOf oSDOWebCut.Bounding Is IJProfile Then
                        'Profile
                        Dim oProfilePart As New StructDetailObjects.ProfilePart
                        Set oProfilePart.object = oSDOWebCut.Bounding
                        Set oBoundingPort = oProfilePart.BasePortBeforeTrim(BPT_Lateral)
                        Set oBoundingPortGeom = oBoundingPort.Geometry
                    ElseIf TypeOf oSDOWebCut.Bounding Is ISPSMemberPartPrismatic Then
                        'Member
                        Dim oMemberPart As New StructDetailObjects.MemberPart
                        Set oMemberPart.object = oSDOWebCut.Bounding
                        Set oBoundingPort = oMemberPart.BasePortBeforeTrim(BPT_Lateral)
                        Set oBoundingPortGeom = oBoundingPort.Geometry
                    Else
                        'Plate
                        Set oBoundingPort = oSDOWebCut.BoundingPort
                        Set oBoundingPortGeom = oBoundingPort.Geometry
                    End If
                    
                    
                    If TypeOf oBoundingPortGeom Is IJSurfaceBody Then
                        On Error Resume Next
                        Set oCommonGeom = Nothing
                        oIntersect.PlaceIntersectionObject Nothing, oExtendedPort, oBoundingPortGeom, Nothing, oCommonGeom
                        On Error GoTo ErrorHandler
                    End If
                    
                    ' ----------------------------------------------
                    ' Get distance from intersection to the end port
                    ' ----------------------------------------------
                    dist = -1000000#
                    If Not oCommonGeom Is Nothing Then
                        If TypeOf oSDOWebCut.Bounding Is IJProfile Or TypeOf oSDOWebCut.Bounding Is ISPSMemberPartPrismatic Then
                            ' If the bounding object is a stiffener or member then loop through all the vertices on the intersecting
                            ' geometry to get the vertex that is the furthest from the bounded member End Surface
                            Dim oVertexColl As Collection
                            oSGOModelUtil.GetVertices oCommonGeom, oVertexColl
                            
                            Dim dVertexDist As Double
                            Dim dVertexMaxDist As Double
                            Dim oVertexPos As IJDPosition
                            Dim oClosestPoint As IJDPosition
                            Dim oVector As IJDVector
                            Dim j As Long
                            Dim oPointColl As IJElements
                            Set oPointColl = New JObjectCollection
                            
                            dVertexMaxDist = -1000000#
                            
                            For Each oVertexPos In oVertexColl
                                oSGOModelUtil.GetClosestPointOnBody oEndSurface, oVertexPos, oClosestPoint, dVertexDist
                                If GreaterThanOrEqualTo(dVertexDist, dVertexMaxDist) Then
                                    dVertexMaxDist = dVertexDist
                                    oPointColl.Add oVertexPos
                                End If
                            Next
                            'We are looping through all the vertices which are at equal distant from end surface of Bounded Member
                            'and we are getting the Point which is nearest towards the bounding member side.
                                For j = 1 To oPointColl.Count
                                        Set oVector = oPointColl.Item(j).Subtract(oPointOnEnd)
                                        If oVector.Dot(oEndPortNorm) < 0# Then
                                            Set oPointOnCommon = oPointColl.Item(j)
                                        End If
                                Next j
                            dist = dVertexMaxDist
                        Else
                        
                            Dim oVerticesColl As Collection
                            Dim oVerticesPosition As IJDPosition
                            Dim oVertexNormal As IJDVector
                            oSGOModelUtil.GetVertices oCommonGeom, oVerticesColl
                            
                            'If it is a plate, get the distance from the EndSurface to the intersecting geometry
                            oSGOModelUtil.GetClosestPointsBetweenTwoBodies oEndSurface, oCommonGeom, oPointOnEnd, oPointOnCommon, dist
                            
                            For Each oVerticesPosition In oVerticesColl
                                Set oVertexNormal = oVerticesPosition.Subtract(oPointOnEnd)
                                If oVertexNormal.Dot(oEndPortNorm) < 0# Then
                                    Set oPointOnCommon = oVerticesPosition
                                End If
                            Next
                        End If
                        
                        If dist > 0.00001 And Not oPointOnCommon Is Nothing Then
                            Set oDirVector = oPointOnCommon.Subtract(oPointOnEnd)
                            If oDirVector.Dot(oEndPortNorm) > 0# Then
                                dist = -dist
                            End If
                        End If
                      
                    End If
                    
                    ' -------------------------------
                    ' Save the boundary furthest away
                    ' -------------------------------
                    If dist > maxDist Then
                        maxDist = dist
                        Set GetFirstPenetrated = oSDOWebCut.Bounding
                    End If
                End If
            End If
        End If
    Next i
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

'*************************************************************************
' Function:
' GetNumberOfFlangeCutsOnAC
'
' Abstract
'
'***************************************************************************
Public Function GetNumberOfFlangeCutsOnAC(oSO As IJSmartOccurrence) As Long

    GetNumberOfFlangeCutsOnAC = 0

    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSO, sACItemName, oACObject
    
    Dim oMemberCut As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oACObject
    
    Dim nWebCuts As Long
    Dim oStructFeature As IJStructFeature
    Dim i As Long

    nWebCuts = 0
    For i = 1 To oMemberObjects.Count
        If Not oMemberObjects.ItemByDispid(i) Is Nothing Then
            If TypeOf oMemberObjects.ItemByDispid(i) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(i)
                If oStructFeature.get_StructFeatureType = SF_FlangeCut Then
                    GetNumberOfFlangeCutsOnAC = GetNumberOfFlangeCutsOnAC + 1
                End If
            End If
        End If
    Next i

End Function
Public Sub Set_FlangeCuttingBehavior(ByVal oFlangeCut As Object, bHasPC As Boolean)

  Const METHOD = "MbrEndCutUtilities::Set_FlangeCuttingBehavior"
  On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim oStructPort As IJStructPort
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oFlangeCut
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools
    Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
    Dim vOldBehaviorType As Variant
    
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As Object
    Dim oBoundingObect As Object
    Dim oBoundedObject As Object
    
    Set oBoundingObect = oSDO_FlangeCut.Bounding
    Set oBoundedObject = oSDO_FlangeCut.Bounded
    Set oBoundingPort = oSDO_FlangeCut.BoundingPort
    Set oBoundedPort = oSDO_FlangeCut.BoundedPort
    
       
    Get_AttributeValue oFlangeCut, "CuttingBehavior", vOldBehaviorType
    
    Dim vNewBehaviorType As Variant
        
    If Not bHasPC Then
        vNewBehaviorType = AvoidWeb
    Else
    
        'When Bounding is Member or Stiffener
        If TypeOf oBoundingObect Is IJStructProfilePart Then
    
            Dim strACSmartItem As String
            Dim oACObject As Object
            AssemblyConnection_SmartItemName oFlangeCut, strACSmartItem, oACObject
    
            'For Miter, Split cases
            Dim eGetACType As eACType
            eGetACType = GetMbrAssemblyConnectionType(oACObject)
            
            If eGetACType = ACType_Miter Or eGetACType = ACType_Split Then
                    
                vNewBehaviorType = ToFlangeInnerSurface
                
            'For Flange Penetrated Cases
            ElseIf Not IsWebPenetrated(oBoundingPort, oBoundedPort) Then
                vNewBehaviorType = ToFlangeInnerSurface
            
            ElseIf TypeOf oBoundingPort Is IJStructPort Or TypeOf oBoundingPort Is ISPSSplitAxisAlongPort And Not IsTubularMember(oBoundingObect) Then
                
                If TypeOf oBoundingPort Is IJStructPort Then
                    Set oStructPort = oBoundingPort
                Else
                    'Get the Flange Intesecting Port ID
                    Dim eBoundingEdge As eBounding_Edge
                    Dim eMappedPort As JXSEC_CODE
                    Dim bIsBottomFlange As Boolean
                    Dim sBottomFlange As String
                                        
                    GetSelectorAnswer oFlangeCut, "BottomFlange", sBottomFlange
                    
                    If sBottomFlange = gsYes Then
                        bIsBottomFlange = True
                    Else
                        bIsBottomFlange = False
                    End If
                    
                    GetNonPenetratedIntersectedEdge oFlangeCut, oBoundingPort, oBoundedPort, eBoundingEdge, eMappedPort, bIsBottomFlange
                    
                    'Mapped Edge will be UNKNOWN when Intersecting is Above, Below or None
                    If Not eMappedPort = JXSEC_UNKNOWN Then
                        Set oStructPort = GetLateralSubPortBeforeTrim(oBoundingObect, eMappedPort)
                    End If
                End If
                
                If Not oStructPort Is Nothing Then
                    If oStructPort.SectionID = JXSEC_TOP_FLANGE_LEFT Or oStructPort.SectionID = JXSEC_TOP_FLANGE_RIGHT Or _
                        oStructPort.SectionID = JXSEC_BOTTOM_FLANGE_LEFT Or oStructPort.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                        oStructPort.SectionID = JXSEC_TOP_FLANGE_TOP Or oStructPort.SectionID = JXSEC_BOTTOM_FLANGE_BOTTOM Then
        
                        vNewBehaviorType = ToBoundingEdge
                        
                    ElseIf (oStructPort.OperatorID = JXSEC_TOP_FLANGE_LEFT Or oStructPort.OperatorID = JXSEC_TOP_FLANGE_RIGHT Or _
                            oStructPort.OperatorID = JXSEC_BOTTOM_FLANGE_LEFT Or oStructPort.OperatorID = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                            oStructPort.OperatorID = JXSEC_TOP_FLANGE_TOP Or oStructPort.OperatorID = JXSEC_BOTTOM_FLANGE_BOTTOM) _
                            And oStructPort.SectionID = -1 Then
                        
                        vNewBehaviorType = ToBoundingEdge
                    Else
                        'If Bounding is Member WebEdge
                        If ((Not HasBottomFlange(oBoundingObect)) And oStructPort.OperatorID = JXSEC_BOTTOM) Or _
                            ((Not HasTopFlange(oBoundingObect)) And oStructPort.OperatorID = JXSEC_TOP) Then
                            
                            vNewBehaviorType = ToBoundingEdge
                        Else
                            vNewBehaviorType = ToFlangeInnerSurface
                        End If
                    End If
                Else
                    vNewBehaviorType = ToFlangeInnerSurface
                End If
            Else 'For Remaining cases
                vNewBehaviorType = ToFlangeInnerSurface
            End If

        'When Bounding is Plate Part
        ElseIf (TypeOf oBoundingObect Is IJPlate) Or (TypeOf oBoundingObect Is SPSWallPart) Or (TypeOf oBoundingObect Is SPSSlabEntity) Then
    
            Set oStructPort = oBoundingPort
    
            'Plate Lateral surface
            If oStructPort.ContextID And CTX_LATERAL Then
                vNewBehaviorType = ToBoundingEdge
            Else 'with Plate Base or Offset
                vNewBehaviorType = ToFlangeInnerSurface
            End If
        Else
            'Unknown Unhandled cases
            Exit Sub
        End If
        
    
    End If
    
    Set_AttributeValue oFlangeCut, "IJShpStrFlangeCut", "CuttingBehavior", vNewBehaviorType
    
 Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    

End Sub

