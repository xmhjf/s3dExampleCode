Attribute VB_Name = "CornerFeatureUtilities"
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-2016, Intergraph Corporation.  All rights reserved.
'
'  File        : CornerFeatureUtilities.bas
'
'  Description : Common Methods for Corner Features
'
'  Author      : Alligators
'
'  History     :
'    18/APR/2011 - Creation
'    21/Sep/2011 - mpulikol
'           DI-CP-200263  Improve performance by caching measurement symbol results
'    22/Sep/2011 - svsmylav TR-202526: edge id of CF is corrected to JXSEC_BOTTOM_FLANGE_RIGHT/JXSEC_TOP_FLANGE_RIGHT in If condition
'                  in GetProjBoundingFlangeThickness method (earlier check was using bottom flange top/top flange bottom: this won't work for CF on outside corner).
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
'   29 Nov 2014 - GHM       DI-259276- Updated GetFlangeThickness() method with validation checks
'   21 Apr 2015 - pkakula   TR-271022- Updated GetCornerFatureEdgeLengths ()
'   07 Jan 2016 -GHM        TR-279987- Updated GetCornerFatureEdgeLengths() method with proper corner feature input object
'   15/June/2016 -knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'*********************************************************************************************

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\MbrAxisEndCutCFSel.bas"
Public Const LINEAR_TOLERANCE = 0.0000001

Public Enum eEndCutCornerLocation

    EndCutCornerLocationUnknown
    BottomEdgeInside
    BottomEdgeOutside
    TopEdgeInside
    TopEdgeOutside
    FaceBottomInsideCorner
    FaceTopInsideCorner
    WebAndFlangeCommon ' Between web cut and flange cut

End Enum

'*********************************************************************************************
' Method      : GetGrandParentName
' Description : Given the Smart Occurence Object return the Owning grand parent Smart item Name
'
'*********************************************************************************************
Public Function GetGrandParentName(oSmartOccurrence As Object, Optional oGrandParent As Object) As String
    Const METHOD = "CornerFeatureUtilities::GetGrandParentName"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oSmartOccurrence, sParentItemName, oParentObj
    
    Dim sGPName As String
    Parent_SmartItemName oParentObj, sGPName, oGrandParent
    GetGrandParentName = sGPName
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*********************************************************************************************
' Method      : GetFlangeThickness
' Description : get flange thickness of bounded and bounding members
'
'*********************************************************************************************
Public Sub GetFlangeThickness(oSmartOccurence As Object, Optional dBoundedFlangeThickness As Double, Optional dBoundingFlangeThickness As Double)
    Const METHOD = "CornerFeatureUtilities::GetFlangeThickness"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    'Get bounding and bounded member objects
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort

    Dim oBoundingObject As Object
    Dim oBoundedObject As Object
    Dim oBoundedMbr As StructDetailObjects.MemberPart
    Dim oBoundedProfile As StructDetailObjects.ProfilePart
    Dim oBoundingMbr As StructDetailObjects.MemberPart
    Dim oBoundingProfile As StructDetailObjects.ProfilePart
    
    GetBoundingAndBoundedForCorner oSmartOccurence, _
        oBoundingPort, oBoundedPort, _
         oBoundingObject, oBoundedObject
    
    If Not oBoundedObject Is Nothing Then
        If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
            Set oBoundedMbr = New StructDetailObjects.MemberPart
            Set oBoundedMbr.object = oBoundedObject
            dBoundedFlangeThickness = oBoundedMbr.flangeThickness
        ElseIf TypeOf oBoundedObject Is IJProfile Then
            Set oBoundedProfile = New StructDetailObjects.ProfilePart
            Set oBoundedProfile.object = oBoundedObject
            dBoundedFlangeThickness = oBoundedProfile.flangeThickness
        End If
    End If
    
    If Not oBoundingObject Is Nothing Then
        If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
            Set oBoundingMbr = New StructDetailObjects.MemberPart
            Set oBoundingMbr.object = oBoundingObject
            dBoundingFlangeThickness = oBoundingMbr.flangeThickness
        ElseIf TypeOf oBoundingObject Is IJProfile Then
            Set oBoundingProfile = New StructDetailObjects.ProfilePart
            Set oBoundingProfile.object = oBoundingObject
            dBoundingFlangeThickness = oBoundingProfile.flangeThickness
        End If
    End If
'
'
'    'oBoundedObject and oBoundingObject will either be SDO.MemberPart or SDO.ProfilePart
'    'Both have a flangeThickness property.
'    dBoundedFlangeThickness = oBoundedObject.flangeThickness
'    dBoundingFlangeThickness = oBoundingObject.flangeThickness
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*************************************************************************
'Function: GetBoundingAndBoundedForCorner
'
'Description:
'   Given corner feature Smart Occurence Object on member,
'   return the bounding and bounded ports of the parent end cut.
'    Also, optionally return bounding and bounded objects
'
'Input
'   oCornerFeature
'
'Return
'   oBoundingPort
'   oBoundedPort
'   Optionally oBoundingObject
'   Optionally oBoundedObject
'Exceptions
'
'***************************************************************************
Public Sub GetBoundingAndBoundedForCorner(ByVal oCornerFeature As Object, _
                                          oBoundingPort As IJPort, _
                                          oBoundedPort As IJPort, _
                                          Optional oBoundingObject As Object, _
                                          Optional oBoundedObject As Object)
                                          
    Const METHOD = "CornerFeatureUtilities::GetBoundingAndBoundedForCorner"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim oFeatureAsChild As IJDesignChild
    Dim oCornerParent As Object
    
    If TypeOf oCornerFeature Is IJDesignChild Then
        Set oFeatureAsChild = oCornerFeature
        Set oCornerParent = oFeatureAsChild.GetParent
    End If
    
    Dim eFeatureType As StructFeatureTypes
    Dim oParentFeature As IJStructFeature
    
    eFeatureType = SF_CornerFeature
    If TypeOf oCornerParent Is IJStructFeature Then
        Set oParentFeature = oCornerParent
        eFeatureType = oParentFeature.get_StructFeatureType
    End If
        
    Dim oBoundedMember As Object
    Dim oBoundingMember As Object
    
    Dim oBoundedProfile As Object
    Dim oBoundingProfile As Object
    
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oParentFeature
            Set oBoundedPort = oWebCut.BoundedPort
            Set oBoundingPort = oWebCut.BoundingPort
            
            If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
                Set oBoundedMember = oWebCut.Bounded
            ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
                Set oBoundedProfile = oWebCut.Bounded
            End If
            
            If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
                Set oBoundingMember = oWebCut.Bounding
            ElseIf TypeOf oBoundingPort.Connectable Is IJProfile Then
                Set oBoundingProfile = oWebCut.Bounding
            End If
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oParentFeature
            Set oBoundedPort = oFlangeCut.BoundedPort
            Set oBoundingPort = oFlangeCut.BoundingPort
            
            If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
                Set oBoundedMember = oFlangeCut.Bounded
            ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
                Set oBoundedProfile.object = oFlangeCut.Bounded
            End If

            If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
                Set oBoundingMember = oFlangeCut.Bounding
            ElseIf TypeOf oBoundingPort.Connectable Is IJProfile Then
                Set oBoundingProfile = oFlangeCut.Bounding
            End If
        Case Else
            sMsg = "Wrong feature type"
            GoTo ErrorHandler
    End Select
    
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
        Set oBoundedObject = oBoundedMember
    ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
        Set oBoundedObject = oBoundedProfile
    End If
    
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
        Set oBoundingObject = oBoundingMember
    ElseIf TypeOf oBoundingPort.Connectable Is IJProfile Then
        Set oBoundingObject = oBoundingProfile
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' ------------------------------------------------------------------------------
' GetProjBoundingFlangeThickness: given corner feature (at the edge of the bounding)
' and two edge ports, computes the bounding flange thickness measured in a plane
' that contains the web of the bounded.
' ------------------------------------------------------------------------------
Public Function GetProjBoundingFlangeThickness(oCornerFeature As IJStructFeature, _
                                                     oPort1 As IJStructPort, _
                                                     oPort2 As IJStructPort) As Double
    Const sMETHOD = "GetProjBoundingFlangeThickness"
        
    ' ----------------------------------
    ' Get the edges used for the feature
    ' ----------------------------------
    Dim xID1 As JXSEC_CODE
    Dim xID2 As JXSEC_CODE
    
    xID1 = oPort1.SectionID
    xID2 = oPort2.SectionID
    
    Dim oFeatureAsChild As IJDesignChild
    Dim oCornerParent As Object
    
    If TypeOf oCornerFeature Is IJDesignChild Then
        Set oFeatureAsChild = oCornerFeature
        Set oCornerParent = oFeatureAsChild.GetParent
    End If
    
    Dim oACObject As IJAppConnection
    AssemblyConnection_SmartItemName oCornerFeature, , oACObject

    Dim eFeatureType As StructFeatureTypes
    Dim oParentFeature As IJStructFeature
    
    eFeatureType = SF_CornerFeature
    If TypeOf oCornerParent Is IJStructFeature Then
        Set oParentFeature = oCornerParent
        eFeatureType = oParentFeature.get_StructFeatureType
    End If
    
    Dim oEndCutBoundedPort As IJPort
    Dim oEndCutBoundingPort As IJPort
    
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oParentFeature
            Set oEndCutBoundedPort = oWebCut.BoundedPort
            Set oEndCutBoundingPort = oWebCut.BoundingPort
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oParentFeature
            Set oEndCutBoundedPort = oFlangeCut.BoundedPort
            Set oEndCutBoundingPort = oFlangeCut.BoundingPort
        Case Else
            GetProjBoundingFlangeThickness = -1#
            Exit Function
    End Select
    
    ' --------------------
    ' Get the measurements
    ' --------------------
    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    Dim oMeasurements As New Collection
    
    GetConnectedEdgeInfo oCornerParent, oEndCutBoundedPort, oEndCutBoundingPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements
    
    ' -------------------------------------------
    ' If an input is the bottom of the top flange
    ' -------------------------------------------
    If xID1 = JXSEC_TOP_FLANGE_RIGHT Or xID2 = JXSEC_TOP_FLANGE_RIGHT Then
        GetProjBoundingFlangeThickness = oMeasurements.Item("DimPt15ToPt17")
    ' -------------------------------------------
    ' If an input is the top of the bottom flange
    ' -------------------------------------------
    ElseIf xID1 = JXSEC_BOTTOM_FLANGE_RIGHT Or xID2 = JXSEC_BOTTOM_FLANGE_RIGHT Then
        GetProjBoundingFlangeThickness = oMeasurements.Item("DimPt21ToPt23")
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function


Public Function GetSurfaceOffsetForEdgeCornerFeature(oCornerFeature As IJStructFeature) As Double

    Const sMETHOD = "GetSurfaceOffsetForEdgeCornerFeature"

    ' -------------------------------------------------------------------
    ' If the part on which the feature is placed is a member or stiffener
    ' -------------------------------------------------------------------
    Dim oSDOFeature As New StructDetailObjects.CornerFeature
    Set oSDOFeature.object = oCornerFeature
    
    ' ------------------------------------------
    ' Get the endcut creating the corner feature
    ' ------------------------------------------
    Dim oFeatureAsChild As IJDesignChild
    Dim oCornerParent As Object
    
    If TypeOf oCornerFeature Is IJDesignChild Then
        Set oFeatureAsChild = oCornerFeature
        Set oCornerParent = oFeatureAsChild.GetParent
    End If
    
    ' -----------------------------------
    ' If bounded by a member or stiffener
    ' -----------------------------------
    If TypeOf oSDOFeature.GetPartObject Is ISPSMemberPartPrismatic Or TypeOf oSDOFeature.GetPartObject Is IJProfile Then
        ' --------------------------------------
        ' Get the position of the corner feature
        ' --------------------------------------
        Dim eLocation As eEndCutCornerLocation
        Dim bIsBottomFlange As Boolean
        
        GetCornerFeaturePositionOnEndCut oCornerFeature, eLocation, bIsBottomFlange
    
        ' ---------------
        ' Get the overlap
        ' ---------------
        Dim dInsideOverlap As Double
        Dim dOutsideOverlap As Double
        
        Dim bIsBottomEdge As Boolean
        bIsBottomEdge = (eLocation = BottomEdgeInside Or eLocation = BottomEdgeOutside)
        GetEdgeOverlapAndClearance oCornerParent, _
                                   bIsBottomEdge, _
                                   bIsBottomFlange, _
                                   dInsideOverlap, _
                                   dOutsideOverlap
        
        If eLocation = BottomEdgeInside Or eLocation = TopEdgeInside Then

            GetSurfaceOffsetForEdgeCornerFeature = dInsideOverlap
        Else
            GetSurfaceOffsetForEdgeCornerFeature = dOutsideOverlap
        End If
    ' ---------------------
    ' If bounded by a plate
    ' ---------------------
    ElseIf TypeOf oSDOFeature.GetPartObject Is IJPlate Then
        ' ----------------------------------
        ' Get the edges used for the feature
        ' ----------------------------------
        Dim oStructFeatUtils As IJSDFeatureAttributes
        Dim oFacePortObj As Object
        Dim oEdgePort1Obj As Object
        Dim oEdgePort2Obj As Object
        
        Set oStructFeatUtils = New SDFeatureUtils
          
        oStructFeatUtils.get_CornerCutInputsEx oCornerFeature, _
                                               oFacePortObj, _
                                               oEdgePort1Obj, _
                                               oEdgePort2Obj
        
        If Not TypeOf oEdgePort1Obj Is IJStructPort Or Not TypeOf oEdgePort2Obj Is IJStructPort Then
            Exit Function
        End If
        
        Dim oStructPort1 As IJStructPort
        Dim oStructPort2 As IJStructPort
        Set oStructPort1 = oEdgePort1Obj
        Set oStructPort2 = oEdgePort2Obj
        
        Dim xID1 As JXSEC_CODE
        Dim xID2 As JXSEC_CODE
        
        xID1 = oStructPort1.SectionID
        xID2 = oStructPort2.SectionID
           
        ' -----------------------
        ' Get the bounding object
        ' -----------------------
    Dim eFeatureType As StructFeatureTypes
    Dim oParentFeature As IJStructFeature
    
    eFeatureType = SF_CornerFeature
    If TypeOf oCornerParent Is IJStructFeature Then
        Set oParentFeature = oCornerParent
        eFeatureType = oParentFeature.get_StructFeatureType
    End If
    
    Dim oEndCutBoundedPort As IJPort
    Dim oEndCutBoundingPort As IJPort
    
    Dim oWebCut As New StructDetailObjects.WebCut
    
    Select Case eFeatureType
        Case SF_WebCut
            Set oWebCut.object = oParentFeature
            Set oEndCutBoundedPort = oWebCut.BoundedPort
            Set oEndCutBoundingPort = oWebCut.BoundingPort
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oParentFeature
            Set oEndCutBoundedPort = oFlangeCut.BoundedPort
            Set oEndCutBoundingPort = oFlangeCut.BoundingPort
            
            Set oWebCut.object = oFlangeCut.WebCut

        Case Else
            GetSurfaceOffsetForEdgeCornerFeature = -1#
            Exit Function
    End Select
    
        ' ---------------------------------------------------------------------
        ' We currently only determine the overlap for the "FlushActual" web cut
        ' ---------------------------------------------------------------------
    ' The conditions below are indicative of a "FlushActual" web cut
        If TypeOf oEndCutBoundingPort.Connectable Is IJPlate And _
          (xID1 = JXSEC_IDEALIZED_BOUNDARY And (xID2 = JXSEC_TOP Or xID2 = JXSEC_BOTTOM)) Or _
          (xID2 = JXSEC_IDEALIZED_BOUNDARY And (xID1 = JXSEC_TOP Or xID1 = JXSEC_BOTTOM)) Then
          
        ' ---------------------------------------------------------------------------------
        ' Determine if plate base or offset port is the one we want based on surface normal
        ' ---------------------------------------------------------------------------------
        ' Get the ports
        Dim oSDOPlate As New StructDetailObjects.PlatePart
        Set oSDOPlate.object = oEndCutBoundingPort.Connectable
        
        Dim oBasePort As IJPort
        Dim oOffsetPort As IJPort
        
        Set oBasePort = oSDOPlate.BasePortBeforeChamfer(BPT_Base)
        Set oOffsetPort = oSDOPlate.BasePortBeforeChamfer(BPT_Offset)
        
        Dim oFlangePort As IJPort
        If xID1 = JXSEC_TOP Or xID2 = JXSEC_TOP Or (xID1 = e_JXSEC_MultipleBounding_5002 And xID2 = e_JXSEC_MultipleBounding_5003) Or (xID1 = e_JXSEC_MultipleBounding_5003 And xID2 = e_JXSEC_MultipleBounding_5002) Then
            
            Set oFlangePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
        Else
            Set oFlangePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
        End If
        
        ' Get the position of the feature
        Dim oLocation As IJDPosition
        Set oLocation = oWebCut.BoundedLocation
        
        ' Get a point on each port
        Dim oModelUtil As IJSGOModelBodyUtilities
        Set oModelUtil = New SGOModelBodyUtilities
        
        Dim oPointOnBase As IJDPosition
        Dim oPointOnOffset As IJDPosition
        Dim oPointOnFlange As IJDPosition
        Dim dist As Double
        
        Dim oBasePortGeom As IJSurfaceBody
        Dim oOffsetPortGeom As IJSurfaceBody
        Dim oFlangePortGeom As IJSurfaceBody
        Set oBasePortGeom = oBasePort.Geometry
        Set oOffsetPortGeom = oOffsetPort.Geometry
        Set oFlangePortGeom = oFlangePort.Geometry
        
        oModelUtil.GetClosestPointOnBody oBasePortGeom, oLocation, oPointOnBase, dist
        oModelUtil.GetClosestPointOnBody oOffsetPortGeom, oLocation, oPointOnOffset, dist
        oModelUtil.GetClosestPointOnBody oFlangePortGeom, oLocation, oPointOnFlange, dist
        
        ' Get the normals
        Dim oBaseNorm As IJDVector
        Dim oOffsetNorm As IJDVector
        Dim oFlangeNorm As IJDVector
        
        oBasePortGeom.GetNormalFromPosition oPointOnBase, oBaseNorm
        oOffsetPortGeom.GetNormalFromPosition oPointOnOffset, oOffsetNorm
        oFlangePortGeom.GetNormalFromPosition oPointOnFlange, oFlangeNorm
        
        Dim oPlateGeom As IJSurfaceBody
        
        If oFlangeNorm.Dot(oBaseNorm) > oFlangeNorm.Dot(oOffsetNorm) Then
            Set oPlateGeom = oBasePortGeom
        Else
            Set oPlateGeom = oOffsetPortGeom
        End If
        
        ' ---------------------------------------------------------
        ' Get distance between inside flange and top/bottom surface
        ' ---------------------------------------------------------
        Dim oPointOnBounding As IJDPosition
        
        oModelUtil.GetClosestPointsBetweenTwoBodies oPlateGeom, oFlangePortGeom, oPointOnBounding, oPointOnFlange, dist
        
        GetSurfaceOffsetForEdgeCornerFeature = dist
        
    End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'determine if the CF is at outer corner, inner corner, base corner
'make sure this method is called only when the CF is on slot
Public Sub DetermineCFPositionOnSlot(ByVal oEP1 As Object, ByVal oEP2 As Object, ByRef bIsBaseCF As Boolean, ByRef bIsInsideCF As Boolean, ByRef bOuterCF As Boolean)
Const sMETHOD = "DetermineCFPositionOnSlot"
On Error GoTo ErrorHandler

        Dim oAssoc As IJDAssocRelation
        Dim oRelColl As IJDRelationshipCol
        Set oRelColl = Nothing
        
        Dim dRelCount As Integer
        Dim oRelationShip As IJDRelationship
        
        Dim oEdge1ID As String, oEdge2ID As String
        oEdge1ID = "": oEdge2ID = ""
        
        If Not TypeOf oEP1 Is IJPort Then
            Set oAssoc = oEP1
            Set oRelColl = oAssoc.CollectionRelations(IID_IJProxyMember, "Member")
            
            For dRelCount = 1 To oRelColl.Count
                Set oRelationShip = oRelColl.Item(dRelCount)
                oEdge1ID = oRelationShip.Name
            Next
        End If
        
        Set oAssoc = Nothing
        Set oRelColl = Nothing
        Set oRelationShip = Nothing
        
        If Not TypeOf oEP2 Is IJPort Then
            Set oAssoc = oEP2
            Set oRelColl = oAssoc.CollectionRelations(IID_IJProxyMember, "Member")
        
            For dRelCount = 1 To oRelColl.Count
                Set oRelationShip = oRelColl.Item(dRelCount)
                oEdge2ID = oRelationShip.Name
            Next
        End If
        
        Select Case LCase(oEdge1ID)
            Case LCase("Slot:OutLine_257")
                If InStr(1, oEdge2ID, "770", vbTextCompare) > 0 Then
                    bIsInsideCF = True
                ElseIf oEdge2ID = "" Then
                    bIsBaseCF = True
                Else
                    bOuterCF = True
                End If
            Case LCase("Slot:OutLine_258")
                If InStr(1, oEdge2ID, "772", vbTextCompare) > 0 Then
                    bIsInsideCF = True
                ElseIf oEdge2ID = "" Then
                    bIsBaseCF = True
                Else
                    bOuterCF = True
                End If
            Case LCase("Slot:OutLine_770")
                If InStr(1, oEdge2ID, "257", vbTextCompare) > 0 Then
                    bIsInsideCF = True
                Else
                    bOuterCF = True
                End If
            Case LCase("Slot:OutLine_772")
                If InStr(1, oEdge2ID, "258", vbTextCompare) > 0 Then
                    bIsInsideCF = True
                Else
                    bOuterCF = True
                End If
            Case ""
                If InStr(1, oEdge2ID, "257", vbTextCompare) > 0 Or InStr(1, oEdge2ID, "258", vbTextCompare) > 0 Then
                    bIsBaseCF = True
                End If
            Case Else
                bOuterCF = True
        End Select

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub
'use this method only if the feature is on slot
Public Sub GetPortFromProxy(oEdgePort1 As Object, oFeature As IJStructFeature, oEP1 As IJPort)
Const sMETHOD = "GetPortFromProxy"
On Error GoTo ErrorHandler

        'check to see if the feature is on slot
        Dim bFeatureOnSlot As Boolean
        bFeatureOnSlot = False
        
        Dim oCFChild As IJDesignChild
        Dim oParentSlot As Object
    
        Set oCFChild = oFeature
        Set oParentSlot = oCFChild.GetParent
        Dim oStructFeature As IJStructFeature
        
        If TypeOf oParentSlot Is IJStructFeature Then
            Set oStructFeature = oParentSlot
    
            If oStructFeature.get_StructFeatureType = SF_Slot Then
                bFeatureOnSlot = True
            End If
        End If
        
        If Not bFeatureOnSlot Then
            Exit Sub
        End If

        Dim oSlotWrapper As New StructDetailObjects.Slot
        Set oSlotWrapper.object = oStructFeature
        
        Dim oAssoc As IJDAssocRelation
        Dim oRelColl As IJDRelationshipCol
        Set oRelColl = Nothing
        
        Dim dRelCount As Integer
        Dim oRelationShip As IJDRelationship
        
        Dim oEdge1ID As String
        oEdge1ID = ""
        
        Set oAssoc = oEdgePort1
        Set oRelColl = oAssoc.CollectionRelations(IID_IJProxyMember, "Member")
        
        For dRelCount = 1 To oRelColl.Count
            Set oRelationShip = oRelColl.Item(dRelCount)
            oEdge1ID = oRelationShip.Name
        Next
        
        If TypeOf oSlotWrapper.Penetrating Is IJProfilePart Then
            
            Dim oPenetrating As New StructDetailObjects.ProfilePart
            Set oPenetrating.object = oSlotWrapper.Penetrating
        
            If InStr(1, oEdge1ID, "257", vbTextCompare) > 0 Then
                Set oEP1 = oPenetrating.SubPort(JXSEC_WEB_LEFT)
            ElseIf InStr(1, oEdge1ID, "258", vbTextCompare) > 0 Then
                Set oEP1 = oPenetrating.SubPort(JXSEC_WEB_RIGHT)
            ElseIf InStr(1, oEdge1ID, "770", vbTextCompare) > 0 Then
                Set oEP1 = oPenetrating.SubPort(JXSEC_TOP_FLANGE_LEFT_BOTTOM)
            ElseIf InStr(1, oEdge1ID, "772", vbTextCompare) > 0 Then
                Set oEP1 = oPenetrating.SubPort(JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
            Else
                Set oEP1 = Nothing
            End If
        ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
            Dim oPenetratingPlate As New StructDetailObjects.PlatePart
            Set oPenetratingPlate.object = oSlotWrapper.Penetrating
            
            Dim oSlotMappingRule As IJSlotMappingRule
            Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
            
            Dim oBasePort As IJPort
            Dim oMappedPorts As JCmnShp_CollectionAlias
            Set oMappedPorts = New Collection
            oSlotMappingRule.GetEmulatedPorts oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, oBasePort, oMappedPorts
            
            Dim pSDOHelper As New StructDetailObjects.Helper
            
            If InStr(1, oEdge1ID, "257", vbTextCompare) > 0 Then
                Set oEP1 = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item("257"))
            ElseIf InStr(1, oEdge1ID, "258", vbTextCompare) > 0 Then
                Set oEP1 = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item("258"))
            ElseIf InStr(1, oEdge1ID, "770", vbTextCompare) > 0 Then
                Set oEP1 = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item("770"))
            ElseIf InStr(1, oEdge1ID, "772", vbTextCompare) > 0 Then
                Set oEP1 = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item("772"))
            Else
                Set oEP1 = Nothing
            End If
            
        End If
        
        Set oAssoc = Nothing
        Set oRelColl = Nothing
        Set oRelationShip = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number


End Sub

'This Method is used to get the CornerFeature Edge Lengths on with respect to bounding
Public Sub GetCornerFatureEdgeLengths(ByVal oCornerFeature As Object, ByVal oAC As Object, ByVal oStructPort1 As IJStructPort, _
                                ByVal oStructPort2 As IJStructPort, dLength1 As Double, dLength2 As Double, Optional oFacePort As IJPort = Nothing)

  On Error GoTo ErrorHandler
  Const sMETHOD = "GetCornerFatureEdgeLengths"
    
    dLength1 = 0
    dLength2 = 0
    
    If oStructPort1 Is Nothing Or oStructPort2 Is Nothing Then
        Exit Sub
    End If
    
    If oStructPort1.SectionID >= e_JXSEC_MultipleBounding_5001 And oStructPort1.SectionID <= e_JXSEC_MultipleBounding_5005 Then
    
        Dim oMD As IJDModelBody
        Dim oGeomIntersect As IMSModelGeomOps.DGeomOpsIntersect
        Set oGeomIntersect = New DGeomOpsIntersect
        
        Dim oSGOUitls As New SGOModelBodyUtilities
        Dim oVertices As New Collection
        Dim oIntersection As Object
        
        Dim oPosition1 As IJDPosition
        Dim oPosition2 As IJDPosition
                    
        Dim oPOrtColl As Collection
        Dim oEdgePort As IJPort
        Dim iCount As Integer
        
        Set oPOrtColl = New Collection
        
        oPOrtColl.Add oStructPort1
        oPOrtColl.Add oStructPort2
        
        dLength1 = 0
        dLength2 = 0
        
        'Get Face port
        If oFacePort Is Nothing Then
            Dim oStructFeatUtils As IJSDFeatureAttributes
            Set oStructFeatUtils = New SDFeatureUtils
            'Get Corner Feature inputs
            oStructFeatUtils.get_CornerCutInputsEx oCornerFeature, oFacePort, Nothing, Nothing
        End If
        
        For iCount = 1 To oPOrtColl.Count
        
            Set oEdgePort = oPOrtColl.Item(iCount)
                                                        
            'Get Edge Intersection on Face port
            On Error Resume Next
            oGeomIntersect.PlaceIntersectionObject Nothing, oFacePort.Geometry, oEdgePort.Geometry, Nothing, oIntersection
        
            On Error GoTo ErrorHandler
                    
            If Not oIntersection Is Nothing Then
                oSGOUitls.GetVertices oIntersection, oVertices
                
                If oVertices.Count = 2 Then
                    Set oPosition1 = oVertices.Item(1)
                    Set oPosition2 = oVertices.Item(2)
                    
                    'Calculate Edge lengths
                    If dLength1 = 0 Then
                        dLength1 = oPosition1.DistPt(oPosition2)
                    Else
                        dLength2 = oPosition1.DistPt(oPosition2)
                    End If
                End If
            Else
                Dim oEdgeFace As IJSurfaceBody
                Set oEdgeFace = oEdgePort.Geometry
                
                Dim oSDOCorner As New StructDetailObjects.CornerFeature
                Set oSDOCorner.object = oCornerFeature
                
                Dim oCornerPos As IJDPosition
                oSDOCorner.GetLocationOfCornerFeature oCornerPos
                ' -------------------------------------------------------------------------------------------
                ' Get all the vertices on the face and find the one furthest from the corner position
                ' -------------------------------------------------------------------------------------------
                Dim oVertices1 As Collection
                Dim oModelUtil As IJSGOModelBodyUtilities
                Set oModelUtil = New SGOModelBodyUtilities
                
                oModelUtil.GetVertices oEdgeFace, oVertices1
                
                Dim dDist As Double
                Dim dmaxDist As Double
                Dim dsecondmaxDist As Double
                Dim oPosition As IJDPosition
                Dim dEdgeLength As Double
                Dim jCount As Integer
                Dim maxpositionIndex As Integer
                dmaxDist = -1#
                 
                For jCount = 1 To oVertices1.Count
                   Set oPosition = oVertices1.Item(jCount)
                   dDist = oPosition.DistPt(oCornerPos)
                   If GreaterThan(dDist, dmaxDist) Then
                       dmaxDist = dDist
                       maxpositionIndex = jCount
                   End If
                Next jCount
                oVertices1.Remove (maxpositionIndex)
                ' ----------------------------------------------------------------------------------------------
                ' To find the second farthest from the corner position,which is the required edge length
                ' ----------------------------------------------------------------------------------------------
                dsecondmaxDist = -1#
                For Each oPosition In oVertices1
                    dDist = oPosition.DistPt(oCornerPos)
                    If GreaterThan(dDist, dsecondmaxDist) Then
                        dsecondmaxDist = dDist
                    End If
                Next oPosition
                dEdgeLength = dsecondmaxDist
                'Calculate Edge lengths
                If dLength1 = 0 Then
                    dLength1 = dEdgeLength
                Else
                    dLength2 = dEdgeLength
                End If
            
            End If

            Set oIntersection = Nothing
            Set oVertices = Nothing
            
        Next
        Exit Sub
    End If
    
    Dim oDesignChild As IJDesignChild
    Dim oCF_Parent As Object
    
    'Get CF Parent
    If TypeOf oCornerFeature Is IJDesignChild Then
        Set oDesignChild = oCornerFeature
        Set oCF_Parent = oDesignChild.GetParent
    End If
    
    Dim eFeatureType As StructFeatureTypes
    Dim oParentFeature As IJStructFeature
    
    'Get Parent Feature Type
    If TypeOf oCF_Parent Is IJStructFeature Then
        Set oParentFeature = oCF_Parent
        eFeatureType = oParentFeature.get_StructFeatureType
    End If
    Set oDesignChild = Nothing
    Set oCF_Parent = Nothing
    Set oCornerFeature = Nothing
    
    Dim oEndCutBoundedPort As IJPort
    Dim oEndCutBoundingPort As IJPort
    Dim oEndCut As Object
    
    'Check the Parent Feature Type
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oParentFeature
            Set oEndCutBoundedPort = oWebCut.BoundedPort
            Set oEndCutBoundingPort = oWebCut.BoundingPort
            Set oEndCut = oParentFeature
            Set oWebCut = Nothing
            
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oParentFeature
            Set oEndCutBoundedPort = oFlangeCut.BoundedPort
            Set oEndCutBoundingPort = oFlangeCut.BoundingPort
            Set oEndCut = oParentFeature
            Set oFlangeCut = Nothing

        Case Else
            Exit Sub
    End Select
        
    Set oParentFeature = Nothing
    
    'Get the Edge Lengths on Bounding if Bounding StiffenerProfilePart
    If TypeOf oEndCutBoundingPort.Connectable Is IJStructProfilePart Then

        Dim oITF As ConnectedEdgeInfo
        Dim oIBF As ConnectedEdgeInfo
        Dim oTF As ConnectedEdgeInfo
        Dim oBF As ConnectedEdgeInfo
        Dim oMeasurements As New Collection
        Dim bWebPenetrated As Boolean
        bWebPenetrated = True
        
        Dim oACOrEC As Object
        Set oACOrEC = IIf(GetMbrAssemblyConnectionType(oAC) = ACType_Mbr_Generic Or _
                          GetMbrAssemblyConnectionType(oAC) = ACType_Stiff_Generic, oEndCut, oAC)

        'Get the Measurements Lengths
        GetConnectedEdgeInfo oACOrEC, oEndCutBoundedPort, oEndCutBoundingPort, _
                                    oTF, oBF, oITF, oIBF, oMeasurements, bWebPenetrated
        Set oEndCutBoundedPort = Nothing
        Set oEndCutBoundingPort = Nothing
        
        If oMeasurements Is Nothing Then
            Exit Sub
        End If
        
        Dim xID1 As JXSEC_CODE
        Dim xID2 As JXSEC_CODE
        
        xID1 = oStructPort1.SectionID
        xID2 = oStructPort2.SectionID

        If bWebPenetrated Then
            'At Top Corner
            If xID1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or xID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
                dLength1 = oMeasurements.Item("DimPt17ToPt18")
                dLength2 = oMeasurements.Item("DimPt18ToBottom")
            'At Bottom Corner
            ElseIf xID1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or xID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
                dLength1 = oMeasurements.Item("DimPt20ToPt21")
                dLength2 = oMeasurements.Item("DimPt20ToTop")
            End If
        Else
            'At Top Corner
            If xID1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or xID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
                dLength1 = oMeasurements.Item("DimPt17ToPt18")
                dLength2 = oMeasurements.Item("DimPt18ToFR")
            'At Bottom Corner
            ElseIf xID1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or xID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
                dLength1 = oMeasurements.Item("DimPt20ToPt21")
                dLength2 = oMeasurements.Item("DimPt20ToFL")
            End If
        End If
        
        
        Set oMeasurements = Nothing
    End If
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Sub


Public Sub GetCornerFeaturePositionOnEndCut(oFeature As IJStructFeature, _
                                            eLocation As eEndCutCornerLocation, _
                                            Optional bIsBottomFlange As Boolean)

    Const sMETHOD = "GetPortFromProxy"
    
    On Error GoTo ErrorHandler
    
    eLocation = EndCutCornerLocationUnknown
    bIsBottomFlange = False
    
    ' ------------------------------------
    ' Determine where the feature is going
    ' ------------------------------------
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Dim oFacePortObj As Object
    Dim oEdgePort1Obj As Object
    Dim oEdgePort2Obj As Object
    
    Set oStructFeatUtils = New SDFeatureUtils
      
    oStructFeatUtils.get_CornerCutInputsEx oFeature, _
                                           oFacePortObj, _
                                           oEdgePort1Obj, _
                                           oEdgePort2Obj
    
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oEdgePort1 = oEdgePort1Obj
    Set oEdgePort2 = oEdgePort2Obj
    
    Dim xID1 As JXSEC_CODE
    Dim xID2 As JXSEC_CODE
    
    xID1 = oEdgePort1.SectionID
    xID2 = oEdgePort2.SectionID
        
    If xID2 = -1 Then
        eLocation = WebAndFlangeCommon
    ElseIf (xID1 = JXSEC_BOTTOM_FLANGE_RIGHT And (xID2 = JXSEC_BOTTOM Or xID2 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM)) Or _
       (xID2 = JXSEC_BOTTOM_FLANGE_RIGHT And (xID1 = JXSEC_BOTTOM Or xID1 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM)) Then
        eLocation = BottomEdgeOutside
    ElseIf (xID1 = JXSEC_BOTTOM_FLANGE_RIGHT And xID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Or _
           (xID2 = JXSEC_BOTTOM_FLANGE_RIGHT And xID1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Then
        eLocation = BottomEdgeInside
    ElseIf (xID1 = JXSEC_TOP_FLANGE_RIGHT And xID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Or _
           (xID2 = JXSEC_TOP_FLANGE_RIGHT And xID1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
        eLocation = TopEdgeInside
    ElseIf (xID1 = JXSEC_TOP_FLANGE_RIGHT And (xID2 = JXSEC_TOP Or xID2 = JXSEC_TOP_FLANGE_RIGHT_TOP)) Or _
           (xID2 = JXSEC_TOP_FLANGE_RIGHT And (xID1 = JXSEC_TOP Or xID1 = JXSEC_TOP_FLANGE_RIGHT_TOP)) Then
        eLocation = TopEdgeOutside
    ElseIf (xID1 = JXSEC_WEB_RIGHT And xID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Or _
           (xID2 = JXSEC_WEB_RIGHT And xID1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Then
        eLocation = FaceBottomInsideCorner
    ElseIf (xID1 = JXSEC_WEB_RIGHT And xID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Or _
           (xID2 = JXSEC_WEB_RIGHT And xID1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
        eLocation = FaceTopInsideCorner
    ElseIf (xID1 = e_JXSEC_MultipleBounding_5001 And xID2 = e_JXSEC_MultipleBounding_5002) Or (xID1 = e_JXSEC_MultipleBounding_5002 And xID2 = e_JXSEC_MultipleBounding_5001) Then
        eLocation = TopEdgeInside
    ElseIf (xID1 = e_JXSEC_MultipleBounding_5002 And xID2 = e_JXSEC_MultipleBounding_5003) Or (xID1 = e_JXSEC_MultipleBounding_5003 And xID2 = e_JXSEC_MultipleBounding_5002) Then
        eLocation = BottomEdgeInside
    End If
    
    ' ---------------------------------------
    ' Determine if it is on the bottom flange
    ' ---------------------------------------
    Dim sIsBottomFlange As String
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oFeature, sParentItemName, oParentObj

    If Not oParentObj Is Nothing Then
        GetSelectorAnswer oParentObj, "BottomFlange", sIsBottomFlange
        If sIsBottomFlange = gsYes Then
            bIsBottomFlange = True
        End If
    End If

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub

