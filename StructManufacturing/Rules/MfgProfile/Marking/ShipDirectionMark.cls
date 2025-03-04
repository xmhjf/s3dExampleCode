VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipDirection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.ShipDirection"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ShipDirection: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oGeomCol3dOut As IJMfgGeomCol3d
    Set oGeomCol3dOut = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_DIRECTION, ReferenceObjColl, oGeomCol3dOut
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3dOut
    
    If bSelectiveRecompute Then
        Dim ind As Integer
        Dim oModifiedElems As IJElements
        Set oModifiedElems = ReferenceObjColl
        
        For ind = 1 To oModifiedElems.Count
            If TypeOf oModifiedElems.Item(ind) Is IJMfgDefinition Then
                GoTo CreateDirectionMark
            End If
        Next ind
        
        Set oModifiedElems = Nothing
        Exit Function
    End If
        
CreateDirectionMark:
        
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

    Dim oProfileCon As IJConnectable
    Set oProfileCon = Part

    Dim oPortColl As IJElements
    oProfileCon.enumPorts oPortColl, PortFace

    Dim i As Long
    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    
    Dim oStructConnectable As IJStructConnectable
    Dim oLateralPortsCol As IJElements
    If Not TypeOf Part Is IJStructConnectable Then GoTo ErrorHandler
    Set oStructConnectable = Part
        
    'Get the Base and Offset Ports
    oStructConnectable.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol

    If oBasePort Is Nothing Or oOffsetPort Is Nothing Then GoTo ErrorHandler
    Dim oSurface As IJSurfaceBody
    Set oSurface = oBasePort.Geometry
    If oSurface Is Nothing Then GoTo ErrorHandler
        
    Dim oUpsideSurface As IUnknown
    Dim oSurfacePort As IJPort
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(UpSide)
    Set oUpsideSurface = oSurfacePort.Geometry
    
    ' Get the edgewire on the UpSide
    Dim oWire As IJWireBody
    Set oWire = m_oMfgRuleHelper.GetCommonGeometry(oUpsideSurface, oSurface, False)
    
    ' Find the middle point of the wire
    Dim oMidPos As IJDPosition
    Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oWire)
    
    Dim oCreateModifyUtilities As IJProfileAttributes
    Set oCreateModifyUtilities = New ProfileUtils
    
    ' Get the complete landing curve from the profile
    Dim oLandCurve As Object
    oCreateModifyUtilities.GetLandingCurveFromProfile Part, oLandCurve
    Dim oWireLandCurve As IJWireBody
    Set oWireLandCurve = oLandCurve
    
    ' Use the landing curve for determining direction string
    Dim oEndPos1 As IJDPosition
    Dim oEndPos2 As IJDPosition
    Dim oLandingDir As IJDVector
    
    'Get positions of LCs two ends
    oWireLandCurve.GetEndPoints oEndPos1, oEndPos2
    Set oLandingDir = oEndPos1.Subtract(oEndPos2)
    oLandingDir.length = 1
    
    Dim oPosLandCurve As IJDPosition
    
    Dim oToolBox As IJDTopologyToolBox
    Set oToolBox = New IMSModelGeomOps.DGeomOpsToolBox
    Dim oDirection As IJDVector
    
    ' The projection might fail in some cases, we don't want to error out, we just try from the other end
    On Error Resume Next
    ' Project the middle point of the port onto the landingcurve
    oToolBox.GetNearestPointOnWireBodyFromPoint oWireLandCurve, oMidPos, Nothing, oPosLandCurve
    oToolBox.ProjectPointOnWireBody oWireLandCurve, oPosLandCurve, oPosLandCurve, oDirection
    On Error GoTo ErrorHandler
    
    Dim oDirString As String
    Dim oCS As IJComplexString
    Dim bSucces As Boolean
    bSucces = False
    If Not oDirection Is Nothing Then
        ' Determine the direction string using landing curve
        oDirString = m_oMfgRuleHelper.GetDirection(oLandingDir)
        
        bSucces = CreateShipDirectionMarkLine(oMidPos, oBasePort, oDirection, UpSide, oCS)
    End If
    
    If oDirection Is Nothing Or Not bSucces Then
        ' If we for some reason fails to creare the mark for the base port, then try the offset port
        Set oWire = m_oMfgRuleHelper.GetCommonGeometry(oUpsideSurface, oOffsetPort.Geometry, False)
        Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oWire)

        ' Project the middle point of the port onto the landingcurve
        oToolBox.GetNearestPointOnWireBodyFromPoint oWireLandCurve, oMidPos, Nothing, oPosLandCurve
        oToolBox.ProjectPointOnWireBody oWireLandCurve, oPosLandCurve, oPosLandCurve, oDirection
        
        ' Determine the direction string using landing curve
        oLandingDir.length = -1         ' reverse landing curve direction since we are computing from other end
        oDirString = m_oMfgRuleHelper.GetDirection(oLandingDir)
        
        If Not CreateShipDirectionMarkLine(oMidPos, oSurfacePort, oDirection, UpSide, oCS) Then GoTo ErrorHandler
    End If

    Dim oGeom3d As IJMfgGeom3d
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As IJMarkingInfo

    If Not oCS Is Nothing Then
        Set oGeom3d = CreateGeom3dObject(oCS, STRMFG_DIRECTION, UpSide, oMarkingInfo)
        oMarkingInfo.name = oDirString
        oGeomCol3dOut.AddGeometry 1, oGeom3d
    End If
        
    '  Determine if a direction mark needs to be placed on the flange
    If IsBuiltUp(Part) = True Then
        Dim oProfileSupport As IJProfilePartSupport
        Set oProfileSupport = New ProfilePartSupport
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oProfileSupport
        Set oPartSupport.Part = Part
        
        Dim pContourColl As Collection, pMonikerColl As Collection
        Dim pFlangeSB As IJSurfaceBody
        
        ' only need to use the flange surfacebody
        oProfileSupport.GetProfileContours TopFlangeTopFace, pFlangeSB, pContourColl, pMonikerColl
        Set pMonikerColl = Nothing
        Set pContourColl = Nothing
        
        ' projection may fail and that is OK
        On Error Resume Next
        Dim oMGHelper As New MfgMGHelper
        Dim pSurfNorm As IJDVector
        Dim pProjCS As IJComplexString
        
        Set oMidPos = Nothing
        Set oWire = m_oMfgRuleHelper.GetCommonGeometry(pFlangeSB, oBasePort.Geometry, False)
        If Not oWire Is Nothing Then Set oMidPos = m_oMfgRuleHelper.GetMiddlePoint(oWire)

        If Not oMidPos Is Nothing Then
            Set oCS = Nothing
            If CreateShipDirectionMarkLine(oMidPos, oBasePort, oDirection, UpSide, oCS) Then
                Set oGeom3d = Nothing
                Set oSystemMark = Nothing
                Set oGeom3d = CreateGeom3dObject(oCS, STRMFG_DIRECTION, JXSEC_TOP, oMarkingInfo)
                oMarkingInfo.name = oDirString
                oGeomCol3dOut.AddGeometry 1, oGeom3d
            End If
        End If
    End If
   
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3dOut

CleanUp:
    Set oGeom3d = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2016, , "RULES")
    GoTo CleanUp
End Function

 
