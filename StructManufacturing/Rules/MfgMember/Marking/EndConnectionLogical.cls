VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndConnectionLogical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.EndConnectionLogical"


Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "EndConnectionLogical: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    'This line is added because the method "Connection_ContourProfileEnd" is not implemented for beam part.
    'It can be removed, once it is implemented
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    'Create the SD profile Wrapper and initialize it
    Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oMfgProfileOutput As IJMfgProfileOutput
        
    If Not oProfileWrapper.ProfileHasMfgPart(oMfgProfilePart) Then Exit Function
    Set oMfgProfileOutput = oMfgProfilePart

    'Getting the Active Faces of the Member to be Manufactured
    'Currently only PC's on WebLeft and WebRight are Marked
    Dim oFaces() As Variant
    Dim varFace As Variant
    
    oMfgProfileOutput.GetPartFaces StructMfgAnySide, StructMfgAsAny, oFaces

    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    Dim oWebLeftPort As IJPort
    Dim oWebRightPort As IJPort
    
    Dim oWebLeftSB As IJSurfaceBody
    Dim oWebRightSB As IJSurfaceBody
    
    If UBound(oFaces) >= LBound(oFaces) Then
        For Each varFace In oFaces
            If varFace = JXSEC_WEB_LEFT Then
                Set oWebLeftPort = oSDMemberWrapper.SubPort(JXSEC_WEB_LEFT)
                Set oWebLeftSB = oWebLeftPort.Geometry
            End If
            If varFace = JXSEC_WEB_RIGHT Then
                Set oWebRightPort = oSDMemberWrapper.SubPort(JXSEC_WEB_RIGHT)
                Set oWebRightSB = oWebRightPort.Geometry
            End If
        Next
    Else
        Set oWebLeftPort = oSDMemberWrapper.SubPort(JXSEC_WEB_LEFT)
        Set oWebLeftSB = oWebLeftPort.Geometry
        Set oWebRightPort = oSDMemberWrapper.SubPort(JXSEC_WEB_RIGHT)
        Set oWebRightSB = oWebRightPort.Geometry
    End If
    
    'Get the Profile Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)

    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData
    Dim iCount As Integer
    
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn

    ' Loop thru each Physical Connections
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
    Dim lGeomCount As Long
    lGeomCount = 1

    Dim dWebLength As Double
    dWebLength = oSDMemberWrapper.WebLength
    
    For iCount = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(iCount)
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Or TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
            oConnectionData = oConObjsCol.Item(iCount)
            
            Dim oThisPort As IJPort
            Dim oAppConnection As IJAppConnection
            
            Set oThisPort = oConnectionData.ConnectingPort
            Set oAppConnection = oConnectionData.AppConnection

            Set oSDPhysicalConn.object = oAppConnection
            
            Dim oCSMarkColl As IJElements
            
            'Creating the End Connection Marks on Members
            If oThisPort Is oWebLeftPort Or oThisPort Is oWebRightPort Then
                Set oCSMarkColl = CreateEndConnLogicalMarks(oConnectionData.ToConnectable, oThisPort, dWebLength)
            Else
                GoTo NextItem
            End If
            
            If oCSMarkColl Is Nothing Then
                GoTo NextItem
            ElseIf oCSMarkColl.Count = 0 Then
                GoTo NextItem
            End If

            Dim oMarkCS1 As IJComplexString
            Dim oMarkCS2 As IJComplexString

            Set oMarkCS1 = oCSMarkColl.Item(1)
            If oCSMarkColl.Count = 2 Then
                Set oMarkCS2 = oCSMarkColl.Item(2)
            End If

            Dim jCount As Long

            For jCount = 1 To oCSMarkColl.Count
                'Create a SystemMark object to store additional information
                 Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

                'Set the marking side
                oSystemMark.SetMarkingSide oProfileWrapper.GetSide(oThisPort)

                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

                oMarkingInfo.thickness = oSDMemberWrapper.WebThickness
                oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle

                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                If jCount = 1 Then
                    oGeom3d.PutGeometry oMarkCS1
                Else
                    oGeom3d.PutGeometry oMarkCS2
                End If
                oGeom3d.PutGeometrytype STRMFG_END_MARK
                oGeom3d.FaceId = UpSide

                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oAppConnection)
                oGeom3d.PutMoniker oMoniker

                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1
            Next jCount

            Set oMarkCS1 = Nothing
            Set oMarkCS2 = Nothing
            Set oCSMarkColl = Nothing
            Set oThisPort = Nothing
            Set oAppConnection = Nothing
        End If
NextItem:
    Next iCount
        
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing
    Set oMfgProfileOutput = Nothing
    Set oMfgProfilePart = Nothing
    Set oResourceManager = Nothing
    Set oWebLeftPort = Nothing
    Set oWebLeftSB = Nothing
    Set oWebRightPort = Nothing
    Set oWebRightSB = Nothing
    Set oMarkingInfo = Nothing
    Set oSDPhysicalConn = Nothing
    Set oGeom3d = Nothing
    Set oSystemMark = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2003, , "RULES")
    GoTo CleanUp
End Function

'**************************************************************************************
' Method       : CreateEndConnLogicalMarks
' Abstract     : This Function gets the Profile part and Plate/Profile Port object as input and gives the
'                EndConnectionMarks as a collection as output
'
'**************************************************************************************

Private Function CreateEndConnLogicalMarks(ByVal oConnectedPart As Object, ByVal oPort As IJPort, ByVal dMemberWebLength As Double) As IJElements
    Const METHOD = "CreateEndConnLogicalMarks"
    On Error GoTo ErrorHandler
    
    '** Declarations required for generating the mark**'
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oWebLeftPort As IJPort
    Dim oTopFacePort As IJPort
    
    Dim oWebLeftSurfaceBody As IJSurfaceBody
    Dim oTopSurfaceBody As IJSurfaceBody
    Dim oPartSurface As IJSurfaceBody
    Dim oWebLeftWB As IJWireBody
    Dim oWebLeftCS As IJComplexString
    
    Dim dWebLength As Double
    
    Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
    
    If TypeOf oConnectedPart Is IJProfilePart Then
        Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = oConnectedPart
        
        dWebLength = oSDProfileWrapper.WebLength
        
        '**Getting the WebLeft and TopFace Port**'
        Set oWebLeftPort = oSDProfileWrapper.SubPort(JXSEC_WEB_LEFT)
        Set oTopFacePort = oSDProfileWrapper.SubPort(JXSEC_TOP)
        
    ElseIf TypeOf oConnectedPart Is ISPSMemberPartPrismatic Then
        Dim oSDMemberWrapper As StructDetailObjects.MemberPart
        Set oSDMemberWrapper = New StructDetailObjects.MemberPart
        Set oSDMemberWrapper.object = oConnectedPart
        
        dWebLength = oSDMemberWrapper.WebLength
        
        '**Getting the WebLeft and TopFace Port**'
        Set oWebLeftPort = oSDMemberWrapper.SubPort(JXSEC_WEB_LEFT)
        Set oTopFacePort = oSDMemberWrapper.SubPort(JXSEC_TOP)
    End If

    'Create an instance of the StrMfg Math Geom helper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    '**Start of Code for End Connection Mark SKDY Specific**
 
    '**Getting the Geometry of the Webleft and TopFace**'
    Set oWebLeftSurfaceBody = oWebLeftPort.Geometry
    Set oTopSurfaceBody = oTopFacePort.Geometry
    
    '**Getting the Plate Surface to which profiles are connected**'
    Set oPartSurface = oPort.Geometry
    
    Dim oGeomOffset As IJGeometryOffset
    Set oGeomOffset = New DGeomOpsOffset
    
    Dim dExtend As Double
    dExtend = 0.005
    
    On Error Resume Next ' As the Extend sheet body routine may fail
    
    'Extendind the SurfaceBody to ensure that intersection occurs between plate and WebLeft of Profile
    Dim oWebLeftExtendedSB As IJSurfaceBody
    oGeomOffset.CreateExtendedSheetBody Nothing, oWebLeftSurfaceBody, Nothing, dExtend, Nothing, oWebLeftExtendedSB
    
    If oWebLeftExtendedSB Is Nothing Then
        Set oWebLeftExtendedSB = oWebLeftSurfaceBody
    End If
    
    'Extendind the SurfaceBody to ensure that intersection occurs between plate and WebLeft of Profile
    Dim oExtendedPlateSB As IJSurfaceBody
    oGeomOffset.CreateExtendedSheetBody Nothing, oPartSurface, Nothing, dExtend, Nothing, oExtendedPlateSB
    
    If oExtendedPlateSB Is Nothing Then
        Set oExtendedPlateSB = oPartSurface
    End If

    '**Getting the intersections of WebLeft with Plate Surface and also TopFace with Plate Surface and assigning it to WireBody
    Dim oExtendedWB As IJWireBody
    Set oExtendedWB = GetIntersection(oWebLeftExtendedSB, oExtendedPlateSB)
    
    If oExtendedWB Is Nothing Then
        Exit Function
    End If
    
    Dim oExtendedTopFlangeWB As IJWireBody
    If Not oTopSurfaceBody Is Nothing Then
        Set oExtendedTopFlangeWB = GetIntersection(oTopSurfaceBody, oPartSurface)
    End If
    On Error GoTo ErrorHandler
    
    'Trimming the WebLeft Geometry Generated as it will be more than the WebLeft Geometry
    oMfgMGHelper.WireBodyToComplexString oExtendedWB, oWebLeftCS
    m_oMfgRuleHelper.TrimCurveEnds oWebLeftCS, dExtend
    
    Set oWebLeftWB = m_oMfgRuleHelper.ComplexStringToWireBody(oWebLeftCS)
    
    Dim oPointOnLine As IJDPosition
    Dim oWebLeftLine As IJLine
    
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    
    'Getting the Point for Creating the Horizontal line
    CreateLineByExtendingWireToInputSurface oWebLeftWB, oTopSurfaceBody, 0.001, oWebLeftLine, oPointOnLine
    
    '**Finding Vector Normal to the WebLeft Surface**'
    
    Dim oCOG As IJDPosition                 'Center of Gravity of Web Left Surface
    Dim oProjectionPoint As IJDPosition
    Dim oProjectionVector As IJDVector
    Dim oSurfaceNormal As IJDVector         'Vector Normal to Web Left Surface
    
    '***Setting the COG
    oWebLeftSurfaceBody.GetCenterOfGravity oCOG
    
    oMfgMGHelper.ProjectPointOnSurfaceBody oWebLeftSurfaceBody, oCOG, oProjectionPoint, oProjectionVector
    oWebLeftSurfaceBody.GetNormalFromPosition oProjectionPoint, oSurfaceNormal
    
    'Normalizing The Vector
    oSurfaceNormal.length = 1
    
    Dim oTopLineLeftPoint As IJDPosition
    Dim oTopLineRightPoint As IJDPosition
    
    Set oTopLineLeftPoint = New DPosition
    Set oTopLineRightPoint = New DPosition
    
    'Generating the Horizontal Line which indicate the flange

    Dim oCSColl As IJElements
    Set oCSColl = New JObjectCollection

    Dim oWebLeftLineCS As IJComplexString
    Set oWebLeftLineCS = New ComplexString3d

    oWebLeftLineCS.AddCurve oWebLeftLine, True

    
    Dim oWebLeftLineProjCS As IJElements
    Dim oWebLeftCurve As IJCurve
    Set oWebLeftCurve = oWebLeftLineCS
    
    oWebLeftCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    Dim oTempPoint1 As New DPosition
    Dim oTempPoint2 As New DPosition

    oTempPoint1.Set dStartX, dStartY, dStartZ
    oTempPoint2.Set dEndX, dEndY, dEndZ
    
    'Projecting The Web Left Line on the Plate Surface that will be manufactured
    oMfgMGHelper.ProjectCSToSurface oWebLeftLineCS, oPartSurface, Nothing, oWebLeftLineProjCS
    'Dim oWebLeftCurve As IJCurve
    Set oWebLeftCurve = oWebLeftLineProjCS.Item(1)
    
    oCSColl.Add oWebLeftCurve
    
    Dim oProjWebLeftWB As IJWireBody
    Set oProjWebLeftWB = m_oMfgRuleHelper.ComplexStringToWireBody(oWebLeftCurve)
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    
    Dim SurfFaces As IJElements
    Set SurfFaces = oMfgGeomHelper.GetPlatePartEdgesInIJElements(oPartSurface, True)
    
    oMfgMGHelper.ProjectPointOnSurfaceBody oPartSurface, oPointOnLine, oProjectionPoint, oProjectionVector
    oWebLeftCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ

    oTempPoint1.Set dStartX, dStartY, dStartZ
    oTempPoint2.Set dEndX, dEndY, dEndZ
    
    If (oProjectionPoint.DistPt(oTempPoint1)) <= (oProjectionPoint.DistPt(oTempPoint2)) Then
        Set oProjectionPoint = oTempPoint1
    Else
        Set oProjectionPoint = oTempPoint2
    End If
    
    Dim iCount As Integer
    Dim dMinDist As Double
    Dim dRequiredDist As Double
    dRequiredDist = 10000
    Dim oSurfaceMB As IJDModelBody
    For iCount = 1 To SurfFaces.Count
        Set oSurfaceMB = SurfFaces.Item(iCount)
        oSurfaceMB.GetMinimumDistanceFromPosition oProjectionPoint, oTempPoint1, dMinDist
        If dRequiredDist > dMinDist Then
            dRequiredDist = dMinDist
        End If
        
    Next

    If dRequiredDist > 0.001 Then
        oTopLineLeftPoint.x = oPointOnLine.x + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.x)
        oTopLineLeftPoint.y = oPointOnLine.y + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.y)
        oTopLineLeftPoint.z = oPointOnLine.z + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.z)
        
        oTopLineRightPoint.x = oPointOnLine.x + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.x))
        oTopLineRightPoint.y = oPointOnLine.y + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.y))
        oTopLineRightPoint.z = oPointOnLine.z + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.z))
        
        Dim oTopFaceLine As IJLine
        Set oTopFaceLine = New Line3d
        
        oTopFaceLine.DefineBy2Points oTopLineLeftPoint.x, oTopLineLeftPoint.y, oTopLineLeftPoint.z, oTopLineRightPoint.x, oTopLineRightPoint.y, oTopLineRightPoint.z
        Dim oTopFaceLineCS As IJComplexString
        Set oTopFaceLineCS = New ComplexString3d
        
        Set oTopFaceLineCS = New ComplexString3d
        oTopFaceLineCS.AddCurve oTopFaceLine, True
        
        oCSColl.Add oTopFaceLineCS
    End If

    Set CreateEndConnLogicalMarks = oCSColl
    
CleanUp:
    Set oMfgMGHelper = Nothing
    Set oWebLeftPort = Nothing
    Set oTopFacePort = Nothing
    Set oWebLeftSurfaceBody = Nothing
    Set oTopSurfaceBody = Nothing
    Set oWebLeftExtendedSB = Nothing
    Set oTopFaceLine = Nothing
    Set oTopLineLeftPoint = Nothing
    Set oTopLineRightPoint = Nothing
    Set oSurfaceNormal = Nothing
    Set oProjectionVector = Nothing
    Set oProjectionPoint = Nothing
    Set oCOG = Nothing
    Set oPointOnLine = Nothing
    Set oWebLeftLine = Nothing
    Set oPartSurface = Nothing
    Set oExtendedPlateSB = Nothing
    Set oWebLeftWB = Nothing
    Set oWebLeftCS = Nothing
    Set oGeomOffset = Nothing
    Set oSDProfileWrapper = Nothing
    Set oMfgGeomUtilWrapper = Nothing
    Set oWebLeftCurve = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'**************************************************************************************
' Method       : CreateLineByExtendingWireToInputSurface
' Abstract     : In this Function given a wire body and a surface,it tries to create a line
'                by extending wire to surface.
'
'       \  Wire
'        \
'         \
'          '  (Line created includes wire portion also)
'           '
'            '-------------  Surface
'**************************************************************************************

Private Sub CreateLineByExtendingWireToInputSurface( _
            ByVal oWireBody As IJWireBody, _
            ByVal oAnotherSurface As Object, _
            ByVal dDistanceTolerance As Double, _
            ByRef oCreatedLine As IJLine, ByRef oTopPos As IJDPosition)
    
    On Error GoTo ErrorHandler
   
    Dim oModelBodyUtils As SGOModelBodyUtilities
    Dim oPointOnWire As IJDPosition
    Dim oPointOnSurface As IJDPosition
    Dim dDistance As Double
   
    Set oCreatedLine = Nothing
    Set oModelBodyUtils = New SGOModelBodyUtilities
    
    'Getting the minimum distance betwwen Profile Top Surface and Web Left Line
    oModelBodyUtils.GetClosestPointsBetweenTwoBodies _
                    oWireBody, oAnotherSurface, _
                     oPointOnWire, oPointOnSurface, dDistance
    
    Dim oTempCS As IJComplexString
    Dim oCurve As IJCurve
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
   
    Dim oCurveSP As IJDPosition
    Dim oCurveEP As IJDPosition
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    Dim dOldStartX As Double, dOldStartY As Double, dOldStartZ As Double, dOldEndX As Double, dOldEndY As Double, dOldEndZ As Double
  
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    Set oTempCS = oMfgRuleHelper.WireBodyToComplexString(oWireBody)
    
    Set oCurve = oTempCS
    oCurve.EndPoints dOldStartX, dOldStartY, dOldStartZ, dOldEndX, dOldEndY, dOldEndZ
    Set oCurve = Nothing
    
    'Ensuring that the Web Left Line is Offset by the Required Distance
    oMfgRuleHelper.ExtendWire oTempCS, dDistance + ENDCONN_EXTENSION
   
    Set oCurve = oTempCS
    oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
  
    Set oCurveSP = New DPosition
    Set oCurveEP = New DPosition
     
    oCurveSP.Set dStartX, dStartY, dStartZ
    oCurveEP.Set dEndX, dEndY, dEndZ
    
    'Creating the Required Line And Getting the Required point
    Set oCreatedLine = New Line3d
    Set oTopPos = New DPosition
    
    If oPointOnSurface.DistPt(oCurveEP) < oPointOnSurface.DistPt(oCurveSP) Then
        oTopPos.Set dEndX, dEndY, dEndZ
        oCreatedLine.DefineBy2Points _
                         oTopPos.x, oTopPos.y, oTopPos.z, _
                        dOldStartX, dOldStartY, dOldStartZ
    Else
        oTopPos.Set dStartX, dStartY, dStartZ
        oCreatedLine.DefineBy2Points _
                         oTopPos.x, oTopPos.y, oTopPos.z, _
                         dOldEndX, dOldEndY, dOldEndZ
    End If
    
CleanUp:
    Set oModelBodyUtils = Nothing
    Set oPointOnWire = Nothing
    Set oCurveSP = Nothing
    Set oPointOnSurface = Nothing
    Set oCurveEP = Nothing
    Set oCurve = Nothing

    Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Sub

