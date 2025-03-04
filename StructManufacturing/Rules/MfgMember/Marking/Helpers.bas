Attribute VB_Name = "Helpers"
'*******************************************************************
'  Copyright (C) 2002 Global Research and Development, Inc.  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Helpers for the MfgPlateMarking rules.
'
'  History:
'       TBH         april 8. 2002   created
'       KONI    may. 17 2002    implementation
'       Venkat  Oct. 11 2006    Review MsgBox usage-TR:103745
'******************************************************************

Option Explicit

Const MODULE = " Helpers.bas"

Public m_bDebug As Boolean

Public m_oMfgRuleHelper As MfgRuleHelpers.Helper
Public m_oSystemMarkFactory As GSCADMfgSystemMark.MfgSystemMarkFactory
Public m_oGeom3dFactory As GSCADMfgGeometry.MfgGeom3dFactory
Public m_oGeom2dFactory As GSCADMfgGeometry.MfgGeom2dFactory
Public m_oGeomCol3dFactory As GSCADMfgGeometry.MfgGeomCol3dFactory
Public m_oGeomCol2dFactory As GSCADMfgGeometry.MfgGeomCol2dFactory

' ***********************************************************************************
' Public Sub Initialize
'
' Description: Helper to Initialize the most used objects
'
' ***********************************************************************************
Public Sub Initialize()
Const METHOD = "Initialize"

    On Error GoTo ErrorHandler
    
    m_bDebug = False
    
    Set m_oGeom3dFactory = New GSCADMfgGeometry.MfgGeom3dFactory
    Set m_oGeom2dFactory = New GSCADMfgGeometry.MfgGeom2dFactory
    Set m_oGeomCol3dFactory = New GSCADMfgGeometry.MfgGeomCol3dFactory
    Set m_oGeomCol2dFactory = New GSCADMfgGeometry.MfgGeomCol2dFactory
    
    Set m_oMfgRuleHelper = New MfgRuleHelpers.Helper
    Set m_oSystemMarkFactory = New GSCADMfgSystemMark.MfgSystemMarkFactory
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub


Public Sub UnInitialize()
    Set m_oMfgRuleHelper = Nothing
    Set m_oSystemMarkFactory = Nothing
    Set m_oGeom3dFactory = Nothing
    Set m_oGeom2dFactory = Nothing
    Set m_oGeomCol3dFactory = Nothing
    Set m_oGeomCol2dFactory = Nothing
End Sub


Public Function CreateFittingMarkLine(oProfile As Object, oPos As IJDPosition, oConnectionData As ConnectionData, UpSide As Long) As IJWireBody
Const METHOD = "CreateFittingMarkLine"
On Error GoTo ErrorHandler
    Dim endX As Double
    Dim endY As Double
    Dim endZ As Double
    
    'Find direction of mark
    Dim oPlateSurface As IJSurfaceBody
    Dim oPlatePort As IJPort
    Set oPlatePort = oConnectionData.ToConnectedPort
    
    ' Make sure the point is on the plate surface
    Dim oNewPos As IJDPosition
    Dim dDistance As Double
    Dim pModelBody As IJModelBody

    If TypeOf oProfile Is IJProfileER Then
        ' For EdgeReinforcement get the surface port from the plate instead of the connected port
        Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
        Set oMfgPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
        
        Set oMfgPlateWrapper.object = oPlatePort.Connectable
        
        Dim oPlateBasePort As IJPort
        Set oPlateBasePort = oMfgPlateWrapper.GetSurfacePort(JXSEC_WEB_LEFT)
        
        Set oPlateSurface = oPlateBasePort.Geometry
    Else
        Set oPlateSurface = oPlatePort.Geometry
    End If

    Set pModelBody = oPlateSurface
    
    Dim oGeomOps As IJSGOModelBodyUtilities
    Set oGeomOps = New SGOModelBodyUtilities
    oGeomOps.GetClosestPointOnBody pModelBody, oPos, oNewPos, dDistance

    Dim oDir As IJDVector
    oPlateSurface.GetNormalFromPosition oNewPos, oDir
        
    'Create temporary line
    Dim oLine As IJLine
    Set oLine = New Line3d
    endX = oPos.x + (oDir.x * FittingMarkLength)
    endY = oPos.y + (oDir.y * FittingMarkLength)
    endZ = oPos.z + (oDir.z * FittingMarkLength)
    oLine.DefineBy2Points oPos.x, oPos.y, oPos.z, endX, endY, endZ
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True
    
    Dim oWire As IJWireBody
    Set oWire = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)
    Set oCS = Nothing
    
    'Offset curve away from profile in order to project it back in (in case of mountingangle!=90)
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfile
    
    Dim oProfilePort As IJPort
    Set oProfilePort = oProfileWrapper.GetSurfacePort(UpSide)
    Dim oProfileSurface As IJSurfaceBody
    Set oProfileSurface = oProfilePort.Geometry
    
    Dim oProfilePos As IJDPosition
    Set oProfilePos = m_oMfgRuleHelper.ProjectPointOnSurface(oPos, oProfileSurface, oDir)
    
    Dim oSurface As IJSurfaceBody
    Set oSurface = m_oMfgRuleHelper.CreateSurfaceByPointNormal(oProfilePos.x, oProfilePos.y, oProfilePos.z, oDir.x, oDir.y, oDir.z)
   
    oSurface.GetNormalFromPosition oProfilePos, oDir
    
    Set oCS = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurface, oWire, oDir)
    
'    oProfileSurface.GetNormalFromPosition oProfilePos, oDir
'
'    Dim oOffSetWire As IJWireBody
'    Set oOffSetWire = m_oMfgRuleHelper.OffsetCurve(oProfileSurface, oProjWire, oDir, 0.2, True)
'
'
'
'    If oOffSetWire Is Nothing Then
'        GoTo ErrorHandler
'    End If
'
'    'Project line back on profilesurface
'    Dim oProfileDir As IJDPosition
'    Dim oFinalWire As IJWireBody
'    m_oMfgRuleHelper.ScaleVector oDir, -1
'
'    Set oFinalWire = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oProfileSurface, oOffSetWire, oDir)
    
    'Create line with correct length
'    Dim oStart As IJDPosition, oEnd As IJDPosition
'    oProjWire.GetEndPoints oStart, oEnd
'
'    Dim x As Double, y As Double, z As Double, x2 As Double, y2 As Double, z2 As Double
'    oStart.Get x, y, z
'    oEnd.Get x2, y2, z2
'
'    Dim oLine2 As IJLine
'    Set oLine2 = New Line3d
'    oLine2.SetStartPoint x, y, z
'    oLine2.SetEndPoint x2, y2, z2
'    oLine2.Length = FittingMarkLength
'
'    Set oCS = New ComplexString3d
'    oCS.AddCurve oLine2, True
    
    Set oWire = Nothing
    Set oWire = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)

    Set CreateFittingMarkLine = oWire
    
    Set oProfileSurface = Nothing
    Set oProfilePort = Nothing
    Set oPlateSurface = Nothing
    Set oPlatePort = Nothing
    Set oDir = Nothing
    Set oLine = Nothing
    Set oCS = Nothing
    Set oProfilePos = Nothing
'    Set oOffSetWire = Nothing
'    Set oProfileDir = Nothing
'    Set oFinalWire = Nothing
'    Set oStart = Nothing
'    Set oEnd = Nothing
    Set oCS = Nothing
'    Set oProjWire = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
'    ReportUnanticipatedError "Helper", "CreateFittingMarkLine"
End Function

'*******************************************************************
'   CreateSeamControlLine
'
'   Currently the same (almost) as CreateFittingMarkLine.
'   Eventually the line should have the same direction as the seam.
'*******************************************************************

Public Function CreateSeamControlLine(oProfile As Object, oPos As IJDPosition, oConnectionData As ConnectionData, UpSide As Long) As IJWireBody
Const METHOD = "CreateSeamControlLine"
On Error GoTo ErrorHandler
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfile
    
    ' Surface to put final mark
    Dim oUpsideSurface As IJPort, oPort1 As IJPort, oPort2 As IJPort
    Set oUpsideSurface = oProfileWrapper.GetSurfacePort(UpSide)

    Set oPort1 = oConnectionData.ConnectingPort
    Set oPort2 = oConnectionData.ToConnectedPort
    Dim oPhysWire As IJWireBody
    Set oPhysWire = m_oMfgRuleHelper.GetCommonGeometry(oPort1.Geometry, oPort2.Geometry)
    
    If (oPhysWire Is Nothing) Then
        Set CreateSeamControlLine = Nothing
        Exit Function
    End If
        
    Dim oSurface As IJSurfaceBody
    Dim bSucceed As Boolean
    
    bSucceed = m_oMfgRuleHelper.GetSurfaceFromWire(oPhysWire, oPos, oSurface)
    
    If bSucceed = False Then
        Exit Function
    End If
    
    Dim oFinalWire As IJWireBody
    Set oFinalWire = m_oMfgRuleHelper.GetCommonGeometry(oSurface, oUpsideSurface.Geometry)

    Set CreateSeamControlLine = oFinalWire

    Set oProfileWrapper = Nothing
    Set oUpsideSurface = Nothing
    Set oPort1 = Nothing
    Set oPort2 = Nothing
    Set oPhysWire = Nothing
    Set oSurface = Nothing
    Set oFinalWire = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
'    ReportUnanticipatedError "Helper", "CreateFittingMarkLine"
End Function

Private Function GetMarginValueOnThePort(oPort As IJPort) As Double

Const METHOD = "GetMarginValueOnThePort"
On Error GoTo ErrorHandler

Dim oMfgDefCol As Collection
Dim oConstMargin As IJConstMargin
Dim oObliqueMargin As IJObliqueMargin

Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)

If oMfgDefCol.Count > 0 Then
    Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
    lFabMargin = 0
    lAssyMargin = 0
    lCustomMargin = 0
    Dim j As Integer
    For j = 1 To oMfgDefCol.Count
        If TypeOf oMfgDefCol.Item(j) Is IJAssyMarginChild Then
            Set oConstMargin = oMfgDefCol.Item(j)
            lAssyMargin = lAssyMargin + oConstMargin.Value
        ElseIf TypeOf oMfgDefCol.Item(j) Is IJObliqueMargin Then
            Set oObliqueMargin = oMfgDefCol.Item(j)
            If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                lFabMargin = lFabMargin + oObliqueMargin.EndValue
            Else
                lFabMargin = lFabMargin + oObliqueMargin.StartValue
            End If
        ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
            Set oConstMargin = oMfgDefCol.Item(j)
            lFabMargin = lFabMargin + oConstMargin.Value
        'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
        End If
        
    Next j
    If lAssyMargin <> 0 Or lFabMargin <> 0 Or lCustomMargin <> 0 Then
         Dim TotMargin As Double
         TotMargin = lAssyMargin + lFabMargin + lCustomMargin
    End If
End If
GetMarginValueOnThePort = TotMargin

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function


'*******************************************************************
'   CreateGeom3dObject
'

'*******************************************************************

Public Function CreateGeom3dObject(oCurve As IJCurve, eType As StrMfgGeometryType, eUpside As Long, oMarkingInfo As IJMarkingInfo) As IJMfgGeom3d
Const METHOD = "CreateGeom3dObject"
On Error GoTo ErrorHandler
    
    Dim oGeom3d As IJMfgGeom3d
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oSystemMark As IJMfgSystemMark

    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)

    'Create a SystemMark object to store additional information
    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    Set oObjSystemMark = oSystemMark
    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oObjSystemMark)

    oGeom3d.PutMoniker oMoniker
    oGeom3d.PutGeometry oCurve
    oGeom3d.PutGeometrytype eType
    oGeom3d.FaceId = eUpside

    'Set the marking side
    oSystemMark.SetMarkingSide eUpside
    oSystemMark.Set3dGeometry oGeom3d
    Set oMarkingInfo = oSystemMark

    Set CreateGeom3dObject = oGeom3d
    
    Set oMoniker = Nothing
    Set oGeom3d = Nothing
    Set oObjSystemMark = Nothing
    Set oSystemMark = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function


'*******************************************************************
'   IsBuiltUp
'
'*******************************************************************

Public Function IsBuiltUp(oPart As Object) As Boolean
Const METHOD = "IsBuiltUp"
On Error GoTo ErrorHandler

    'Create the SD profile Wrapper and initialize it
    Dim oSDProfileWrapper As Object
    If TypeOf oPart Is IJStiffenerPart Then
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = oPart
    ElseIf TypeOf oPart Is IJBeamPart Then
        Set oSDProfileWrapper = New StructDetailObjects.BeamPart
        Set oSDProfileWrapper.object = oPart
    ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
        Set oSDProfileWrapper = New StructDetailObjects.MemberPart
        Set oSDProfileWrapper.object = oPart
    End If
    
    ' determine if cross-section is not a builtup - we may need to add direction mark to multiple surfaces
    IsBuiltUp = oSDProfileWrapper.IsCrossSectionABuiltUp
    Set oSDProfileWrapper = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Function GetPhysicalConnectionData(ByVal oThisPart As Object, ByVal oReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal bReturnAll As Boolean) As Collection

Const METHOD = "GetPhysicalConnectionData"
On Error GoTo ErrorHandler

    Dim index As Long
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oConnection As IJAppConnection
    Dim aConnectionData As ConnectionData
    
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = oThisPart
    
    Set GetPhysicalConnectionData = New Collection
    
    For index = 1 To oReferenceObjColl.Count
        If TypeOf oReferenceObjColl.Item(index) Is IJStructPhysicalConnection Then
            Dim bIsCrossOfTee As Boolean
            Dim oConnType As ContourConnectionType
            
            Set oConnection = oReferenceObjColl.Item(index)
            oSDPartSupport.GetConnectionTypeForContour oConnection, _
                                                       oConnType, _
                                                       bIsCrossOfTee
    
            If bReturnAll = True Or (oConnType = PARTSUPPORT_CONNTYPE_LAP) Or _
               (oConnType = PARTSUPPORT_CONNTYPE_PROFILE_END) Or _
               (oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsCrossOfTee) Then
               
               Dim oPortElements As IJElements
               oConnection.enumPorts oPortElements
               
               Dim oPort1 As IJPort
               Dim oPort2 As IJPort
               
               Set oPort1 = oPortElements.Item(1)
               Set oPort2 = oPortElements.Item(2)
               
               If (oPort1.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oPort1
                    Set aConnectionData.ToConnectable = oPort2.Connectable
                    Set aConnectionData.ToConnectedPort = oPort2
                    GetPhysicalConnectionData.Add aConnectionData
               ElseIf (oPort2.Connectable Is oThisPart) Then
                    Set aConnectionData.AppConnection = oConnection
                    Set aConnectionData.ConnectingPort = oPort2
                    Set aConnectionData.ToConnectable = oPort1.Connectable
                    Set aConnectionData.ToConnectedPort = oPort1
                    GetPhysicalConnectionData.Add aConnectionData
               End If
            End If
        End If
    
    Next index
    
    Set oSDPartSupport = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function

' ***********************************************************************************
' Public Function CreateShipDirectionMarkLine
'
' Description:  Method will create a small line as a ComplexString,
'               and the direction information should be put in the MarkingInfo.Name as Direction.
'               The line is only there to indicate a position on the plate.
'
' ***********************************************************************************
Public Function CreateShipDirectionMarkLine(oPos As IJDPosition, oSurfacePort As IJPort, oDirection As IJDVector, UpSide As Long, ByRef oMark As IJComplexString) As Boolean
Const METHOD = "CreateShipDirectionMarkLine"
On Error GoTo ErrorHandler
    Dim dLineLength As Double
    dLineLength = 0.01 ' As small as possible (No need for visible representation)
    Dim endX As Double
    Dim endY As Double
    Dim endZ As Double
        
    ' make sure the start point is on the surface
    Dim oSurface As IJSurfaceBody
    Set oSurface = oSurfacePort.Geometry
    If (oSurface Is Nothing) Then GoTo ErrorHandler

    Dim oMGHelper As New MfgMGHelper
    Dim oProjPos As IJDPosition
    Dim oNorm As IJDVector
    oMGHelper.ProjectPointOnSurfaceBody oSurface, oPos, oProjPos, oNorm
    
    If (oProjPos Is Nothing) Then
        Dim lErrNumber As Long
        lErrNumber = LogMessage(Err, MODULE, METHOD, " projection failed ")
        GoTo ErrorHandler
    End If
   
    Dim oLine As IJLine
    Set oLine = New Line3d
    endX = oProjPos.x + (oDirection.x * dLineLength)
    endY = oProjPos.y + (oDirection.y * dLineLength)
    endZ = oProjPos.z + (oDirection.z * dLineLength)
    oLine.DefineBy2Points oProjPos.x, oProjPos.y, oProjPos.z, endX, endY, endZ
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oLine, True

    Set oMark = oCS
    If oMark Is Nothing Then GoTo ErrorHandler
        
    CreateShipDirectionMarkLine = True

Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    CreateShipDirectionMarkLine = False
End Function



Public Function CreateMemberLocationMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, bAddFittingMark As Boolean) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
     Const METHOD = "Location: CreateMemberLocationMark"
    
    On Error GoTo ErrorHandler
    
    Dim oSDMemberWrapper As Object
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oPartInfo As IJDPartInfo
    Set oPartInfo = New PartInfo
    
    'Get the Profile Part Physically Connected Objects
    Dim oConObjsCol As Collection
    'Set oConObjsCol = oSDMemberWrapper.ConnectedObjects
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the function
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long
    Dim oSDConProfileWrapper As New StructDetailObjects.ProfilePart
    Dim oSDConMemberWrapper As New StructDetailObjects.MemberPart
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim oVector As IJDVector
    Dim bContourTee As Boolean
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
    Dim name As String
    Dim direction As String
    Dim thickness As Double
    Dim MfgType As StrMfgGeometryType
    Dim lGeomCount As Long
    lGeomCount = 1
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        'Get the sub port having section information
        Dim oSubPort As IJStructPort
        Set oSubPort = oProfileWrapper.GetProfileSubPort(oConnectionData.ConnectingPort, oConnectionData.AppConnection)
        If Not ((oSubPort.SectionID = JXSEC_WEB_RIGHT) Or (oSubPort.SectionID = JXSEC_OUTER_TUBE) Or _
            (oSubPort.SectionID = JXSEC_WEB_LEFT)) Then
            Set oSubPort = Nothing
            GoTo NextItem
        End If
        
        Dim eMoldedDir As StructMoldedDirection
        Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
        
        eMoldedDir = invalidDirection
        eSideOfConnectedObjectToBeMarked = SideA
        Dim sMoldedSide As String
        'Initialize the profile wrapper and the Physical Connection wrapper
        If TypeOf oConnectionData.ToConnectable Is IJPlatePart Or _
           TypeOf oConnectionData.AppConnection Is IJSmartPlate Or _
           TypeOf oConnectionData.AppConnection Is IJCollarPart Then
            Set oSDConPlateWrapper.object = oConnectionData.ToConnectable
            name = oSDConPlateWrapper.name
            thickness = oSDConPlateWrapper.PlateThickness
            
             'Initialize the profile wrapper and the Physical Connection wrapper
            Set oSDConPlateWrapper.object = oConnectionData.ToConnectable
            
            eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(oConnectionData.ToConnectable)
            
            If eMoldedDir = Centered Then
                eSideOfConnectedObjectToBeMarked = SideUnspecified
            Else
                ' Fix for TR#62896
                ' Earlier we were filling the side to be marked always SideA which is wrong
                ' It should be decided based on the MoldedSide of the Connected plate
                
                ' Get the moulded side of the connected plate
                sMoldedSide = oSDConPlateWrapper.MoldedSide
                If sMoldedSide = "Base" Then
                    eSideOfConnectedObjectToBeMarked = SideA
                ElseIf sMoldedSide = "Offset" Then
                    eSideOfConnectedObjectToBeMarked = SideB
                End If
            End If
            MfgType = STRMFG_PLATELOCATION_MARK
        ElseIf TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDConProfileWrapper.object = oConnectionData.ToConnectable
            name = oSDConProfileWrapper.name
            thickness = oSDConProfileWrapper.WebThickness
            
            MfgType = STRMFG_PROFILELOCATION_MARK
            
            'Get the Thickness side to be marked for Profile which is based on the position of web
            'w.r.t load point
            Dim oProfilePartSupport As IJProfilePartSupport
            Dim oPartSupp As IJPartSupport
            
            Set oPartSupp = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupp.Part = oSDConProfileWrapper.object
            
            Set oProfilePartSupport = oPartSupp
            eSideOfConnectedObjectToBeMarked = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
            
            Set oProfilePartSupport = Nothing
            Set oPartSupp = Nothing
        ElseIf TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
            eSideOfConnectedObjectToBeMarked = SideA
            MfgType = STRMFG_MEMBER_LOCATION_MARK
            
            Set oSDConMemberWrapper.object = oConnectionData.ToConnectable
            name = oSDConMemberWrapper.name
            thickness = oSDConMemberWrapper.WebThickness
            
        End If
        
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        bContourTee = oSDMemberWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oWB, oVector)
                                                   
        If (bContourTee = True) And (Not (oWB Is Nothing)) Then
            Dim oCSColl As IJElements
            'Convert the IJWireBody to ComplexStrings
            Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWB)

            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
            If (oCSColl Is Nothing) Then
                GoTo NextItem
            End If

            For Each oCS In oCSColl
                On Error GoTo NextComplexString
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
                'Set the marking side
                Dim MarkingFace As Long
                MarkingFace = oProfileWrapper.GetSide(oSubPort)
                oSystemMark.SetMarkingSide MarkingFace
    
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
                oMarkingInfo.name = name
                oMarkingInfo.thickness = thickness
                oMarkingInfo.FittingAngle = oSDPhysicalConn.TeeMountingAngle
                
                If eSideOfConnectedObjectToBeMarked = SideUnspecified Then
                    oMarkingInfo.direction = "centered"
                Else
                    direction = m_oMfgRuleHelper.GetDirection(oVector)
                    oMarkingInfo.direction = direction
                    If Not oSDConPlateWrapper Is Nothing Then
                        oMarkingInfo.ThicknessDirection = GetThicknessDirectionVector(oCS, oSDConPlateWrapper, sMoldedSide) 'oVector
                    ElseIf Not oSDConProfileWrapper Is Nothing Then
                        oMarkingInfo.ThicknessDirection = GetThicknessDirectionVector(oCS, oSDConProfileWrapper, sMoldedSide) 'oVector
                    End If
                End If
                
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype MfgType
                oGeom3d.FaceId = MarkingFace
                Set oObjSystemMark = oSystemMark
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    
                oGeom3d.PutMoniker oMoniker
                oSystemMark.Set3dGeometry oGeom3d
                    
                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1
NextComplexString:
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oGeom3d = Nothing
            Next
            oCSColl.Clear
            Set oCSColl = Nothing
        End If
NextItem:
        Set oSubPort = Nothing
        Set oWB = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConProfileWrapper = Nothing
        Set oSDPhysicalConn = Nothing
        name = ""
        thickness = 0
    Next nIndex
    
    'Return the 3d collection
    Set CreateMemberLocationMark = oGeomCol3d
    
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing

CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oPartInfo = Nothing
    Set oSDConProfileWrapper = Nothing
    Set oSDConMemberWrapper = Nothing
    Set oSDConPlateWrapper = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2008, , "RULES")
    GoTo CleanUp
End Function

Public Sub GetCorrectPort(ByVal pInPort As IJPort, ByRef pOutPort As IJPort)
    Const METHOD = "GetCorrectPort"
        
    Dim oConnectable As IJStructConnectable
    If TypeOf pInPort.Connectable Is IJProfilePart Or TypeOf pInPort.Connectable Is IJPlatePart Then
        Set oConnectable = pInPort.Connectable
    Else 'This is Member -- get IJStructconnectable from MemberPartSupport
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New MemberPartSupport
        Set oPartSupport.Part = pInPort.Connectable
        Set oConnectable = oPartSupport
        Set oPartSupport = Nothing
    End If
    
    If TypeOf pInPort.Connectable Is IJStructProfilePart Then 'If Member or Stiffener
        Dim oElements As IJElements
        Dim pBasePort As IJPort, pOffsetPort As IJPort
        
        'OperationProgID is NULL.So, get the latest geometry
        oConnectable.GetBaseOffsetLateralPorts vbNullString, False, pBasePort, pOffsetPort, oElements
        
        Dim oStructPort As IJStructPort
        Set oStructPort = pInPort
        
        If ((oStructPort.ContextID And CTX_BASE) = CTX_BASE) Then
            Set pOutPort = pBasePort
        Else
            Set pOutPort = pOffsetPort
        End If
        
        Set oConnectable = Nothing
        Set oElements = Nothing
        Set pBasePort = Nothing
        Set pOffsetPort = Nothing
    Else
        Set pOutPort = pInPort
    End If
    
End Sub
'*******************************************************************
'   CreateSeamControlLineForPlPr
'
'   Currently the same (almost) as CreateFittingMarkLine.
'   Eventually the line should have the same direction as the seam.
'*******************************************************************

Public Function CreateSeamControlLineForPlPr(oProfile As Object, oPos As IJDPosition, _
                                      oConnectionData As ConnectionData, UpSide As Long, _
                                      oDirToOffset As IJDVector, OffsetAmount As Double) As IJWireBody
    Const METHOD As String = "CreateSeamControlLineForPlPr"
    On Error GoTo ErrorHandler

    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfile
    
    Dim oUpsidePort As IJPort
    Set oUpsidePort = oProfileWrapper.GetSurfacePort(UpSide)
    If oUpsidePort Is Nothing Then Exit Function
    
    Dim oUpsideSurface As IJSurface ' Surface to put final mark
    Set oUpsideSurface = oUpsidePort.Geometry

    'Since the port dosent have global port so get Correct global port
    Dim oPort As IJPort
    Set oPort = oConnectionData.ConnectingPort
    
    Dim oCorrectPort As IJPort
    GetCorrectPort oPort, oCorrectPort
           
    Dim oPortSurface As IJSurfaceBody
    Set oPortSurface = oCorrectPort.Geometry
    
    Dim oWire As IJWireBody
    Set oWire = m_oMfgRuleHelper.GetCommonGeometry(oPortSurface, oUpsideSurface)
        
    Dim oFinalWire As IJWireBody
    Set oFinalWire = m_oMfgRuleHelper.OffsetCurve(oUpsideSurface, oWire, oDirToOffset, OffsetAmount, False)
    
    Set CreateSeamControlLineForPlPr = oFinalWire

CleanUp:
    Set oProfileWrapper = Nothing
    Set oUpsidePort = Nothing
    Set oUpsideSurface = Nothing
    Set oPort = Nothing
    Set oCorrectPort = Nothing
    Set oPortSurface = Nothing
    Set oWire = Nothing
    Set oFinalWire = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo CleanUp
End Function



' ***********************************************************************************
' CreateSeamControlMark3D
'
' Description:  This function creates the Seam control marks based on given inputs.
'
' ***********************************************************************************
Public Function CreateSeamControlMark3D(ByVal Part As Object, ByVal UpSide As Long, _
                                        ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                        Optional ByVal bProcessAllBUTTConns As Boolean) As IJMfgGeomCol3d
                                        
    Const METHOD = "CreateSeamControlMark3D"
    On Error GoTo ErrorHandler
    
    'Create the SD profile Wrapper and initialize it
    Dim oSDProfileWrapper As Object
    'Create the SD profile Wrapper and initialize it
    If TypeOf Part Is IJStiffenerPart Then
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = Part
    ElseIf TypeOf Part Is IJBeamPart Then
        Set oSDProfileWrapper = New StructDetailObjects.BeamPart
        Set oSDProfileWrapper.object = Part
    End If
    
    Dim oSCProfileHelper As New StructDetailObjects.Helper

    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oWBLandingCurve As IJWireBody
    Set oWBLandingCurve = oProfileWrapper.GetLandingCurve
    
    'Get the physical connected objects
    Dim oConObjsCol As Collection
    'Set oConObjsCol = oSDProfileWrapper.ConnectedObjects
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    'CreateAPSProfileMarkings STRMFG_SEAM_MARK, ReferenceObjColl, oGeomCol3d
    Set CreateSeamControlMark3D = oGeomCol3d

    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can skip this function
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long
    
    ' For some reason connected objects return the same object twice on profile endconnections.
    ' Maybe a web/flange related problem !!
    ' Make sure we don't draw the same SeamControlMark twice
    Dim bBaseFound As Boolean, bOffsetFound As Boolean, bIsConnWithPlate As Boolean
    bBaseFound = False
    bOffsetFound = False
    
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
                
        'Check if the connected object is a profile part or bProcessAllBUTTConns = True else goto next item
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Or bProcessAllBUTTConns = True Then
        
            bIsConnWithPlate = False
            'If bAllBUTTConns = True then process Plate-Profile connections also
            If bProcessAllBUTTConns = True Then
                If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
                    bIsConnWithPlate = True
                End If
            End If
            
            ' Check if it is a butt-to-butt connection (both are connected through base or offset ports)
            Dim res1 As Integer, res2 As Integer, res3 As Integer, res4 As Integer
            Dim port1Flag As IMSStructConnection.eUSER_CTX_FLAGS
            Dim port2Flag As IMSStructConnection.eUSER_CTX_FLAGS

            Dim oStructPort1 As IJStructPort, oStructPort2 As IJStructPort
            Set oStructPort1 = oConnectionData.ConnectingPort
            Set oStructPort2 = oConnectionData.ToConnectedPort

            port1Flag = oStructPort1.ContextID
            port2Flag = oStructPort2.ContextID

            res1 = port1Flag And IMSStructConnection.CTX_BASE
            res2 = port1Flag And IMSStructConnection.CTX_OFFSET
            
            If bIsConnWithPlate = True Then
                res3 = port2Flag And IMSStructConnection.CTX_LATERAL
                res4 = port2Flag And IMSStructConnection.CTX_LATERAL
            Else
                res3 = port2Flag And IMSStructConnection.CTX_OFFSET
                res4 = port2Flag And IMSStructConnection.CTX_BASE
            End If
            
            If (res1 > 0 Or res2 > 0) And (res3 > 0 Or res4 > 0) Then
                ' Check for doublettes
                If (res1 > 0) Then  'Base
                    If bBaseFound Then
                        GoTo NextItem
                    Else
                        bBaseFound = True
                    End If
                End If
                
                If (res2 > 0) Then 'Offset
                    If bOffsetFound Then
                        GoTo NextItem
                    Else
                        bOffsetFound = True
                    End If
                End If
                
                Dim oStartPos As IJDPosition, oEndPos As IJDPosition, oSeamPos As IJDPosition
                Dim oDirVector As IJDVector
                
                oWBLandingCurve.GetEndPoints oStartPos, oEndPos
                If (res1 > 0) Then
                    Set oSeamPos = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oWBLandingCurve, oStartPos, GetSeamDistance, oEndPos)
                    Set oDirVector = oSeamPos.Subtract(oStartPos)
                Else
                    Set oSeamPos = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oWBLandingCurve, oEndPos, GetSeamDistance, oStartPos)
                    Set oDirVector = oSeamPos.Subtract(oEndPos)
                End If
                
                Dim oFinalWireBody As IJWireBody
                Set oFinalWireBody = CreateSeamControlLineForPlPr(Part, oSeamPos, oConnectionData, UpSide, oDirVector, 0.01)
                
                If Not oFinalWireBody Is Nothing Then
                    
                    Dim oCS As IJComplexString
                    Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oFinalWireBody)
                    
                    If Not oCS Is Nothing Then
                        ' Create in systemmark object and add to output-collection
                        Dim oSystemMark As IJMfgSystemMark
                        Dim oMarkingInfo As MarkingInfo
                        Dim oGeom3d As IJMfgGeom3d
                        Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
        
                        'Create a SystemMark object to store additional information
                        Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                        'Set the marking side
                        oSystemMark.SetMarkingSide UpSide '(should be connecting side)
        
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
        
                        oMarkingInfo.name = "SEAM CONTROL"
        
                        oSystemMark.Set3dGeometry oGeom3d
                        oGeom3d.PutGeometry oCS
                        oGeom3d.PutGeometrytype STRMFG_SEAM_MARK
                        oGeom3d.FaceId = UpSide
        
                        oGeom3d.PutMoniker m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                        'ApplyDirection oSystemMark, oConnectionData.ConnectingPort
                        oGeomCol3d.AddGeometry 1, oGeom3d
                    End If
                End If
            End If
        End If
NextItem:
        Set oStructPort1 = Nothing
        Set oStructPort2 = Nothing
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oSeamPos = Nothing
        Set oCS = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
   Next nIndex

    'Return the 3d collection
    Set CreateSeamControlMark3D = oGeomCol3d
    
CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oWBLandingCurve = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function CreateFrameMarks(oGeomCol3d As IJMfgGeomCol3d, eJXSEC_CODE As JXSEC_CODE, sAxis As String, oSDMemberWrapper As Object, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias)
    Const METHOD = ":CreateFrameMarks"
    On Error GoTo ErrorHandler

    Dim eGeomType As StrMfgGeometryType
    If sAxis = "X" Then
        eGeomType = STRMFG_FRAMELINE_MARK
    ElseIf sAxis = "Y" Then
        eGeomType = STRMFG_BUTTOCKLINE_MARK
    ElseIf sAxis = "Z" Then
        eGeomType = STRMFG_WATERLINE_MARK
    End If
    
    
     Dim oResourceManager As IUnknown
     Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)


         Dim oCSColl As IJElements
         Dim oMfgGeomHelper As IJMfgGeomHelper
         Set oMfgGeomHelper = New MfgGeomHelper
        
         'Get the Coordinate System object
         Dim oMfgFrameSystem As IJDCoordinateSystem
         Dim nIndex As Long
         For nIndex = 1 To ReferenceObjColl.Count
             If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
                 Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
             End If
         Next nIndex
         
         ' Within V7, the Mfg object inherit the Frame System. When not found we do
         ' leave by the default rule. Customers can activate the Main CS.
         ' TR-CP·121466  manufacturing Plate profile being marked even though Frame system not defined
         If oMfgFrameSystem Is Nothing Then
             Exit Function
         End If
        
         Dim oFrameSysMoniker As IMoniker
         Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oMfgFrameSystem)
        
         Dim oMfgHFrameSys As IHFrameSystem
         Set oMfgHFrameSys = oMfgFrameSystem
        
         'With radial CS being supported, a call to the same function with "R" as axis parameter
         'is made and the obtained planes, curves are added to collection and processed further.
         '   "R" - stands for Radial Planes in Cylindrical CS
         
         Dim oConObjsCol As Collection
         Dim oRefplanes As Collection
         oSDMemberWrapper.ReferencePlaneRange oMfgHFrameSys.name, sAxis, oRefplanes, oConObjsCol, eJXSEC_CODE
         
         'Below line needs to uncommented for the system to demonstrate the capability to
         'support Radial Co-ordinate System.
         'oSDMemberWrapper.ReferencePlaneRange oMfgHFrameSys.Name, "R", oRefplanes, oConObjsCol, JXSEC_WEB_LEFT
         
         If Not (oConObjsCol Is Nothing) Then
             
             On Error Resume Next
             For nIndex = 1 To oConObjsCol.Count
             
                 Dim oWB             As IJWireBody
                 Dim oCS             As IJComplexString
                 Dim oSystemMark     As IJMfgSystemMark
                 Dim oMarkingInfo    As MarkingInfo
                 Dim oGeom3d         As IJMfgGeom3d
         
                 Set oWB = oConObjsCol.Item(nIndex)
         
             '   Convert the IJWireBody to a IJComplexString
                 Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWB)
                 
                 If Not oCSColl Is Nothing Then
                     If oCSColl.Count = 0 Then
                         Set oCSColl = Nothing
                     End If
                 End If
                     
                 If (oCSColl Is Nothing) Then
                     GoTo NextItem
                 End If
                 
                 For Each oCS In oCSColl
                 '   Create a SystemMark object to store additional information
                     Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
         
                 '   Set the marking side
                     oSystemMark.SetMarkingSide eJXSEC_CODE
         
                 '   QI for the MarkingInfo object on the SystemMark
                     Set oMarkingInfo = oSystemMark
         
                 '   Get the Frame Name
                     oMarkingInfo.name = oRefplanes(nIndex)
         
                     Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                     oGeom3d.PutGeometry oCS
                     oGeom3d.PutGeometrytype eGeomType
                     oGeom3d.FaceId = eJXSEC_CODE
         
                     oGeom3d.PutMoniker oFrameSysMoniker
                     oSystemMark.Set3dGeometry oGeom3d
                 
                     oGeomCol3d.AddGeometry 1, oGeom3d
                     
                     Set oSystemMark = Nothing
                     Set oGeom3d = Nothing
                     Set oMarkingInfo = Nothing
                     Set oCS = Nothing
                     
                 Next oCS
NextItem:
             
                 Set oWB = Nothing
                 Set oCS = Nothing
                 Set oSystemMark = Nothing
                 Set oGeom3d = Nothing
                 Set oMarkingInfo = Nothing
                 
             Next nIndex
             
         End If

CleanUp:
    Set oConObjsCol = Nothing
    Set oRefplanes = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2012, , "RULES")
    GoTo CleanUp
End Function

Public Function CreateFrameMarksByBlock(oGeomCol3d As IJMfgGeomCol3d, eJXSEC_CODE As JXSEC_CODE, sAxis As String, oSDMemberWrapper As Object, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias)
    Const METHOD = "RefPlanesByBlockX: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim eGeomType As StrMfgGeometryType
    Dim eAxisType As AxisType
    If sAxis = "X" Then
        eGeomType = STRMFG_FRAMELINE_MARK
        eAxisType = x
    ElseIf sAxis = "Y" Then
        eGeomType = STRMFG_BUTTOCKLINE_MARK
        eAxisType = y
    ElseIf sAxis = "Z" Then
        eGeomType = STRMFG_WATERLINE_MARK
        eAxisType = z
    End If
    
    Dim oCSColl As IJElements
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    'Get the web-left Surface
    Dim oWebLeftPort            As IJPort
    Dim oMemberSurfGeom         As Object
    
    Set oWebLeftPort = oSDMemberWrapper.SubPort(eJXSEC_CODE)
    Set oMemberSurfGeom = oWebLeftPort.Geometry

    Dim nIndex                  As Long
    Dim oConObjsCol             As Collection
    Dim oRefplanes              As Collection

     ' LEAD_BLOCK_REFERENCE  -- Gets the frames only from the immediate parent block
    m_oMfgRuleHelper.GetBlockFrames oSDMemberWrapper.object, oMemberSurfGeom, eAxisType, LEAD_BLOCK_REFERENCE, oRefplanes, oConObjsCol
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If Not (oConObjsCol Is Nothing) Then
    
        Dim oSPGMiddleHelper    As New SP3dGSMiddleHelper.SPGMiddleHelper
        
        On Error Resume Next
        For nIndex = 1 To oConObjsCol.Count
        
            Dim oWB                 As IJWireBody
            Dim oCS                 As IJComplexString
            Dim oSystemMark         As IJMfgSystemMark
            Dim oMarkingInfo        As MarkingInfo
            Dim oGeom3d             As IJMfgGeom3d
            Dim oNamedItem          As IJNamedItem
            Dim oFrameSystem        As IJDCoordinateSystem
            Dim oFrameSysMoniker    As IMoniker
    
            Set oWB = oConObjsCol.Item(nIndex)
    
        '   Convert the IJWireBody to a IJComplexString
            Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWB)
            
            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
                
            If (oCSColl Is Nothing) Then
                GoTo NextItem
            End If
            
            For Each oCS In oCSColl
            '   Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
            '   Set the marking side
                oSystemMark.SetMarkingSide eJXSEC_CODE
    
            '   QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
            '   Get the Frame Name
                Set oNamedItem = oRefplanes(nIndex)
                oMarkingInfo.name = oNamedItem.name
    
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype eGeomType
                oGeom3d.FaceId = eJXSEC_CODE
                
            '   Get the Frame System
                oSPGMiddleHelper.GetCoordinateSystemOfPlane oRefplanes(nIndex), oFrameSystem
                
            '   Get the Frame System Moniker and store on the Geom3d object
                Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oFrameSystem)
                oGeom3d.PutMoniker oFrameSysMoniker
               
                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d
                
                Set oSystemMark = Nothing
                Set oGeom3d = Nothing
                Set oMarkingInfo = Nothing
                Set oCS = Nothing
                Set oNamedItem = Nothing
                
            Next oCS
NextItem:
        
            Set oWB = Nothing
            Set oCS = Nothing
            Set oSystemMark = Nothing
            Set oGeom3d = Nothing
            Set oMarkingInfo = Nothing
            
        Next nIndex
    End If

CleanUp:

    Set oConObjsCol = Nothing
    Set oRefplanes = Nothing
    Set oWebLeftPort = Nothing
    Set oMemberSurfGeom = Nothing
    Set oFrameSystem = Nothing
    Set oSPGMiddleHelper = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2012, , "RULES")
    GoTo CleanUp
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : The purpose of this routine is to calculate the intersection between any two given
'            objects. The call will be delegated to the G&T PlaceIntersectionObject routine
'--------------------------------------------------------------------------------------------------
Public Function GetIntersection(pIntersectedObject As Object, pIntersectingObject As Object) As Object
On Error GoTo ErrorHandler
Const METHOD = "GetIntersection"

    ' Find the intersection.
    Dim oGeometryIntersector    As IMSModelGeomOps.DGeomOpsIntersect
    Set oGeometryIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    On Error Resume Next 'Needed for continuing with next skid mark if intersection fails
    Dim oIntersectionUnknown    As IUnknown        ' Resultant intersection.
    oGeometryIntersector.PlaceIntersectionObject Nothing, pIntersectedObject, pIntersectingObject, Nothing, oIntersectionUnknown
    
    On Error GoTo ErrorHandler
    Set GetIntersection = oIntersectionUnknown
    Set oGeometryIntersector = Nothing
    Set oIntersectionUnknown = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************************************
' Public Sub CreateMemberHoleMarks
'
' Description:  Helper function to create hole reference marks for the input part.
'               Input arguments: this part, Wire Body of the feature/hole(2D), marking side.
'
' ***********************************************************************************
Public Sub CreateMemberHoleMarks(oThisPart As Object, oFeaturePort As IJPort, _
                                            oFaceId As Long, _
                                            oCurve As IJWireBody, _
                                            UpSide As Long, _
                                            oFeature As IUnknown, _
                                            bMarkOnBothFlanges As Boolean, _
                                            oGeom3dCustom As IJMfgGeom3d, _
                                            ByRef oGeomCol3d As MfgGeomCol3d)

    Const sMETHOD = "CreateMemberHoleMarks"
    On Error GoTo ErrorHandler
        
    'Getting the Center of Gravity
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oFeaturePort.Geometry
    
    Dim oCOG As IJDPosition
    Dim oSurfaceNormal As IJDVector
    
    oSurfaceBody.GetCenterOfGravity oCOG
    oSurfaceBody.GetNormalFromPosition oCOG, oSurfaceNormal
    
    Dim oFeatureWireBody As IJWireBody
    Set oFeatureWireBody = oCurve
        
    Dim oWBUtil As IJSGOWireBodyUtilities
    Set oWBUtil = New SGOWireBodyUtilities
    
    'Checking the Wirebody is Closed
    If oWBUtil.IsWireBodyClosed(oFeatureWireBody) = True Then
                
        Dim oCurveElems As IJElements
        
        Dim oMfgMGHelper As IJMfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper
        
        oMfgMGHelper.WireBodyToComplexStrings oFeatureWireBody, oCurveElems
        'oCurveElems.AddElements oTempElems
                
    End If
    
    'Finding the Minimum  Bounding Box for the Feature
    Dim oMfgGeomHelper As New MfgGeomHelper
    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElems, oSurfaceNormal)
                            
    Dim oPoints(1 To 4) As IJDPosition
    
    Set oPoints(1) = oBoxPoints.Item(1)
    Set oPoints(2) = oBoxPoints.Item(2)
    Set oPoints(3) = oBoxPoints.Item(3)
    
    Dim oPoints4 As IJDPosition
    Set oPoints4 = New DPosition
    
    oPoints4.x = oPoints(1).x + oPoints(3).x - oPoints(2).x
    oPoints4.y = oPoints(1).y + oPoints(3).y - oPoints(2).y
    oPoints4.z = oPoints(1).z + oPoints(3).z - oPoints(2).z
                            
    Dim dLength(1 To 2) As Double
    dLength(1) = oPoints(1).DistPt(oPoints(2))
    dLength(2) = oPoints(2).DistPt(oPoints(3))
    
    If Abs(dLength(1) - dLength(2) < 0.0001) Then
        If IsALine(oCurveElems.Item(1)) = False Then

            Dim oBoxVector As IJDVector
            Set oBoxVector = oPoints(1).Subtract(oPoints(2))
            oBoxVector.length = 1

            Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
            Set oXVector = New DVector
            Set oYVector = New DVector
            Set oZVector = New DVector

            oXVector.Set 1, 0, 0
            oYVector.Set 0, 1, 0
            oZVector.Set 0, 0, 1

            If Not (Abs(oBoxVector.Dot(oXVector)) > 0.999 Or Abs(oBoxVector.Dot(oYVector)) > 0.999 Or Abs(oBoxVector.Dot(oZVector)) > 0.999) Then
                Dim oVecElems As IJElements
                Set oVecElems = New JObjectCollection
                
                Dim dXDotP As Double, dYDotP As Double, dZDotP As Double
                
                ' Take absolute value of dot product as surface normal can be either positive or negative
                dXDotP = Abs(oSurfaceNormal.Dot(oXVector))
                dYDotP = Abs(oSurfaceNormal.Dot(oYVector))
                dZDotP = Abs(oSurfaceNormal.Dot(oZVector))
                
                'The order in which vectors are added to the collection is important.
                ' The vector aligning with the surface normal need to be added at the end of collection
                If ((dXDotP > dYDotP) And (dXDotP > dZDotP)) Then
                    oVecElems.Add oYVector
                    oVecElems.Add oZVector
                    oVecElems.Add oXVector
                ElseIf ((dYDotP > dXDotP) And (dYDotP > dZDotP)) Then
                    oVecElems.Add oXVector
                    oVecElems.Add oZVector
                    oVecElems.Add oYVector
                Else
                    oVecElems.Add oXVector
                    oVecElems.Add oYVector
                    oVecElems.Add oZVector
                End If

                Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVectors(oCurveElems, oVecElems)

                Set oPoints(1) = oBoxPoints.Item(1)
                Set oPoints(2) = oBoxPoints.Item(2)
                Set oPoints(3) = oBoxPoints.Item(3)

                oPoints4.x = oPoints(1).x + oPoints(3).x - oPoints(2).x
                oPoints4.y = oPoints(1).y + oPoints(3).y - oPoints(2).y
                oPoints4.z = oPoints(1).z + oPoints(3).z - oPoints(2).z
            End If
       End If
    End If
        
    'Calculating the Cross Points
    Dim oCrossPoint1 As IJDPosition
    Set oCrossPoint1 = New DPosition
    
    oCrossPoint1.x = (oPoints(1).x + oPoints(2).x) / 2
    oCrossPoint1.y = (oPoints(1).y + oPoints(2).y) / 2
    oCrossPoint1.z = (oPoints(1).z + oPoints(2).z) / 2
    
    Dim oCrossPoint2 As IJDPosition
    Set oCrossPoint2 = New DPosition
    
    oCrossPoint2.x = (oPoints(2).x + oPoints(3).x) / 2
    oCrossPoint2.y = (oPoints(2).y + oPoints(3).y) / 2
    oCrossPoint2.z = (oPoints(2).z + oPoints(3).z) / 2
    
    Dim oCrossPoint3 As IJDPosition
    Set oCrossPoint3 = New DPosition
    
    oCrossPoint3.x = (oPoints(3).x + oPoints4.x) / 2
    oCrossPoint3.y = (oPoints(3).y + oPoints4.y) / 2
    oCrossPoint3.z = (oPoints(3).z + oPoints4.z) / 2
    
    Dim oCrossPoint4 As IJDPosition
    Set oCrossPoint4 = New DPosition
    
    oCrossPoint4.x = (oPoints(1).x + oPoints4.x) / 2
    oCrossPoint4.y = (oPoints(1).y + oPoints4.y) / 2
    oCrossPoint4.z = (oPoints(1).z + oPoints4.z) / 2
    
    'Constructing the Cross Lines
    Dim oCrossLine1 As IJLine
    Dim oCrossLine2 As IJLine
    Set oCrossLine1 = New Line3d
    Set oCrossLine2 = New Line3d
    
    oCrossLine1.DefineBy2Points oCrossPoint1.x, oCrossPoint1.y, oCrossPoint1.z, oCrossPoint3.x, oCrossPoint3.y, oCrossPoint3.z
    oCrossLine2.DefineBy2Points oCrossPoint2.x, oCrossPoint2.y, oCrossPoint2.z, oCrossPoint4.x, oCrossPoint4.y, oCrossPoint4.z

    Dim oCS1 As IJComplexString
    Dim oCS2 As IJComplexString
    Set oCS1 = New ComplexString3d
    Set oCS2 = New ComplexString3d
    
    oCS1.AddCurve oCrossLine1, True
    oCS2.AddCurve oCrossLine2, True
                                                                                           
    Dim oWireBody1 As IJWireBody
    Set oWireBody1 = m_oMfgRuleHelper.ComplexStringToWireBody(oCS1)
    
    Dim oWireBody2 As IJWireBody
    Set oWireBody2 = m_oMfgRuleHelper.ComplexStringToWireBody(oCS2)
                                                                             
    On Error Resume Next
    Set oCS1 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody1, Nothing)
    Err.Clear
    Set oCS2 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody2, Nothing)
    Err.Clear
    On Error GoTo ErrorHandler
    
    If oCS1 Is Nothing Or oCS2 Is Nothing Then Exit Sub
                                                                                                                                       
    Dim oMfgGeomUtilwrapper1 As New MfgGeomUtilWrapper
    oMfgGeomUtilwrapper1.ExtendWire oCS1, 0.05
    Dim oMfgGeomUtilwrapper2 As New MfgGeomUtilWrapper
    oMfgGeomUtilwrapper2.ExtendWire oCS2, 0.05
    
        
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
       
    'Creating the Marks
    Dim jCount As Integer
    For jCount = 1 To 2
        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        'Set the marking side
        oSystemMark.SetMarkingSide UpSide
        
        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark
                                                                                          
        'Checking whether the Mark is required on Both Flanges or Not
        If bMarkOnBothFlanges = True Then
            Dim kCount As Integer
            For kCount = 1 To 2
               Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            
               If jCount = 1 Then
                   oGeom3d.PutGeometry oCS1
               Else
                   oGeom3d.PutGeometry oCS2
               End If
                            
               oGeom3d.PutGeometrytype STRMFG_HOLE_REF_MARK 'Hole Ref mark
               
               If kCount = 1 Then
                    oGeom3d.FaceId = oFaceId
                Else
                    If oFaceId = JXSEC_TOP Then
                        oGeom3d.FaceId = JXSEC_BOTTOM
                    Else
                        oGeom3d.FaceId = JXSEC_TOP
                    End If
                End If
                              
               oGeom3d.TrimToBoundaries = True
                
               oGeomCol3d.AddGeometry 1, oGeom3d
               oSystemMark.Set3dGeometry oGeom3d
                
               If oGeom3dCustom Is Nothing Then
                   Set oMoniker = m_oMfgRuleHelper.GetMoniker(oFeature)
               Else
                   Set oMoniker = oGeom3dCustom.GetMoniker
               End If
                
               oGeom3d.PutMoniker oMoniker
            Next kCount
        Else
            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
         
            If jCount = 1 Then
                oGeom3d.PutGeometry oCS1
            Else
                oGeom3d.PutGeometry oCS2
            End If
                         
            oGeom3d.PutGeometrytype STRMFG_HOLE_REF_MARK 'Hole Ref mark
            oGeom3d.FaceId = oFaceId
            oGeom3d.TrimToBoundaries = True
             
            oGeomCol3d.AddGeometry 1, oGeom3d
            oSystemMark.Set3dGeometry oGeom3d
             
            If oGeom3dCustom Is Nothing Then
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oFeature)
            Else
                Set oMoniker = oGeom3dCustom.GetMoniker
            End If
             
            oGeom3d.PutMoniker oMoniker
        End If
                       
        
    Next jCount
    
    Set oSystemMark = Nothing
    Set oMoniker = Nothing
    Set oMarkingInfo = Nothing
    Set oGeom3d = Nothing

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ***********************************************************************************
' Method:
'
' Description:This Function is written to check whether given geometry is curved or linear
' ***********************************************************************************
Public Function IsALine(oCS As Object) As Boolean
Const METHOD = "IsALine"
On Error GoTo ErrorHandler
    
    Dim oCurveElems As IJElements
    Set oCurveElems = New JObjectCollection

    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper

    oCurveElems.Add oCS

    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElems)

    Dim Length1 As Double, length2 As Double, length3 As Double
    Dim Points(1 To 4) As IJDPosition

    Set Points(1) = oBoxPoints.Item(1)
    Set Points(2) = oBoxPoints.Item(2)
    Set Points(3) = oBoxPoints.Item(3)
    Set Points(4) = oBoxPoints.Item(4)

    Length1 = Points(1).DistPt(Points(2))
    length2 = Points(2).DistPt(Points(3))
    length3 = Points(1).DistPt(Points(4))
   
    If (Length1 < 0.01 And length2 < 0.01) Or (length2 < 0.01 And length3 < 0.01) Or (Length1 < 0.01 And length3 < 0.01) Then
        IsALine = True
    End If
    
CleanUp:
    Set Points(1) = Nothing
    Set Points(2) = Nothing
    Set Points(3) = Nothing
    Set Points(4) = Nothing
    Set oBoxPoints = Nothing
    Set oCurveElems = Nothing
    Set oMfgGeomHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
