Attribute VB_Name = "PrMrkHelpers"
'*******************************************************************
'  Copyright (C) 2002 Global Research and Development, Inc.  All rights reserved.
'
'  Project: MfgProfieMarking
'
'  Abstract:    Helpers for the MfgPlateMarking rules.
'
'  History:
'       TBH         april 8. 2002   created
'       KONI        may. 17 2002    implementation
'       Venkat      Oct. 11 2006    Review MsgBox usage-TR:103745
'       Koushik     Sep. 09 2008    TR-CP-148665    ErrorLog message should not have been localized
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

Public Function CreateFittingMarkLine(oProfile As IJProfilePart, oPos As IJDPosition, oConnectionData As ConnectionData, UpSide As Long) As IJWireBody
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
        Set oPlateBasePort = oMfgPlateWrapper.GetSurfacePort(BaseSide)
        
        Set oPlateSurface = oPlateBasePort.Geometry
    Else
        Set oPlateSurface = oPlatePort.Geometry
    End If

    Set pModelBody = oPlateSurface
    
    Dim oGeomOps As IJSGOModelBodyUtilities
    Set oGeomOps = New SGOModelBodyUtilities
    oGeomOps.GetClosestPointOnBody pModelBody, oPos, oNewPos, dDistance

    Dim oDir As IJDVector, oLineDir As IJDVector
    oPlateSurface.GetNormalFromPosition oNewPos, oDir
        
    'Create temporary line
    Dim oLine As IJLine
    Set oLine = New Line3d
    endX = oPos.x + (oDir.x * FittingMarkLength)
    endY = oPos.y + (oDir.y * FittingMarkLength)
    endZ = oPos.z + (oDir.z * FittingMarkLength)
    oLine.DefineBy2Points oPos.x, oPos.y, oPos.z, endX, endY, endZ
    
    Set oLineDir = oDir
    
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
    
    If Abs(oLineDir.Dot(oDir)) = 1# Then
        Exit Function
    End If
    
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

Public Function CreateSeamControlLine(oProfile As IJProfilePart, oPos As IJDPosition, _
                                      oConnectionData As ConnectionData, UpSide As Long, _
                                      oDirToOffset As IJDVector, OffsetAmount As Double) As IJWireBody
    Const METHOD As String = "CreateSeamControlLine"
    On Error GoTo ErrorHandler

    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfile
    
    Dim oUpsidePort As IJPort
    Set oUpsidePort = oProfileWrapper.GetSurfacePort(UpSide)
    If oUpsidePort Is Nothing Then Exit Function
    
    Dim oUpsideSurface As IJSurface ' Surface to put final mark
    Dim oUpsideSurfaceObj As Object
    Set oUpsideSurfaceObj = oUpsidePort.Geometry
    
    If Not TypeOf oUpsideSurfaceObj Is IJSurface Then Exit Function
    Set oUpsideSurface = oUpsideSurfaceObj

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
    
    Set CreateSeamControlLine = oFinalWire

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

Public Function CreateGeom3dObject(oCurve As IJCurve, eType As StrMfgGeometryType, lUpside As Long, oMarkingInfo As IJMarkingInfo) As IJMfgGeom3d
Const METHOD = "CreateGeom3dObject"
On Error GoTo ErrorHandler
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)

    oGeom3d.PutGeometry oCurve
    oGeom3d.PutGeometrytype eType
    oGeom3d.FaceId = lUpside

    'Create a SystemMark object to store additional information
    Dim oSystemMark As IJMfgSystemMark
    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
    'Set the marking side
    oSystemMark.SetMarkingSide lUpside
    oSystemMark.Set3dGeometry oGeom3d
    Set oMarkingInfo = oSystemMark

    Set CreateGeom3dObject = oGeom3d
    
    Set oGeom3d = Nothing
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

Private Function GetDirAsFAUDIO(ByVal oVec As IJDVector, ByVal oPt As IJDPosition) As String
    Const METHOD = "GetDirAsFAUDIO"
    On Error GoTo ErrorHandler

    If oVec Is Nothing Then Exit Function
    If oVec.length < 0.000001 Then Exit Function ' Null Vector has no direction!

    If Abs(oVec.x) > Abs(oVec.y) Then
        If Abs(oVec.x) > Abs(oVec.z) Then
            If oVec.x > 0 Then
                GetDirAsFAUDIO = "F"
            Else
                GetDirAsFAUDIO = "A"
            End If
        Else
            If oVec.z > 0 Then
                GetDirAsFAUDIO = "U"
            Else
                GetDirAsFAUDIO = "D"
            End If
        End If
    Else
        If Abs(oVec.z) > Abs(oVec.y) Then
            If oVec.z > 0 Then
                GetDirAsFAUDIO = "U"
            Else
                GetDirAsFAUDIO = "D"
            End If
        Else
            If oPt Is Nothing Then Exit Function

            If oVec.y > 0 And oPt.y >= 0 Or oVec.y < 0 And oPt.y <= 0 Then
                GetDirAsFAUDIO = "O"
            ElseIf oVec.y > 0 And oPt.y < 0 Or oVec.y < 0 And oPt.y > 0 Then
                GetDirAsFAUDIO = "I"
            End If
        End If
    End If

CleanUp:

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Sub RotateCSaboutPointNormal(oCS As IJComplexString, _
                                     oPivot As IJDPosition, _
                                     oNormal As IJDVector, _
                                     dAngle As Double)

    Const METHOD = "RotateCSaboutPointNormal"
    On Error GoTo ErrorHandler

    Dim oRotMat As New DT4x4
    oRotMat.LoadIdentity

    Dim oPivotPointAsVec As New DVector
    oPivotPointAsVec.Set oPivot.x, oPivot.y, oPivot.z
    oRotMat.Translate oPivotPointAsVec

    oRotMat.Rotate dAngle, oNormal

    oPivotPointAsVec.Set -oPivot.x, -oPivot.y, -oPivot.z
    oRotMat.Translate oPivotPointAsVec
    
    Dim oTransformCS As IJTransform
    Set oTransformCS = oCS

    oTransformCS.Transform oRotMat

    Set oPivotPointAsVec = Nothing
    Set oTransformCS = Nothing
    Set oRotMat = Nothing

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub
Public Function CreateDeclivityMarksOnProfile(Part As Object, oConnectionData As ConnectionData, _
                                              oConnGeomCS As IJComplexString, lFaceId As Long, _
                                              oRefGeom3D As IJMfgGeom3d, _
                                              oGeomCol3d As IJMfgGeomCol3d, _
                                              lGeomCount As Long, sMoldedSide As String, oRelatedPart As Object)

    Const METHOD = "CreateDeclivityMarksOnProfile"
    
    If Part Is Nothing Then Exit Function
    
    ' Check the connected part and initialize some helpers
    Dim oSDObjectWrapper As Object
    Dim oProfilePart As IJProfilePart
    Dim oPlatePart As IJPlatePart
    
    Dim eSectionType As ProfileSectionType
    eSectionType = UnknownProfile
    Dim oConnectingPort As IJPort
    
    If oRelatedPart Is Nothing Then
        ' Get the connecting face port (for THIS profile part) and its surface geometry
        
        Set oConnectingPort = oConnectionData.ConnectingPort
    
        If Not Part Is oConnectingPort.Connectable Then Exit Function
        
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            ' If the connected part is a profile, get the Section type
            Set oProfilePart = oConnectionData.ToConnectable
            
            Set oSDObjectWrapper = New StructDetailObjects.ProfilePart
            Set oSDObjectWrapper.object = oConnectionData.ToConnectable
            
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
            Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport.Part = oProfilePart
            
            Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
            Set oProfilePartSupport = oPartSupport
    
            eSectionType = oProfilePartSupport.SectionType
            
            Set oPartSupport = Nothing
            Set oProfilePartSupport = Nothing
        ElseIf TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
            Set oPlatePart = oConnectionData.ToConnectable
            Set oSDObjectWrapper = New StructDetailObjects.PlatePart
            Set oSDObjectWrapper.object = oConnectionData.ToConnectable
        Else
            Exit Function
        End If
    
    Else ' In case of APS Marks
        
        Dim oMarkingLineAE      As IJMfgMarkingLines_AE
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oRefGeom3D)
        
        'get the port on which marking line is placed
        Dim oProfilePort As IJPort
        Set oProfilePort = oMarkingLineAE.GetMfgMarkingPortForProfilePart(oMarkingLineAE.GetMfgMarkingPart)
        
        Set oConnectingPort = oProfilePort
        
        If TypeOf oRelatedPart Is IJProfilePart Then
            ' If the connected part is a profile, get the Section type
            Set oProfilePart = oRelatedPart
            
            Set oSDObjectWrapper = New StructDetailObjects.ProfilePart
            Set oSDObjectWrapper.object = oRelatedPart
            
            Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport.Part = oProfilePart
            
            Set oProfilePartSupport = oPartSupport
    
            eSectionType = oProfilePartSupport.SectionType
            
            Dim eTSide As GSCADSDPartSupport.ThicknessSide
            eTSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
            
            If eTSide = SideA Then
                sMoldedSide = "Base"
            ElseIf eTSide = SideB Then
                sMoldedSide = "Offset"
            End If
            
            Set oPartSupport = Nothing
            Set oProfilePartSupport = Nothing
        
        ElseIf TypeOf oRelatedPart Is IJPlatePart Then
        
            Set oPlatePart = oRelatedPart
            Set oSDObjectWrapper = New StructDetailObjects.PlatePart
            Set oSDObjectWrapper.object = oRelatedPart
            
            sMoldedSide = oSDObjectWrapper.MoldedSide
        
        End If
    
    End If

    ' Convert the wirebody to complex string
    
    Dim oConnGeomWire As IJWireBody
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    Set oConnGeomWire = oMfgRuleHelper.ComplexStringToWireBody(oConnGeomCS)

    Dim oMidPos As IJDPosition
    Dim oToolBox As New DGeomOpsToolBox
    oToolBox.GetMiddlePointOfCompositeCurve oConnGeomWire, oMidPos

    ' Calculate declivity angle at position
    
    Dim oConnectingPortSurface As IJSurfaceBody
    Set oConnectingPortSurface = oConnectingPort.Geometry

    ' Get THIS profile's normal at the input frame position
    Dim oProjPos As IJDPosition
    Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oMidPos, oConnectingPortSurface, Nothing)
    
    Dim oMyProfileNormal As IJDVector
    oConnectingPortSurface.GetNormalFromPosition oProjPos, oMyProfileNormal

    Dim oConnectedObjNormal As IJDVector
    Dim oConnectedObjPlaneVec As IJDVector
    
    If Not oProfilePart Is Nothing Then
        ' Get the profile cross section matrix at the mid PC geometry position
        Dim oCrossSectionMatrix As AutoMath.DT4x4
        Dim oTopologyLocate As New TopologyLocate
        Set oCrossSectionMatrix = oTopologyLocate.GetPenetratingCrossSectionMatrix(oProfilePart, oMidPos)
        
        Dim dMat() As Double
        dMat = oCrossSectionMatrix.GetMatrix
        
        ' oConnectedObjNormal is the normal on WebRight surface
        Set oConnectedObjNormal = New DVector
        oConnectedObjNormal.Set dMat(0), dMat(1), dMat(2)
        
        ' Get the plane on which the angle is measured
        Set oConnectedObjPlaneVec = New DVector
        oConnectedObjPlaneVec.Set dMat(8), dMat(9), dMat(10)
        
        Set oCrossSectionMatrix = Nothing
    Else
        Dim oPartSupp As IJPartSupport
        Set oPartSupp = New PlatePartSupport
        
        Dim oPlatePartSupp As IJPlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = oPlatePart
        
        Dim oMarkedSurfaceBody As IJSurfaceBody
        ' Get the normal of the connected surface at the input position
        If sMoldedSide = "Base" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oMarkedSurfaceBody
        Else
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oMarkedSurfaceBody
        End If

        Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oMidPos, oMarkedSurfaceBody, Nothing)
        oMarkedSurfaceBody.GetNormalFromPosition oProjPos, oConnectedObjNormal
        Set oConnectedObjPlaneVec = oConnectedObjNormal.Cross(oMyProfileNormal)

        Set oMarkedSurfaceBody = Nothing
        Set oPartSupp = Nothing
        Set oPlatePartSupp = Nothing
    End If

    Const PI As Double = 3.14159265358979

    ' Calculate the declivity angle
    Dim dDeclivityAngle As Double
    dDeclivityAngle = oConnectedObjNormal.Angle(oMyProfileNormal, oConnectedObjPlaneVec)

    ' If the angle is exactly 90 degrees ( 0.5 degree tolerance ),
    ' There is no need to create the mark. In this case, just goto the next position
    If Abs(Cos(dDeclivityAngle)) < Sin(DECLIVITY_SHOW_TOLERANCE * PI / 180#) Then
        GoTo CleanUp
    End If

    If dDeclivityAngle > PI Then
        dDeclivityAngle = 2 * PI - dDeclivityAngle
    End If

    Dim bDeclivityIsOnBase As Boolean
    If eSectionType = Angle Or eSectionType = Bulb_Angle Or eSectionType = Bulb_Tee Or _
       eSectionType = Bulb_Type Or eSectionType = Fab_Angle Then
        bDeclivityIsOnBase = True
    Else
    'Ref: TSN  Customization - mounting angle marked on the obtuse angle (Section 6.16)
    'Ref: SKDY Customization - mounting angle marked on the Acute side (Section 9.2.2)
        If dDeclivityAngle < PI / 2# And DECLIVITY_MARK_OBTUSE_ANGLE Or _
           dDeclivityAngle > PI / 2# And Not DECLIVITY_MARK_OBTUSE_ANGLE Then
            dDeclivityAngle = PI - dDeclivityAngle
            bDeclivityIsOnBase = False
        Else
            bDeclivityIsOnBase = True
        End If
    End If

    Dim oConnGeomCurve As IJCurve
    Set oConnGeomCurve = oConnGeomCS

    Dim dPointPar As Double
    oConnGeomCurve.Parameter oMidPos.x, oMidPos.y, oMidPos.z, dPointPar
    
    Dim dTanX As Double, dTanY As Double, dTanZ As Double, dDummy As Double
    oConnGeomCurve.Evaluate dPointPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
    
    
    Dim oDeclPosOnCS As IJDPosition
    Set oDeclPosOnCS = oMidPos
    
    Dim oMarkTangentVec As New DVector
    oMarkTangentVec.Set dTanX, dTanY, dTanZ
    oMarkTangentVec.length = DECLIVITY_MARKING_LINE_LENGTH

    Dim oThicknessDirVec As IJDVector
    Set oThicknessDirVec = GetThicknessDirectionVectorAtAGivenPos(oDeclPosOnCS, oSDObjectWrapper, oConnectingPortSurface, sMoldedSide)
    If Not oProfilePart Is Nothing Then
        If ((sMoldedSide = "Base" And bDeclivityIsOnBase = False) Or _
           (sMoldedSide = "Offset" And bDeclivityIsOnBase = True)) Then
            oThicknessDirVec.length = DECLIVITY_OFFSET_FROM_LOCATION
        Else
            oThicknessDirVec.length = -DECLIVITY_OFFSET_FROM_LOCATION
        End If
    Else
        If bDeclivityIsOnBase Then
            oThicknessDirVec.length = -DECLIVITY_OFFSET_FROM_LOCATION
        Else
            oThicknessDirVec.length = DECLIVITY_OFFSET_FROM_LOCATION
        End If
    End If

    Dim oDeclRootPos As IJDPosition
    Set oDeclRootPos = oDeclPosOnCS.Offset(oThicknessDirVec)
    
    Dim oDeclStartPos As IJDPosition
    Set oDeclStartPos = oDeclRootPos.Offset(oMarkTangentVec)


    '*** Check if Declivity Mark is intersecting outer contour or outside outer contour ***'

    ' Get the MfgProfile and 3D Contour Collection
    
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

    Dim oOutrCntGeom3dColl As IJMfgGeomCol3d
    Dim oMfgPart As IJMfgProfilePart
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oOutrCntGeom3dColl = oMfgPart.GeometriesBeforeUnfold
    Else
        'Exit Function?
    End If

    Set oProfileWrapper = Nothing
    Set oMfgPart = Nothing
    
    'Loop through all outer and inner contours to get the distance from root point
    '.. If the distance is less than (declivity mark length + offset), we will flip the direction and take complement angle.
    Dim j As Long
    For j = 1 To oOutrCntGeom3dColl.Getcount

        Dim objGeom As IJMfgGeom3d
        Set objGeom = oOutrCntGeom3dColl.GetGeometry(j)

        'Check if the contour is outer / inner
        If objGeom.GetGeometryType = STRMFG_OUTER_CONTOUR Or objGeom.GetGeometryType = STRMFG_INNER_CONTOUR Then
            Dim oTCurve As IJCurve
            Set oTCurve = objGeom.GetGeometry

            'Get the distance of root of declivity mark from the contour
            Dim dMinDist As Double
            Dim dPtX As Double, dPtY As Double, dPtZ As Double
            Dim dPtX2 As Double, dPtY2 As Double, dPtZ2 As Double
            oTCurve.DistanceBetween oDeclRootPos, dMinDist, dPtX, dPtY, dPtZ, dPtX2, dPtY2, dPtZ2

            'If distance is less than (declivity mark length + offset) then flip the direction and take complement angle.
            If dMinDist < DECLIVITY_MARKING_LINE_LENGTH + DECLIVITY_OFFSET_FROM_LOCATION Then

                Dim oPosOnContour As New DPosition
                oPosOnContour.Set dPtX, dPtY, dPtZ

                'oRootContourVec is a vector from root to the outer contour
                Dim oRootContourVec As New DVector
                Set oRootContourVec = oPosOnContour.Subtract(oDeclRootPos)

                ' Dot product > 0 means that declivity mark and contour are on same side
                '.. so flip the position
                If oRootContourVec.Dot(oThicknessDirVec) > 0 Then

                    oThicknessDirVec.length = oThicknessDirVec.length * -1#
                    dDeclivityAngle = PI - dDeclivityAngle

                    'Reset the Root Position and Start Position
                    Set oDeclRootPos = oDeclPosOnCS.Offset(oThicknessDirVec)
                    Set oDeclStartPos = oDeclRootPos.Offset(oMarkTangentVec)
                    
                    bDeclivityIsOnBase = (Not bDeclivityIsOnBase)

                    'Quit the loop and continue marking
                    GoTo Marking
                End If
            End If
        End If
    Next

Marking:

    Dim oLocalZVec As IJDVector
    Set oLocalZVec = oMarkTangentVec.Cross(oThicknessDirVec)

    Dim oRotMat As New DT4x4
    oRotMat.LoadIdentity
    oRotMat.Rotate dDeclivityAngle, oLocalZVec
    
    Dim oDeclVec As IJDVector
    Set oDeclVec = oRotMat.TransformVector(oMarkTangentVec)
    oDeclVec.length = DECLIVITY_MARKING_LINE_LENGTH
    
    Dim oDeclEndPos As IJDPosition
    Set oDeclEndPos = oDeclRootPos.Offset(oDeclVec)
    
    ' Construct a two segment geometry as declivity mark

    ' TODO: Some contract to decide if only label is required.

    Dim oLines As ILines3d
    Set oLines = New GeometryFactory

    Dim oLine1 As IJLine
    Set oLine1 = oLines.CreateBy2Points(Nothing, oDeclStartPos.x, oDeclStartPos.y, oDeclStartPos.z, _
                                        oDeclRootPos.x, oDeclRootPos.y, oDeclRootPos.z)

    Dim oLine2 As IJLine
    Set oLine2 = oLines.CreateBy2Points(Nothing, oDeclRootPos.x, oDeclRootPos.y, oDeclRootPos.z, _
                                        oDeclEndPos.x, oDeclEndPos.y, oDeclEndPos.z)

    Dim oCurveColl As IJElements
    Set oCurveColl = New JObjectCollection
    oCurveColl.Add oLine1
    oCurveColl.Add oLine2

    Dim oComplexStrings As IComplexStrings3d
    Set oComplexStrings = oLines

    Dim oLineCS As IJComplexString
    Set oLineCS = oComplexStrings.CreateByCurves(Nothing, oCurveColl)

    If DECLIVITY_POINTS_TO_LOCATION Then
        ' Rotate the Declivity mark to point towards location mark
        RotateCSaboutPointNormal oLineCS, oDeclRootPos, oLocalZVec, (PI - dDeclivityAngle) / 2#
    End If

    'Create a SystemMark object to store additional information

    Dim oSystemMarkFactory As New MfgSystemMarkFactory
    
    Dim oSystemMark As MfgSystemMark
    Set oSystemMark = oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    'QI for the MarkingInfo object on the SystemMark
    Dim oMarkingInfo As MarkingInfo
    Set oMarkingInfo = oSystemMark

'    oMarkingInfo.name = ""
'    oMarkingInfo.name = GetDirAsFAUDIO(oThicknessDirVec, oDeclRootPos) & _
'                        CStr(Round(dDeclivityAngle * 180 / PI, 1))
'    If Not bDeclivityIsOnBase Then
'        oMarkingInfo.name = oMarkingInfo.name & vbLf & "C"
'    End If
    
    oMarkingInfo.FittingAngle = dDeclivityAngle
    oMarkingInfo.direction = GetDirAsFAUDIO(oThicknessDirVec, oDeclRootPos)
    oMarkingInfo.ThicknessDirection = oThicknessDirVec
    
    oMarkingInfo.SetAttributeNameAndValue "DECLIVITY", CVar(dDeclivityAngle)
    oMarkingInfo.SetAttributeNameAndValue "MOUNTINGDIRECTION", CVar(oMarkingInfo.direction)
    If bDeclivityIsOnBase Then
        oMarkingInfo.SetAttributeNameAndValue "MEASURESIDE", CVar("base")
    Else
        oMarkingInfo.SetAttributeNameAndValue "MEASURESIDE", CVar("offset")
    End If
    
    Dim oGeom3dFactory As New MfgGeom3dFactory
    
    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oGeom3d.PutGeometry oLineCS
    oGeom3d.PutGeometrytype STRMFG_MOUNT_ANGLE_MARK
    oGeom3d.FaceId = lFaceId

    Dim oMoniker As IMoniker
    
    If oRelatedPart Is Nothing Then
        Set oMoniker = oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
    Else
        Set oMoniker = oRefGeom3D.GetMoniker
    End If
        
    oGeom3d.PutMoniker oMoniker

    oSystemMark.Set3dGeometry oGeom3d

    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
    lGeomCount = lGeomCount + 1

CleanUp:
    Set oLocalZVec = Nothing
    Set oConnectedObjNormal = Nothing
    Set oConnectedObjPlaneVec = Nothing
    Set oProjPos = Nothing
    Set oDeclStartPos = Nothing
    Set oDeclEndPos = Nothing
    Set oDeclPosOnCS = Nothing
    Set oDeclRootPos = Nothing
    Set oRotMat = Nothing
    Set oDeclVec = Nothing
    Set oCurveColl = Nothing
    Set oLine1 = Nothing
    Set oLine2 = Nothing
    Set oLineCS = Nothing
    Set oMoniker = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oGeom3d = Nothing

    Set oConnGeomCurve = Nothing
    Set oConnGeomWire = Nothing
    Set oMfgRuleHelper = Nothing
    Set oToolBox = Nothing
    Set oProfilePart = Nothing
    Set oTopologyLocate = Nothing
    Set oPlatePart = Nothing
    Set oConnectingPort = Nothing
    Set oConnectingPortSurface = Nothing
    Set oMyProfileNormal = Nothing
    Set oSystemMarkFactory = Nothing
    Set oGeom3dFactory = Nothing
    Set oSDObjectWrapper = Nothing
    Set oLines = Nothing
    Set oComplexStrings = Nothing
    Set oMarkTangentVec = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

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


' ***********************************************************************************
' Private Sub CreateProfileHoleMarks
'
' Description:  Helper function to create hole mark for the input part.
'               Input arguments: this part, Wire Body of the marking line geometry, marking side.
'
' ***********************************************************************************
Public Sub CreateProfileHoleMarks(oThisPart As Object, _
                                            oCurve As IJWireBody, _
                                            UpSide As Long, _
                                            oFeature As IUnknown, _
                                            oGeom3dCustom As IJMfgGeom3d, _
                                            ByRef oGeomCol3d As MfgGeomCol3d)
    
    Const sMETHOD = "CreateProfileHoleMarks"
    On Error GoTo ErrorHandler
        
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = oThisPart
    
    Dim oProfileSupport As IJProfilePartSupport
    Set oProfileSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfileSupport
    Set oPartSupport.Part = oThisPart

    Dim oFeatureWireBody As IJWireBody
    Set oFeatureWireBody = oCurve
            
        Dim oWBUtil As IJSGOWireBodyUtilities
        Set oWBUtil = New SGOWireBodyUtilities
        
        Dim oMfgMGHelper As IJMfgMGHelper
    
        If oWBUtil.IsWireBodyClosed(oFeatureWireBody) = True Then
                                        
            Dim oCurveElems As IJElements
            Set oCurveElems = New JObjectCollection
                               
            Dim oTempElems As IJElements
            Set oMfgMGHelper = New MfgMGHelper
            
            oMfgMGHelper.WireBodyToComplexStrings oFeatureWireBody, oTempElems
            oCurveElems.AddElements oTempElems
                                    
            Dim oCOG As IJDPosition
            Dim oProjPoint As IJDPosition
            Dim oProjVector As IJDVector
                                  
            Dim oPortMoniker As IMoniker
            Dim oSubPortMonikerColl As Collection
            Dim oCountourColl As Collection
            
            Dim oWebSB As IJSurfaceBody
            Dim oSurfaceContourColl As Collection, oMonikerColl As Collection
            
            oProfileSupport.GetProfileContours WebLeftFace, oWebSB, oSurfaceContourColl, oMonikerColl
            
            oWebSB.GetCenterOfGravity oCOG
            oMfgMGHelper.ProjectPointOnSurfaceBody oWebSB, oCOG, oProjPoint, oProjVector

            Dim oWebSBNormal As IJDVector
            oWebSB.GetNormalFromPosition oProjPoint, oWebSBNormal
            
            If Not oFeature Is Nothing Then
                Dim oGenericContour As IJStructGenericContour
                Set oGenericContour = oFeature 'oFeatureColl.Item(icount)
            
                Dim oProVector As IJElements
            
                oGenericContour.GetProjectionVectors oProVector
                                                                                                           
                Dim oVector As IJDVector
                Set oVector = oProVector.Item(1)
            Else ' In case of APS Marks
                If UpSide = JXSEC_WEB_LEFT Or UpSide = JXSEC_WEB_RIGHT Then
                    Set oVector = oWebSBNormal
                Else
                    Dim oFlangeSB As IJSurfaceBody
                    oProfileSupport.GetProfileContours TopFlangeTopFace, oFlangeSB, oSurfaceContourColl, oMonikerColl
                        
                    oFlangeSB.GetCenterOfGravity oCOG
                    oMfgMGHelper.ProjectPointOnSurfaceBody oFlangeSB, oCOG, oProjPoint, oProjVector
                    
                    Dim oFlangeSBNormal As IJDVector
                    oFlangeSB.GetNormalFromPosition oProjPoint, oFlangeSBNormal
                    
                    Set oVector = oFlangeSBNormal
                End If
            
            End If
                        
            Dim dDotPro As Double
            dDotPro = oVector.Dot(oWebSBNormal)
            
            Dim lFaceId As Long
            Dim oSurfaceBody As IJSurfaceBody
            
            Dim bMarkOnBothFlanges   As Boolean
            
            If Abs(dDotPro) >= 0.99 Then
                Set oSurfaceBody = oWebSB
                lFaceId = UpSide
            Else
                Dim oProfileSection As IJDProfileSection
                Set oProfileSection = oThisPart
                    
                Dim oCrossSection As IJCrossSection
                Set oCrossSection = oProfileSection.crossSection
                
                Dim oCsSmartClass As IJDCrossSectionPartClass
                Set oCsSmartClass = oCrossSection.GetPartClass
                
                Dim strProfileCrossSection As String
                strProfileCrossSection = oCsSmartClass.CrossSectionTypeName
                        
                If UCase(strProfileCrossSection) = "TEEBAR" Then
                            
                    Set oProjVector = Nothing
                    Dim oTopFlangeSurface As IJSurfaceBody
                    oProfileSupport.GetProfileContours TopFlangeTopFace, oTopFlangeSurface, oSurfaceContourColl, oMonikerColl
                    Set oSurfaceBody = oTopFlangeSurface
                    lFaceId = JXSEC_TOP_FLANGE_TOP
                        
                Else
                    
                    Set oProjVector = Nothing
                    oProfileSupport.GetProfileContours TopFlangeTopFace, oTopFlangeSurface, oSurfaceContourColl, oMonikerColl
                    
                    Dim oPrimaryPortMoniker As IMoniker
                    Dim oSubPortMonikers As Collection
                    
                    oSDPartSupport.GetFeaturePortMonikers oFeature, oPrimaryPortMoniker, oSubPortMonikers
                    
                    Dim oFeaturePort As IJPort
                    oMfgMGHelper.BindMoniker oThisPart, oPrimaryPortMoniker, oFeaturePort
                    
                    Dim oModelBody As IJDModelBody
                    Set oModelBody = oFeaturePort.Geometry
                    
                    Dim oClosestPos1 As IJDPosition
                    Dim oClosestPos2 As IJDPosition
                    Dim dMinDistTopFlange As Double
                    Dim dMinDistBottomFlange As Double
                    
                    oModelBody.GetMinimumDistance oTopFlangeSurface, oClosestPos1, oClosestPos2, dMinDistTopFlange
                                            
                    If dMinDistTopFlange < 0.001 Then
                        oTopFlangeSurface.GetCenterOfGravity oCOG
                        oMfgMGHelper.ProjectPointOnSurfaceBody oTopFlangeSurface, oCOG, oProjPoint, oProjVector

                        Set oSurfaceBody = oTopFlangeSurface
                        lFaceId = JXSEC_TOP_FLANGE_TOP
                    End If
                    
                    Set oProjVector = Nothing
                    Dim oBottomFlangeSurface As IJSurfaceBody
                    oProfileSupport.GetProfileContours BottomFlangeBottomFace, oBottomFlangeSurface, oSurfaceContourColl, oMonikerColl
                        
                    oModelBody.GetMinimumDistance oBottomFlangeSurface, oClosestPos1, oClosestPos2, dMinDistBottomFlange
                    
                    If dMinDistBottomFlange < 0.001 Then
                        
                        oBottomFlangeSurface.GetCenterOfGravity oCOG
                        oMfgMGHelper.ProjectPointOnSurfaceBody oBottomFlangeSurface, oCOG, oProjPoint, oProjVector
                        
                        Set oSurfaceBody = oBottomFlangeSurface
                        lFaceId = JXSEC_BOTTOM_FLANGE_BOTTOM
                                                    
                    End If
                    
                    If dMinDistTopFlange < 0.001 And dMinDistBottomFlange < 0.001 Then
                        bMarkOnBothFlanges = True
                    Else
                        bMarkOnBothFlanges = False
                    End If
            
                End If
         
            End If
                                   
            Dim oMfgGeomHelper As New MfgGeomHelper
                                                                               
            Dim oBoxPoints As IJElements
            Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElems, oVector)
                                    
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
                        dXDotP = Abs(oWebSBNormal.Dot(oXVector))
                        dYDotP = Abs(oWebSBNormal.Dot(oYVector))
                        dZDotP = Abs(oWebSBNormal.Dot(oZVector))
                        
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
                                                                                     
            Set oCS1 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody1, oProjVector)
            Set oCS2 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody2, oProjVector)
                                                                                                                                           
            Dim oMfgGeomUtilwrapper1 As New MfgGeomUtilWrapper
            oMfgGeomUtilwrapper1.ExtendWire oCS1, 0.05
            Dim oMfgGeomUtilwrapper2 As New MfgGeomUtilWrapper
            oMfgGeomUtilwrapper2.ExtendWire oCS2, 0.05
            
            Dim oCS As IJComplexString
            Dim oSystemMark As IJMfgSystemMark
            Dim oMoniker As IMoniker
            Dim oMarkingInfo As MarkingInfo
            Dim oGeom3d As IJMfgGeom3d
            Dim oGeom3dBottom As IJMfgGeom3d
            Dim oNamedItem As IJNamedItem
            ''Set oNamedItem = oFeatureColl.Item(icount)
                                    
            Dim jCount As Integer
            For jCount = 1 To 2
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                
                'Set the marking side
                oSystemMark.SetMarkingSide UpSide
                
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
                                                                                                  
                If bMarkOnBothFlanges = True Then
                    
                    Dim kCount As Integer
                    For kCount = 1 To 2
                        Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        
                        If jCount = 1 Then
                            oGeom3d.PutGeometry oCS1
                        Else
                            oGeom3d.PutGeometry oCS2
                        End If
                        
                        If oCrossLine1.length < HOLE_MINIMUM_DIAMETER Or oCrossLine2.length < HOLE_MINIMUM_DIAMETER Then
                            oGeom3d.PutGeometrytype STRMFG_HOLE_TRACE_MARK 'Hole Trace mark
                        Else
                            oGeom3d.PutGeometrytype STRMFG_HOLE_REF_MARK 'Hole Ref mark
                        End If
                                                            
                        If kCount = 1 Then
                            oGeom3d.FaceId = JXSEC_TOP_FLANGE_TOP
                        Else
                            oGeom3d.FaceId = JXSEC_BOTTOM_FLANGE_BOTTOM
                        End If
                        
                        oGeomCol3d.AddGeometry 1, oGeom3d
                        oSystemMark.Set3dGeometry oGeom3d
                        
                        If oGeom3dCustom Is Nothing Then
                            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oFeature)
                        Else
                            Set oMoniker = oGeom3dCustom.GetMoniker
                        End If
                        
                        oGeom3d.PutMoniker oMoniker
                    Next
                Else
                
                    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    
                    If jCount = 1 Then
                        oGeom3d.PutGeometry oCS1
                    Else
                        oGeom3d.PutGeometry oCS2
                    End If
                    
                    If oCrossLine1.length < HOLE_MINIMUM_DIAMETER Or oCrossLine2.length < HOLE_MINIMUM_DIAMETER Then
                        oGeom3d.PutGeometrytype STRMFG_HOLE_TRACE_MARK 'Hole Trace mark
                    Else
                        oGeom3d.PutGeometrytype STRMFG_HOLE_REF_MARK 'Hole Ref mark
                    End If
                    
                    oGeom3d.FaceId = lFaceId
                    
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
            Set oMarkingInfo = Nothing
            Set oGeom3d = Nothing
            Set oMoniker = Nothing
            Set oGeom3dBottom = Nothing
                            
        End If

    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
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

Public Function CreatePenetrationMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, bFullMark As Boolean, bWebPenMark As Boolean, bMarkOnlyTripElem As Boolean) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "CreatePenetrationMark"
    On Error GoTo ErrorHandler
    
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSProfileMarkings STRMFG_PROFILE_TO_PROFILE_PENETRATION_MARK, ReferenceObjColl, oGeomCol3d
    
    Set CreatePenetrationMark = oGeomCol3d
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    '*** Get SD Wrapper Object ***'
    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = Part

    '*** Get Rule Helper Object ***'
    Dim oProfRuleHelper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfRuleHelper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfRuleHelper.object = Part
    
    'Hull plates cannot be processed
'    Dim ePlateType As StructProfileType
'    ePlateType = oSDProfileWrapper.ProfileType
'    If ePlateType = Hull Then
'        Set oPlateWrapper = Nothing
'        Set oSDProfileWrapper = Nothing
'        Exit Function
'    End If
    
    'Non-planar plates cannot be processed
'    If Not oPlateWrapper.CurvatureType = PLATE_CURVATURE_Flat Then
'        GoTo CleanUp
'    End If
        
    'Get plate's marking surface
    Dim oSurface As IJSurfaceBody
    Dim oThiknessSide As PlateThicknessSide
    
    'Convert between enums
    If UpSide = BaseSide Then
        oThiknessSide = PlateBaseSide
    Else
        oThiknessSide = PlateOffsetSide
    End If
    
    
'    Dim oThisPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Dim oUpSideGeometry As IUnknown
    
'    Set oThisPlateWrapper.object = Part
    
    'Set oSurface = oSDProfileWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry
    Dim oMfgPart As IJMfgProfilePart
    If oProfRuleHelper.ProfileHasMfgPart(oMfgPart) Then
        Set oSurface = oMfgPart.BeforeUnfoldSurfaceGeometry
    End If
    
'    Dim oModelBody As IJDModelBody
'    Set oModelBody = oSDProfileWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry
'    Dim sPBTFileName As String
'    sPBTFileName = Environ("TEMP")
'    If sPBTFileName = "" Or sPBTFileName = vbNullString Then
'        sPBTFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'    End If
'    sPBTFileName = sPBTFileName & "ProfBefTrim.sat"
'    oModelBody.DebugToSATFile sPBTFileName
    
    Set oUpSideGeometry = oSurface
    
    'Retrieve collection of Feature objects for this plate
    Dim oFeatureObjectsColl As Collection
    Set oFeatureObjectsColl = oSDProfileWrapper.ProfileFeatures
    
    If oFeatureObjectsColl Is Nothing Then
        'Since there are not features on this plate there will be no penetration marks to be made
        GoTo CleanUp
    End If
    
    Dim oPlateParentSystem As Object
    Set oPlateParentSystem = m_oMfgRuleHelper.GetTopMostParentSystem(Part)
        
    'Dim oConnectionData As StructDetailObjects.ConnectionData
    Dim iFeatureIndex As Integer 'iteration object
    Dim oPt As New DPosition
    Dim oCol1 As Collection, oPortCol As Collection
    Dim bWebLeftMark As Boolean
    
    
    'iterate through the collection of features
    For iFeatureIndex = 1 To oFeatureObjectsColl.Count

        Set oPortCol = New Collection
        ' Sastry
        ' TR#40585
        If Not TypeOf oFeatureObjectsColl.Item(iFeatureIndex) Is IJStructFeature Then
            GoTo NextFeatureIndex
        End If
        
        'Drop if not Profile/Plate penetration (Slot).
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oFeatureObjectsColl.Item(iFeatureIndex)
        If Not oStructFeature.get_StructFeatureType = SF_Slot Then GoTo NextFeatureIndex
        
        'Wrap the slot
        Dim oSlotWrapper As New StructDetailObjects.Slot
        Set oSlotWrapper.object = oFeatureObjectsColl.Item(iFeatureIndex)
    
        'Get penetrating profile object
        Dim oPenetratingProfile As IJProfilePart
        Set oPenetratingProfile = oSlotWrapper.Penetrating
        
        Dim bIsThereAnyTrippingElement As Boolean
        Dim bIsCrossSectionAllowed As Boolean
        Dim oProfileParentSystem As Object
        
        Set oProfileParentSystem = m_oMfgRuleHelper.GetTopMostParentSystem(oPenetratingProfile)
        
        '*** Check for tripping elements ***'
        bIsThereAnyTrippingElement = m_oMfgRuleHelper.CheckIfThereIsAnyTrippingElem(oPlateParentSystem, oProfileParentSystem)
        
        If bMarkOnlyTripElem <> True Then
            bIsThereAnyTrippingElement = False
        End If
        '***********************************'
        
        If bMarkOnlyTripElem = True And bIsThereAnyTrippingElement = False Then
            GoTo NextFeatureIndex
        End If
        
        bIsCrossSectionAllowed = CheckIfCrossSectionAllowed(oPenetratingProfile)
        
        Set oProfileParentSystem = Nothing
        
        'wrap penetrating profile
        Dim oPenetratingProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
        Set oPenetratingProfileWrapper.object = oPenetratingProfile
        
        'Get profile port opposite mounting face
        Dim oPortOppositMount As IJPort
        Set oPortOppositMount = oPenetratingProfileWrapper.GetPortOppositeMountingFace
        
        oPortCol.Add oPortOppositMount
        
        '*** Case - If Web Penetration Mark is needed ***'
'                               |
'                              _|___
'                           _ /     \_
'                            |   __ |
'                            |  |
'                     _______|  |__________

        If bWebPenMark = True Then
            Set oPortOppositMount = oPenetratingProfileWrapper.GetSurfacePort(JXSEC_WEB_LEFT)
            oPortCol.Add oPortOppositMount
        End If
        '**************************************************'

        For Each oPortOppositMount In oPortCol
        
            Set oCol1 = New Collection
            
            'Get intersection product between two surfaces
            Dim oIntersectionWB As IJWireBody
            Dim oPortGeometry As IUnknown
            
            If Not oPortOppositMount Is Nothing Then
                Set oPortGeometry = oPortOppositMount.Geometry
            End If
            
'            Dim oMB1 As IJDModelBody
'            Set oMB1 = oUpSideGeometry
'            Dim sUGFileName As String
'            sUGFileName = Environ("TEMP")
'            If sUGFileName = "" Or sUGFileName = vbNullString Then
'                sUGFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'            sUGFileName = sUGFileName & "oUpsideGeometry.sat"
'            oMB1.DebugToSATFile sUGFileName
'
'            Set oMB1 = oPortGeometry
'            Dim sPGFileName As String
'            sPGFileName = Environ("TEMP")
'            If sPGFileName = "" Or sPGFileName = vbNullString Then
'                sPGFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'            sPGFileName = sPGFileName & "oPortGeometry.sat"
'            oMB1.DebugToSATFile sPGFileName
            
            Set oIntersectionWB = m_oMfgRuleHelper.GetCommonGeometry(oUpSideGeometry, oPortGeometry)
            If oIntersectionWB Is Nothing Then
'                Set oMB1 = oIntersectionWB
'                Dim sIWBFileName As String
'                sIWBFileName = Environ("TEMP")
'                If sIWBFileName = "" Or sIWBFileName = vbNullString Then
'                    sIWBFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'                End If
'                sIWBFileName = sIWBFileName & "oIntersectionWB.sat"
'                oMB1.DebugToSATFile sIWBFileName
                GoTo NextFeatureIndex
                'Exit Function
            End If
            
            '*** Offset the penetration mark by 1mm  as per customer request ***'
            'Get SurfaceNormal at Flange Top Surface - We need to shift it by 0.001m
            ' ... As per Customer request.
            
            Dim oFlangeNormal As IJDVector
            Dim oDummy As IJDPosition
            Dim oTL As GSCADStructGeomUtilities.TopologyLocate
            Set oTL = New TopologyLocate
            oTL.FindApproxCenterAndNormal oPortGeometry, oDummy, oFlangeNormal
            
            Dim oIntersectionWB2 As IJWireBody
            
            If Not oFlangeNormal Is Nothing Then
                Set oIntersectionWB2 = m_oMfgRuleHelper.OffsetCurve(oUpSideGeometry, oIntersectionWB, oFlangeNormal, PROFILE_TO_PLATE_PEN_STRETCH_LENGTH, False)
            End If
            
            If Not oIntersectionWB2 Is Nothing Then
                Set oIntersectionWB = oIntersectionWB2
                Set oIntersectionWB2 = Nothing
            End If
            '*******************************************************************'
            
            'Get two end positions
            Dim oStartPos  As IJDPosition
            Dim oEndPos   As IJDPosition
            Dim oStartPosDir As IJDVector
            Dim oEndPosDir As IJDVector
            oIntersectionWB.GetEndPoints oStartPos, oEndPos, oStartPosDir, oEndPosDir
        
            'Create  line through the positions
            
            Dim xStart As Double
            Dim yStart As Double
            Dim zStart As Double
            Dim xEnd As Double
            Dim yEnd As Double
            Dim zEnd As Double
            Dim endX As Double
            Dim endY As Double
            Dim endZ As Double
        
            Dim oEndPosColl As Collection
            Set oEndPosColl = New Collection
            
            oEndPosColl.Add oStartPos
            oEndPosColl.Add oEndPos
            
            Dim oTempPos As IJDPosition
            Dim iMarkCount As Integer
            iMarkCount = 0
            
            For Each oTempPos In oEndPosColl
            
                iMarkCount = iMarkCount + 1
                Dim oLine As New Line3d
                oStartPos.Get xStart, yStart, zStart
                oEndPos.Get xEnd, yEnd, zEnd
                
                'Create line
                If oTempPos Is oStartPos Then
                    oLine.DefineBy2Points xStart, yStart, zStart, (2# * xStart) - xEnd, (2# * yStart) - yEnd, (2# * zStart) - zEnd
                ElseIf oTempPos Is oEndPos Then
                    oLine.DefineBy2Points xEnd, yEnd, zEnd, (2# * xEnd) - xStart, (2# * yEnd) - yStart, (2# * zEnd) - zStart
                End If
                
                oLine.length = 5
                
                Dim oCS As New ComplexString3d
                Dim oLineWB As IJWireBody
                oCS.AddCurve oLine, False
                Set oLineWB = m_oMfgRuleHelper.ComplexStringToWireBody(oCS)
            
                'Retrieve symbol contour
                oStructFeature.CacheSymbolOutput
                Dim oAttribWire As IJWireBody
                Dim dMaximumExtrusionDistance As Double
                Dim dMinimumExtrusionDistance As Double
                Dim bUseExtrusionDistance As Boolean
                Dim oProjectionVecors As IJElements
                oStructFeature.GetCachedSymbolOutput oAttribWire, dMaximumExtrusionDistance, _
                    dMinimumExtrusionDistance, bUseExtrusionDistance, oProjectionVecors
                    
                Dim oMDBody As IJDModelBody
                Dim oPosOnFeature As New DPosition
                Dim oPosOnLine As New DPosition
                Dim dDist As Double
                
                Set oMDBody = oAttribWire
                oMDBody.GetMinimumDistance oLineWB, oPosOnFeature, oPosOnLine, dDist
            
                'Prepare mark point collection
                Dim oMarkPosColl As New Collection
                Set oMarkPosColl = New Collection
                
                'retrieve intersection positions
    '            Set oMarkPosColl = m_oMfgRuleHelper.GetPositionsFromPointsGraph(oCommonGeom)
                oMarkPosColl.Add oPosOnFeature
                
                If oMarkPosColl.Count = 0 Then Exit Function
            
                'Iterate through the mark points collection and create ML
                Dim oMarkPointPos As IJDPosition
                Dim iIndex As Integer
                
                'For Each oMarkPointPosCount In oMarkPosColl (supposed to be only one)
                For iIndex = 1 To oMarkPosColl.Count
                    
                    'Project oLine on plate's surface if nessesary
                    If TypeOf oMarkPosColl.Item(iIndex) Is IJDPosition Then
                        Set oMarkPointPos = oMarkPosColl.Item(iIndex)
                    Else
                        GoTo NextMarkPointPos
                    End If
                    
                    'Create mark line from the part of oLine
                    Dim oMarkLine As New Line3d
            
                    oMarkPointPos.Get xStart, yStart, zStart
                    oLine.GetDirection xEnd, yEnd, zEnd
                    endX = xStart + (xEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    endY = yStart + (yEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    endZ = zStart + (zEnd * PROFILE_TO_PLATE_PEN_MARK_LENGTH)
                    oMarkLine.DefineBy2Points xStart, yStart, zStart, endX, endY, endZ
                    
                    '*** Case I - Partial Marks ***'
'                              _|___
'                           _ /     \_
'                            |   __ |
'                            |  |
'                     _______|  |__________
                            
                    If bFullMark = False And bIsThereAnyTrippingElement = False Then
                        Set oGeomCol3d = GetPenMarkGeom3D(oGeomCol3d, oMarkLine, UpSide, oStructFeature)
                    End If
                    '*******************************'
                    
                    Set oPt = New DPosition
    
                    oPt.x = endX
                    oPt.y = endY
                    oPt.z = endZ
                    
                    oCol1.Add oPt
                    
                    Set oMarkLine = Nothing
                    
                Next iIndex
    
            'CleanUp
            Set oLine = Nothing
            Set oCS = Nothing
            Set oLineWB = Nothing
        
            Next oTempPos
            
            Dim oPt1 As New DPosition, oPt2 As New DPosition
                
            Set oPt1 = oCol1.Item(1)
            Set oPt2 = oCol1.Item(2)
            
            '*** Case II - Full Marks ***'
'                              _|___
'                           __/_|___\_
'                            |  |__ |
'                            |  |
'                     _______|  |__________
            
            If bFullMark = True Or (bIsThereAnyTrippingElement = True And bIsCrossSectionAllowed = True) Then
                oMarkLine.DefineBy2Points oPt1.x, oPt1.y, oPt1.z, oPt2.x, oPt2.y, oPt2.z
                Set oGeomCol3d = GetPenMarkGeom3D(oGeomCol3d, oMarkLine, UpSide, oStructFeature)
            End If
            '*******************************'
               
            
    
NextMarkPointPos:
            'CleanUp
            Set oMarkPointPos = Nothing
 
        Next oPortOppositMount
        
        Set oCol1 = Nothing
        bWebLeftMark = False
        Set oEndPosColl = Nothing
        'Set oPortCol = Nothing
        
NextFeatureIndex:
    Next iFeatureIndex
    
    'Return collection of ML's geometry
    Set CreatePenetrationMark = oGeomCol3d
    
CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oPlateParentSystem = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1024, , "RULES")
    GoTo CleanUp
End Function

Private Function CheckIfCrossSectionAllowed(oProfilePart As Object) As Boolean

    Dim oProfileSection As IJDProfileSection
    Set oProfileSection = oProfilePart
    
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = oProfileSection.crossSection
    
    Dim oCSPartClass As IJDCrossSectionPartClass
    Set oCSPartClass = oCrossSection.GetPartClass
    
    Dim strProfileCrossSection As String
    strProfileCrossSection = oCSPartClass.CrossSectionTypeName
    
'    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
'    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    CheckIfCrossSectionAllowed = False
    
    Select Case UCase(strProfileCrossSection)
   
        Case "EQUALANGLE", "UNEQUALANGLE", "BULBFLAT", "FLATBAR", "IBAR", "CHANNEL", "TEEBAR"
            CheckIfCrossSectionAllowed = True
        Case Else
            
    End Select
    
CleanUp:
    
    Set oProfileSection = Nothing
    Set oCrossSection = Nothing
    Set oCSPartClass = Nothing
        
End Function

Private Function GetPenMarkGeom3D(oGeomCol3d As IJMfgGeomCol3d, oMarkLine As Line3d, ByVal UpSide As Long, oStructFeature As IJStructFeature) As IJMfgGeomCol3d

    Dim oCS_ML As New ComplexString3d
    oCS_ML.AddCurve oMarkLine, False
    
    Dim oSystemMark As IJMfgSystemMark
    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oSystemMark.SetMarkingSide UpSide
    
    Dim oMarkingInfo As MarkingInfo
    Dim oMoniker As IMoniker
    Dim oObjSystemMark As IUnknown
    
    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark
    oMarkingInfo.name = "PROFILE PEN FITTING MARK"
    
    'Add ML's geometry to collection
    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    oGeom3d.FaceId = JXSEC_WEB_LEFT
    oGeom3d.PutGeometry oCS_ML
    oGeom3d.PutGeometrytype STRMFG_PROFILE_TO_PROFILE_PENETRATION_MARK
    
    Set oObjSystemMark = oSystemMark
    
    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oStructFeature)
    oGeom3d.PutMoniker oMoniker
        
    oSystemMark.Set3dGeometry oGeom3d
    oGeomCol3d.AddGeometry 1, oGeom3d
    
    Set GetPenMarkGeom3D = oGeomCol3d
    
    Set oCS_ML = Nothing
    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oMoniker = Nothing
    Set oObjSystemMark = Nothing
    Set oGeom3d = Nothing
    
Exit Function
ErrorHandler:

End Function

Public Function CreateLocationMark(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, Optional ByVal bAddFittingMark As Boolean) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
     Const METHOD = "CreateLocationMark"
    
    On Error GoTo ErrorHandler
    
    Dim oSDProfileWrapper As Object
    'Create the SD profile Wrapper and initialize it
    If TypeOf Part Is IJStiffenerPart Then
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = Part
    ElseIf TypeOf Part Is IJBeamPart Then
        Set oSDProfileWrapper = New StructDetailObjects.BeamPart
        Set oSDProfileWrapper.object = Part
    End If
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oPartInfo As IJDPartInfo
    Set oPartInfo = New PartInfo
    
    'Get the Profile Part Physically Connected Objects
    Dim oConObjsCol As Collection
    'Set oConObjsCol = oSDProfileWrapper.ConnectedObjects
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_PLATELOCATION_MARK, ReferenceObjColl, oGeomCol3d
    CreateAPSProfileMarkings STRMFG_MOUNT_ANGLE_MARK, ReferenceObjColl, oGeomCol3d
    
    Set CreateLocationMark = oGeomCol3d
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the function
        GoTo CleanUp
    End If
    
    Dim oMfgMathGeom As New MfgMathGeom
    
    Dim oUpsideSurface As IUnknown
    Dim oSurfacePort As IJPort
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(UpSide)
    Set oUpsideSurface = oSurfacePort.Geometry
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long
    Dim oSDConProfileWrapper As New StructDetailObjects.ProfilePart
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    Dim IsPartBuiltUp As Boolean
    IsPartBuiltUp = IsBuiltUp(Part)
    
    ' Loop thru each Physical Connections
    Dim oVector As IJDVector
    Dim bContourTee As Boolean
    Dim oWB As IJWireBody, oTeeWire As IJWireBody
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
    
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        ' Check if this physical connection is a Root PC, which particpates in Split operation
        ' If so, Skip the marking line line creation for this Root PC.
        Dim oStructEntityOperation As IJDStructEntityOperation
        Dim opeartionProgID As String
        Dim opeartionID As StructOperation
        Dim oOperColl As New Collection
    
        Set oStructEntityOperation = oConnectionData.AppConnection
        oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl
        
        Set oStructEntityOperation = Nothing
        Set oOperColl = Nothing
        
        ' If the RootPC has Split operation in its graph, just goto the next pc.
        If opeartionID = ConnectionSplitOperation Then
            ' No need to create marking line.
            GoTo NextItem
        End If
        
        'Get the sub port having section information
        Dim oSubPort As IJStructPort
        Set oSubPort = oProfileWrapper.GetProfileSubPort(oConnectionData.ConnectingPort, oConnectionData.AppConnection)
        If Not ((oSubPort.SectionID = JXSEC_WEB_RIGHT) Or _
            (oSubPort.SectionID = JXSEC_WEB_LEFT) Or _
             oSubPort.SectionID = JXSEC_TOP) Then
            Set oSubPort = Nothing
            GoTo NextItem
        End If
        
        Dim eMoldedDir As StructMoldedDirection
        Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
        
        eMoldedDir = invalidDirection
        eSideOfConnectedObjectToBeMarked = SideA
        Dim sMoldedSide As String

        'Initialize the profile wrapper and the Physical Connection wrapper
        If TypeOf oConnectionData.ToConnectable Is IJPlatePart Or TypeOf oConnectionData.AppConnection Is IJSmartPlate Or TypeOf oConnectionData.AppConnection Is IJCollarPart Then
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
        End If
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDConProfileWrapper = New StructDetailObjects.ProfilePart
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
            
            
        End If
        
        If TypeOf Part Is IJStiffenerPart Then

            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oConnectionData.AppConnection

            bContourTee = oSDProfileWrapper.Connection_ContourTee(oConnectionData.AppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)

            If (bContourTee = True) And (Not (oTeeWire Is Nothing)) Then
                ' Bound the wire based on split points, if there are any.
                Set oWB = TrimTeeWireForSplitPCs(oConnectionData.AppConnection, oTeeWire)

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
                        
                    'oMfgMathGeom.ProjectCurveToSurface oGeom3d, oUpsideSurface, Nothing
                        
                    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                    lGeomCount = lGeomCount + 1
                
                    If IsPartBuiltUp Then
                        CreateDeclivityMarksOnProfile Part, oConnectionData, oCS, MarkingFace, _
                                                      oGeom3d, oGeomCol3d, lGeomCount, sMoldedSide, Nothing
                    End If
NextComplexString:
                    Set oSystemMark = Nothing
                    Set oMarkingInfo = Nothing
                    Set oGeom3d = Nothing
                Next
                oCSColl.Clear
                Set oCSColl = Nothing
            End If
        End If
NextItem:
        Set oSubPort = Nothing
        Set oWB = Nothing
        Set oTeeWire = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConProfileWrapper = Nothing
        Set oSDPhysicalConn = Nothing
        name = ""
        thickness = 0
    Next nIndex
    
    'Return the 3d collection
    Set CreateLocationMark = oGeomCol3d
    
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing
    Set oMfgMathGeom = Nothing
    Set oSurfacePort = Nothing
    Set oUpsideSurface = Nothing

CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oPartInfo = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2008, , "RULES")
    GoTo CleanUp
End Function

'********************************************************************
' Routine: CreateEndFittingMark
' Description: This function Creates the End Fitting Mark(s) at the Free End(s) of the
'               ... Profile Parts.
' Inputs:  Plate Part, Collection of Profile Location Marking Lines, Molded Side, UpSide,
'          ... Geom3D Collection, NormalVector to the Plate Part
' Outputs: Procedure creates the "End Fitting Marks"
' Notes:
'********************************************************************

Public Sub CreateEndFittingMark(Part As Object, _
                                    oCS As IJComplexString, _
                                    oWB As IJWireBody, _
                                    oPartSupport As Object, _
                                    oSDPlateWrapper As StructDetailObjects.ProfilePart, _
                                    oSDProfileWrapper As Object, _
                                    sMoldedSide As String, _
                                    UpSide As Long, _
                                    oConnectionData As ConnectionData, _
                                    oGeomCol3d As IJMfgGeomCol3d, _
                                    oPlateNormal As IJDVector, _
                                    bFlip As Boolean)
        
        Const METHOD = "CreateEndFittingMark"
        On Error GoTo ErrorHandler
                
        '*** Declaration of Variables ***'
        Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
        Dim lNoOfObj As Long, i As Long
        Dim oObj As Object, oConnectedObj As Object
        Dim oPort1 As IJPort
        Dim onamed As IJNamedItem
        Dim oPlateName As IJNamedItem
        Dim oPCWB As IJWireBody
        Dim oProjectedPt As IJDPosition
        Dim dDistFromStart As Double, dDistFromEnd As Double, dDepth As Double
        Dim bIsStartFreeEnd As Boolean, bIsEndFreeEnd As Boolean
        Dim oGeom3d As IJMfgGeom3d
        Dim oToolBox As IJDTopologyToolBox
        Dim oProfileWB As IJWireBody
        Dim oStartPos As New DPosition, oEndPos As New DPosition
        Dim oStartDir As New DVector, oEndDir As New DVector, oTDir As New DVector
        Dim oProfilePartSupport As IJProfilePartSupport
        Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
        
        Dim oPlatePartSupport As IJPlatePartSupport
        Dim dThickness As Double, dCriterion As Double
        
        '*** Initialization of Variables ***'
        Set oToolBox = New IMSModelGeomOps.DGeomOpsToolBox
          
        '*** For the Given Profile, Get collections of
        '... 1. Physical Connections
        '... 2. Connected Objects
        '... 3. Objects Connected at Base Port
        '... 4. Objects Connected at Offset Port
        oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol
          
        '*** Get the # of Physical Connections ***'
        lNoOfObj = oConObjCol.Count
        
        '*** Get the Start & End Points of the WireBody ***'
        oWB.GetEndPoints oStartPos, oEndPos
        
        bIsStartFreeEnd = True
        bIsEndFreeEnd = True
        
        '*** Get the Profile WireBody ***'
        Set oProfileWB = oMfgRuleHelper.ComplexStringToWireBody(oCS)
        
        '*** Get the start and end positions and their directions ***'
        oProfileWB.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir
        'The oEndDir is nothing but the Tangent Vector to the Profile Location Marking Line/Curve
                
                
        '*** Get the Profile Depth to set the distance tolerence to check if a physical connection
        '... exists at the end of the profile part ***'
        'Set oProfilePartSupport = oPartSupport
        'oProfilePartSupport.GetWebDepth dDepth
        If TypeOf oPartSupport Is IJPlatePartSupport Then
            Set oPlatePartSupport = oPartSupport
            oPlatePartSupport.GetThickness dThickness
            dCriterion = dThickness
        ElseIf TypeOf oPartSupport Is IJProfilePartSupport Then
            Set oProfilePartSupport = oPartSupport
            oProfilePartSupport.GetWebDepth dDepth
            dCriterion = dDepth
        End If
        
        '*** lNoOfObj > 1 means that there are other objects (other than base plate)
        '... having physical connection with the profile.
        '... So check which end of the profile is/are Free Ends, and place the marks accordingly.
        If lNoOfObj > 1 Then
           
            For i = 1 To lNoOfObj
            
                ' Get the object connected at the other port
                Set oPort1 = oOtherPortCol.Item(i)
                Set oObj = oPort1.Connectable
                
                ' This gives the name of the object connected at the other port
                Set onamed = oObj

                ' This gives the name of the Base Plate
                Set oPlateName = oSDPlateWrapper.object
                
                'If the Connected Object is Base Plate, ignore it...
                '... and if not, then get the distance of PC from StartPt and EndPt.
                '... if the distance is greater than certain tolerance, that means the object is not
                '... connected, so this is a free end cut. Then place the End Fitting Mark.
                
                'If oNamed.Name <> oPlateName.Name Then
                If Not (oObj Is oSDPlateWrapper.object) Then
                    
                    '*** Get the Geometry of the Connected Object from the Port ***'
                    Set oConnectedObj = oPort1.Geometry
                    
                    '*** Get the WireBody object for the Location Marking Line ***'
                    Set oPCWB = oConnCol.Item(i)
                    
                    '*** Find the closest point on the WireBody of Physical Connection
                    '... in order to find the minimum distance from Start Point and End Point
                    oToolBox.GetNearestPointOnWireBodyFromPoint oPCWB, oStartPos, Nothing, oProjectedPt
                    
                    '*** Get the distance of the point on WB from StartPoint***'
                    dDistFromStart = oProjectedPt.DistPt(oStartPos)
                    
                    oToolBox.GetNearestPointOnWireBodyFromPoint oPCWB, oEndPos, Nothing, oProjectedPt
                    
                    '*** Get the distance of the point on WB from EndPoint***'
                    dDistFromEnd = oProjectedPt.DistPt(oEndPos)
                    
                    '*** If the distance is less than WebDepth of the profile, then the object has
                    '... physical connection with the Profile End
'                    If dDistFromStart < dDepth Then bIsStartFreeEnd = False
'                    If dDistFromEnd < dDepth Then bIsEndFreeEnd = False
                    If dDistFromStart < dCriterion Then bIsStartFreeEnd = False
                    If dDistFromEnd < dCriterion Then bIsEndFreeEnd = False
                
                End If
                
                
            Next i
            
            
            '*** This means that the Start End is Free End, Place the Fitting Mark Here ***'
            If bIsStartFreeEnd Then
                '*** Create the End Fitting Mark at StartPoint***'
                Set oTDir = oStartDir
                Set oGeom3d = GetEndFittingMarkAtPoint(oStartPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oStartDir, oPlateNormal, oTDir, bFlip)
                '*** Add the Geom3d object to collection ***'
                oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
            End If
            
            '*** This means that the Other End is Free End, Place the Fitting Mark Here ***'
            If bIsEndFreeEnd Then
                '*** Create the End Fitting Mark at EndPoint***'
                Set oTDir = oEndDir
                oTDir.length = -1#
                Set oGeom3d = GetEndFittingMarkAtPoint(oEndPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oEndDir, oPlateNormal, oTDir, bFlip)

                '*** Add the Geom3d object to collection ***'
                oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
            End If
            
        '*** lNoOfObj = 1 means that the only physical connection that is present is
        '... between the profile and the base plate.
        '... So place the fitting marks at both ends of the profile.
        Else
        
            '*** Create the End Fitting Mark at StartPoint***'
            Set oTDir = oStartDir
            Set oGeom3d = GetEndFittingMarkAtPoint(oStartPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oStartDir, oPlateNormal, oTDir, bFlip)
            
            '*** Add the Geom3d object to collection ***'
            oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
            
            '*** Create the End Fitting Mark at EndPoint***'
            Set oTDir = oEndDir
            oTDir.length = -1#
            Set oGeom3d = GetEndFittingMarkAtPoint(oEndPos, oSDProfileWrapper, sMoldedSide, UpSide, oConnectionData, oEndDir, oPlateNormal, oTDir, bFlip)
            
            '*** Add the Geom3d object to collection ***'
            oGeomCol3d.AddGeometry oGeomCol3d.GetColcount + 1, oGeom3d
        
        End If
        
        
CleanUp:
        Set oConObjCol = Nothing
        Set oConnCol = Nothing
        Set oThisPortCol = Nothing
        Set oOtherPortCol = Nothing
        Set oObj = Nothing
        Set oConnectedObj = Nothing
        Set oPort1 = Nothing
        Set onamed = Nothing
        Set oPlateName = Nothing
        Set oPCWB = Nothing
        Set oProjectedPt = Nothing
        Set oToolBox = Nothing
        Set oProfileWB = Nothing
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oStartDir = Nothing
        Set oEndDir = Nothing
        Set oProfilePartSupport = Nothing
        
        Exit Sub


ErrorHandler:

        Err.Raise Err.Number, Err.Source, Err.Description
        
End Sub

'********************************************************************
' Routine: GetEndFittingMarkAtPoint
' Description: Creates the End Fitting Mark Line (of specified Length) normal to the
'              ... Location mark line.
' Inputs: Point at which Mark is to be placed, ProfileWrapper object, Molded Side, UpSide,
'              ... Tangent and Normal Vectors to the surface body
' Outputs:  End Fitting Mark Line
' Notes:
'********************************************************************
Public Function GetEndFittingMarkAtPoint(oStartPos As IJDPosition, _
                                            oSDProfileWrapper As Object, _
                                            sMoldedSide As String, _
                                            UpSide As Long, _
                                            oConnectionData As ConnectionData, _
                                            oTanVec As IJDVector, _
                                            oPlateNormal As IJDVector, _
                                            oTDir As IJDVector, _
                                            bFlip As Boolean, _
                                            Optional dLength As Double = 0) As IJMfgGeom3d

        
        Const METHOD = "GetEndFittingMarkAtPoint"
        On Error GoTo ErrorHandler
        
        '*** Declaration of Variables ***'
        Dim oThickVecDir2D As New DVector
        Dim oNewPos As New DPosition
        Dim oLineCS As IJComplexString
        Dim oLine1 As IJLine
        Dim oSystemMark As IJMfgSystemMark
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3d As IJMfgGeom3d
        Dim oMoniker As IMoniker
        Dim oNormalVec As New DVector
        Dim dDot As Double
        
        
        If dLength = 0 Then
             dLength = END_FITTING_MARK_LENGTH
        End If
     
        '*** Initialization of Variables ***'
        Set oLine1 = New Line3d
        Set oLineCS = New ComplexString3d
        
        '*** Get the thickness direction ***'
        Set oThickVecDir2D = GetThicknessDirectionVectorAtAGivenPos(oStartPos, oSDProfileWrapper, Nothing, sMoldedSide)
        oThickVecDir2D.length = dLength
        
        '*** Get the Normal to the Plate surface in order to place the Fitting Mark ***'
        Set oNormalVec = oTanVec.Cross(oPlateNormal)
        
        dDot = oNormalVec.Dot(oThickVecDir2D)
        
        '*** Determine the thickness direction  ***'
        If dDot > 0 Then
            oNormalVec.length = dLength
        Else
            oNormalVec.length = dLength * -1#
        End If
        
        '*** Flip the Mark  ***'
        If bFlip Then
            oNormalVec.length = dLength * -1#
        End If
        
        '*** Get the other co-ordinates of the mark ***'
        Set oNewPos = oStartPos.Offset(oNormalVec)
        
        '*** Construct the Mark ***'
        oLine1.DefineBy2Points oStartPos.x, oStartPos.y, oStartPos.z, oNewPos.x, oNewPos.y, oNewPos.z
        
        '*** Create the Complex String ***'
        oLineCS.AddCurve oLine1, True

        Dim oSystemMarkFactory As New GSCADMfgSystemMark.MfgSystemMarkFactory
        
        Set oSystemMark = oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        '*** Set the marking side ***'
        oSystemMark.SetMarkingSide UpSide
        
        '*** QI for the MarkingInfo object on the SystemMark ***'
        Set oMarkingInfo = oSystemMark
        
        '*** Set the name and thickness for marking info ***'
        oMarkingInfo.name = ""
        
        'Set the Fitting Mark's thickness direction (i.e. towards Material or not)
        oMarkingInfo.ThicknessDirection = oTDir
        
        Dim oGeom3dFactory As New GSCADMfgGeometry.MfgGeom3dFactory
        Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
        
        Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        oGeom3d.PutGeometry oLineCS
        oGeom3d.FaceId = UpSide
        
        '*** Set the Type of the Mark ***'
        oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
        oSystemMark.Set3dGeometry oGeom3d
        
        Set oMoniker = oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
        oGeom3d.PutMoniker oMoniker
        
        
        '*** Return the Complex String ***'
        Set GetEndFittingMarkAtPoint = oGeom3d
        
        
CleanUp:
        
        Set oMarkingInfo = Nothing
        Set oThickVecDir2D = Nothing
        Set oNewPos = Nothing
        Set oLineCS = Nothing
        Set oLine1 = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oMoniker = Nothing
        Set oNormalVec = Nothing
        
        Exit Function

ErrorHandler:

        Err.Raise Err.Number, Err.Source, Err.Description
        
End Function

Public Function GetThicknessDirectionVectorAtAGivenPos(oInputPos As IJDPosition, oSDConWrapper As Object, oUpSideSurf As IJSurfaceBody, sMoldedSide As String) As IJDVector
    Const METHOD = "GetThicknessDirectionVectorAtAGivenPos"
    
    On Error GoTo ErrorHandler
    Dim oProjPos As IJDPosition
    
    'we must adjust the returned TD vector for each of the new seperate complex strings
    Dim oSurfaceBodyCon1 As IJSurfaceBody
    If TypeOf oSDConWrapper.object Is IJPlatePart Then
        Dim oPartSupp As IJPartSupport
        Dim oPlatePartSupp As IJPlatePartSupport
        Set oPartSupp = New PlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = oSDConWrapper.object

        If sMoldedSide = "Base" Then
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateOffsetSide, oSurfaceBodyCon1
        Else
            oPlatePartSupp.GetSurfaceWithoutFeatures PlateBaseSide, oSurfaceBodyCon1
        End If
        
        '*** For Debug ***'
'        Dim oModelBody As IJDModelBody
'        Set oModelBody = oSurfaceBodyCon1
'        Dim FileName As String
'        FileName = Environ("TEMP")
'        If FileName = "" Or FileName = vbNullString Then
'            FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'        End If
'        FileName = FileName & "\oSurfaceBodyCon1.sat"
'        oModelBody.DebugToSATFile FileName
        '*****************'
                
        Set oPlatePartSupp = Nothing
        Set oPartSupp = Nothing
    ElseIf TypeOf oSDConWrapper.object Is IJProfilePart Then
        If sMoldedSide = "Base" Then
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_RIGHT).Geometry
        Else
            Set oSurfaceBodyCon1 = oSDConWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry
        End If
    End If
    
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    
    Set oProjPos = oMfgRuleHelper.ProjectPointOnSurface(oInputPos, oSurfaceBodyCon1, Nothing)

    If Not oUpSideSurf Is Nothing Then
        Dim oProjPosOnUpsideSurf As IJDPosition
        Set oProjPosOnUpsideSurf = oMfgRuleHelper.ProjectPointOnSurface(oProjPos, oUpSideSurf, Nothing)
        Set oProjPos = oProjPosOnUpsideSurf
    End If
    Set GetThicknessDirectionVectorAtAGivenPos = oProjPos.Subtract(oInputPos)
    
CleanUp:
    Set oSurfaceBodyCon1 = Nothing
Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1001, , "RULES")
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
    
    CreateAPSProfileMarkings STRMFG_SEAM_MARK, ReferenceObjColl, oGeomCol3d
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
    Dim strSectionSide As String
    strSectionSide = "WEB"
    
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
                        
            Dim oAppConnection As IJAppConnection
            Set oAppConnection = oConnectionData.AppConnection
            
            
            If Not oConObjsCol.Count = 1 Then
                If Not DoesPortBelongToSection(oStructPort1, oAppConnection, strSectionSide) _
                   And Not DoesPortBelongToSection(oStructPort2, oAppConnection, strSectionSide) Then
                   
                    GoTo NextItem
                   
                End If
            End If
            

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
                Set oFinalWireBody = CreateSeamControlLine(Part, oSeamPos, oConnectionData, UpSide, oDirVector, 0.01)
                
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
        Set oAppConnection = Nothing
        
        If nIndex = oConObjsCol.Count Then
            If Not bOffsetFound And Not bBaseFound Then
                Select Case UCase(strSectionSide)
                    Case "WEB"
                        strSectionSide = "TOPFLANGE"
                        nIndex = 0
                    Case "TOPFLANGE"
                        strSectionSide = "BOTTOMFLANGE"
                        nIndex = 0
                End Select
            End If
        End If
        
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

' ***********************************************************************************
' DoesPortBelongToSection
'
' Description:  This function will tell if a given port belongs to the desired section.
'                   if sectionName is left as Default it will return true if the port
'                   does not belong to a top or bottom flange.
'
' ***********************************************************************************
Public Function DoesPortBelongToSection(ByVal oPort As IJStructPort, ByVal oAppConnection As IJAppConnection, ByVal sectionName As String) As Boolean
    Const METHOD = "DoesPortBelongToSection"
    On Error GoTo ErrorHandler
    
    Dim bBelongsToSection As Boolean
    bBelongsToSection = False
    
    Dim oMfgEntityHelper As IJMfgEntityHelper
    Set oMfgEntityHelper = New MfgEntityHelper
    
    Dim portSection As JXSEC_CODE
    portSection = oPort.SectionID
    
    If portSection = -1 Or _
       portSection = JXSEC_UNKNOWN Or _
       portSection = JXSEC_IDEALIZED_BOUNDARY Then
               
        Dim oSubPort As IJStructPort
        Set oSubPort = oMfgEntityHelper.GetProfileSubPort(oPort, oAppConnection)
        
        portSection = oSubPort.SectionID
    End If
    
    Select Case UCase(sectionName)
        Case "WEB"
            
            If portSection = JXSEC_WEB_LEFT Or _
               portSection = JXSEC_WEB_RIGHT Or _
               portSection = JXSEC_WEB_LEFT_BOTTOM Or _
               portSection = JXSEC_WEB_LEFT_TOP Or _
               portSection = JXSEC_WEB_RIGHT_BOTTOM Or _
               portSection = JXSEC_WEB_RIGHT_TOP Or _
               portSection = JXSEC_INNER_TUBE Or _
               portSection = JXSEC_OUTER_TUBE Then
               
                bBelongsToSection = True
                
            Else
            
                bBelongsToSection = False
                
            End If
            
        Case "TOPFLANGE"
        
            If portSection = JXSEC_TOP Or _
               portSection = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
               portSection = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
               portSection = JXSEC_TOP_FLANGE_LEFT_TOP Or _
               portSection = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
               portSection = JXSEC_TOP_FLANGE_LEFT Or _
               portSection = JXSEC_TOP_FLANGE_RIGHT Or _
               portSection = JXSEC_TOP_FLANGE_TOP Or _
               portSection = JXSEC_TOP_FLANGE_BOTTOM Then
               
                bBelongsToSection = True
                
            Else
                
                bBelongsToSection = False
                
            End If
            
        Case "BOTTOMFLANGE"
        
            If portSection = JXSEC_BOTTOM Or _
               portSection = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
               portSection = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
               portSection = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
               portSection = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
               portSection = JXSEC_BOTTOM_FLANGE_LEFT Or _
               portSection = JXSEC_BOTTOM_FLANGE_RIGHT Or _
               portSection = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
               portSection = JXSEC_BOTTOM_FLANGE_TOP Then
                
                bBelongsToSection = True
                
            Else
            
                bBelongsToSection = False
                
            End If
               
    End Select
    
    DoesPortBelongToSection = bBelongsToSection
    
CleanUp:
    Set oSubPort = Nothing
    Set oMfgEntityHelper = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
