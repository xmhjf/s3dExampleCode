Attribute VB_Name = "GlobalHelpers"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    GlobalHelpers.bas
'
'  History:
'       TBH         april 8. 2002   created
'       MJV         april 14, 2004  Modified the error handling
'******************************************************************

Option Explicit
Const MODULE = "GlobalHelpers.bas"




' ***********************************************************************************
' Public Function GetSeamDistance
'
' Description:  Helper function to get length of a curve. Input is a Wirebody and
'               output is a double which is the length of the curve.
'
' ***********************************************************************************
Public Function GetSeamDistance() As Double
    GetSeamDistance = 0.1
End Function
Public Property Get ERFittingMarkLength() As Double
    ERFittingMarkLength = 0.06
End Property

Public Property Get FittingMarkLength() As Double
    FittingMarkLength = 0.015
End Property

Public Property Get FittingMarkDistance() As Double
    FittingMarkDistance = 0.1
End Property

Public Property Get EdgeCheckMarkVDistance() As Double
    EdgeCheckMarkVDistance = 0.1
End Property

Public Property Get EdgeCheckMarkUDistance() As Double
    EdgeCheckMarkUDistance = 0.2
End Property

Public Property Get WebFrameMarkDistanceFromHull() As Double
    WebFrameMarkDistanceFromHull = -0.2
End Property

' ***********************************************************************************
' Public Function GetActiveConnection
'
' Description: Helper function to get the active connection
'
' ***********************************************************************************
Public Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    oCmnAppGenericUtil.GetActiveConnection GetActiveConnection

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'****************************************************************************************************
'Description
'   Returns the GetActiveConnectionName.
'****************************************************************************************************
Public Function GetActiveConnectionName() As String

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing
    
End Function


' ***********************************************************************************
' Public Function GetFittingMarkPos
'
' Description:  Helper function to find fitting mark positions for the input part.
'               Input arguments: this part, connection data, marking side.
'               Output argument: collection of found mark point positions.
'               Mehod should be able to handle plates and profiles as input.
'
' Unresolved issues: 1. Avoidance of cutout areas is not implemented
'                    2. Method is not ready for profile marking (Plate/Profile T connection, this part = Profile)
' ***********************************************************************************
Public Function GetFittingMarkPositions(oThisPart As Object, _
    oConnectionData As ConnectionData, _
    UpSide As Long, _
    Optional oNeutralSurface As Object = Nothing) As Collection
    
    Const sMETHOD = "GetFittingMarkPos"

    On Error GoTo ErrorHandler
    'Following sequence is to be implemented in this method:

    '- Check if input parts are Plate/Plate or Profile/Plate, using part's wrappers

    '- If Plate/Plate then:
    '- Get connected ports of both plate parts
    '- Check if one port is lateral, and the other one is base (means 'T- conn')
    '- Take two ends positions of the lateral port edge and create mark points positions
    '- Check if this plate is the subject for marking with the found mark point
    '- Check if mark points positions are not in the feature area and move it as required
    '- Save final mark points positions in the oMarkPosColl and return

    '- If Profile/Plate then:
    '- Get landing curve of the profile part
    '- Take two ends positions of the landing curve and create mark points positions
    '- Check if this plate is the subject for marking with the found mark point
    '- Check if mark points positions are not in the feature area and move it as required
    '- Save final mark points positions in the oMarkPosColl and return

    'Prepare different helper objects
    Dim oMarkPosColl As New Collection
    Dim oCurve As IJWireBody
    Dim m_oMfgRuleHelper As New MfgRuleHelpers.Helper
    Dim oCS As IJComplexString
    Dim oEndPos1 As IJDPosition
    Dim oEndPos2 As IJDPosition
    Dim oConnPart As Object
    Dim oMarkPointPos1 As IJDPosition
    Dim oMarkPointPos2 As IJDPosition
    Dim oThisPartPort As IJPort
    Dim oConnPartPort As IJPort
    Dim oPosObj As IJDPosition
    Dim iIndex As Integer
    Dim oValidMarkPointColl As Collection 'temporary collection to keep valid points
    Dim oIntersectionWB As IUnknown
    Dim oLineGeometry As IUnknown
    Dim oPortGeometry As IUnknown
    Dim xStart As Double
    Dim yStart As Double
    Dim zStart As Double
    Dim xEnd As Double
    Dim yEnd As Double
    Dim zEnd As Double
    Dim endX As Double
    Dim endY As Double
    Dim endZ As Double
    Dim oTestedPortGeometry As IUnknown
    Dim oConnType As ContourConnectionType
    Dim bIsBaseOfTee As Boolean
    
    Set oConnPart = oConnectionData.ToConnectable

    '*********************** Plate/Plate case *******************************************
    If TypeOf oConnPart Is IJPlatePart And TypeOf oThisPart Is IJPlatePart Then

        'wrap plates
        Dim oConnPlateWrapper As New StructDetailObjects.PlatePart
        Set oConnPlateWrapper.object = oConnectionData.ToConnectable
        Dim oThisPlateWrapper As New StructDetailObjects.PlatePart
        Set oThisPlateWrapper.object = oThisPart


        'Get connected ports on both parts
        Set oThisPartPort = oConnectionData.ConnectingPort
        Set oConnPartPort = oConnectionData.ToConnectedPort

        'get some helper objects
        Dim oLateralPort As IJPort
        Dim oPlateConnByLateralPort As IJPlatePart
        Dim oFacePortOfLateralPlate As IJPort
        Dim oFaceSuface As IUnknown
        
        'Define connection type for this connection
        oConnType = oThisPlateWrapper.GetGetConnectionTypeForContour(oConnectionData.AppConnection, bIsBaseOfTee)

        If oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsBaseOfTee = True Then
    '       Connected plate is engaged by its lateral
            Set oPlateConnByLateralPort = oConnectionData.ToConnectable
            Set oLateralPort = oConnectionData.ToConnectedPort
            Set oFacePortOfLateralPlate = oConnPlateWrapper.BasePort(BPT_Offset)
            Set oFaceSuface = oFacePortOfLateralPlate.Geometry
        ElseIf oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsBaseOfTee = False Then
    '       This plate is engaged by its lateral
            Set oPlateConnByLateralPort = oThisPart
            Set oLateralPort = oConnectionData.ConnectingPort
            Set oFacePortOfLateralPlate = oThisPlateWrapper.BasePort(BPT_Offset)
            If oNeutralSurface Is Nothing Then
                Set oFaceSuface = oFacePortOfLateralPlate.Geometry
            Else
                Set oFaceSuface = oNeutralSurface
            End If
        Else
    '        Plates are not T connected
            Exit Function
        End If

        'Get surfaces of ports
        Dim oLateralSurface As IUnknown
        Set oLateralSurface = oLateralPort.Geometry

        'get cross product of surfaces
        Dim oCommonGeom As IUnknown
        Set oCommonGeom = m_oMfgRuleHelper.GetCommonGeometry(oLateralSurface, oFaceSuface)
        If oCommonGeom Is Nothing Then
            StrMfgLogError Err, MODULE, sMETHOD, , "SMCustomWarningMessages", 2004, , "RULES"
            GoTo NextPlateIndex
        End If
        
        Set oCurve = oCommonGeom

        'get end points of WB
        oCurve.GetEndPoints oEndPos1, oEndPos2
            
        'find two mark points at 'd' from both ends
       
        Dim oCompStr As IJComplexString
        Dim o3dCurve As IJCurve
        Dim dLength As Double
        Dim oCurveElems As IJElements
        
        Set oCurveElems = m_oMfgRuleHelper.WireBodyToComplexStrings(oCurve)
        
        For Each oCompStr In oCurveElems
            Set o3dCurve = oCompStr
            dLength = dLength + o3dCurve.Length
        Next
        If dLength < (2 * FittingMarkDistance) Then
            Exit Function
        End If
        
        Set oMarkPointPos1 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos1, FittingMarkDistance, oEndPos2)

        Set oMarkPointPos2 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos2, FittingMarkDistance, oEndPos1)

        'Add found mark points position to the collection
        oMarkPosColl.Add oMarkPointPos1
        oMarkPosColl.Add oMarkPointPos2

        Set oValidMarkPointColl = New Collection

        'Check each mark point: Do we need ML for this mark point?
        For iIndex = 1 To oMarkPosColl.Count
            Set oPosObj = oMarkPosColl.Item(iIndex)

            'Create normal to lateral port in oMarkPosObj
            Dim oLateralSufaceBody As IJSurfaceBody
            Dim oLateralNormalVector As IJDVector
            Set oPortGeometry = oLateralPort.Geometry
            Set oLateralSufaceBody = oPortGeometry
            oLateralSufaceBody.GetNormalFromPosition oPosObj, oLateralNormalVector

            'create line from normal
            oPosObj.Get xStart, yStart, zStart
            oLateralNormalVector.Get xEnd, yEnd, zEnd
            Dim oLateralNormalLine As Line3d
            Set oLateralNormalLine = New IngrGeom3D.Line3d
            endX = xStart + (xEnd * 100)
            endY = yStart + (yEnd * 100)
            endZ = zStart + (zEnd * 100)
             
            oLateralNormalLine.DefineBy2Points xStart, yStart, zStart, endX, endY, endZ

            'get WB of oLateralNormalLine
            Dim oLateralNormalLineWB As IJWireBody
            Dim oLateralNormalLineCS As New ComplexString3d
            oLateralNormalLineCS.AddCurve oLateralNormalLine, False
            Set oLateralNormalLineWB = m_oMfgRuleHelper.ComplexStringToWireBody(oLateralNormalLineCS)

            'Prepare input arguments for GetCommonGeometry()
            Set oLineGeometry = oLateralNormalLineWB
            If oPlateConnByLateralPort Is oThisPart Then
                Set oTestedPortGeometry = oConnPartPort.Geometry
            Else 'this->face, conn->lateral
                Set oTestedPortGeometry = oThisPartPort.Geometry
            End If
            
            'check if there is a common geom between oLateralNormalLine and oThisPartSuface
            Set oIntersectionWB = m_oMfgRuleHelper.GetCommonGeometry(oLineGeometry, oTestedPortGeometry)

            'CleanUp
            Set oPosObj = Nothing
            Set oLineGeometry = Nothing
            Set oPortGeometry = Nothing
            Set oLateralNormalLineWB = Nothing
            Set oLateralNormalLineCS = Nothing
            Set oLateralNormalLine = Nothing
            Set oLateralNormalVector = Nothing
            Set oLateralSufaceBody = Nothing
            Set oTestedPortGeometry = Nothing

            'take another mark point in processing if there is no common geometry
            If Not oIntersectionWB Is Nothing Then
                oValidMarkPointColl.Add oMarkPosColl.Item(iIndex)
            'Else Do nothing
            End If

            Set oIntersectionWB = Nothing
NextPlateIndex:
    Next iIndex

    'Prepare to return the collection of valid mark points
    Set oMarkPosColl = oValidMarkPointColl
    End If




    '*********************** Profile/Plate case *******************************************
    If (TypeOf oConnPart Is IJProfilePart And TypeOf oThisPart Is IJPlatePart) Or _
    (TypeOf oThisPart Is IJProfilePart And TypeOf oConnPart Is IJPlatePart) Then
        'Wrap profile object:

        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Dim oProfilePartSupport As IJProfilePartSupport
        Dim oPlateWrapper As New StructDetailObjects.PlatePart
        
        Dim oUpsidePort As IJPort
        Dim dThickness As Double

        If TypeOf oThisPart Is IJProfilePart Then
            Set oPartSupport.Part = oThisPart
            Set oPlateWrapper.object = oConnPart
            
            Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
            Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
            Set oProfileWrapper.object = oThisPart
            Set oUpsidePort = oProfileWrapper.GetSurfacePort(UpSide)
            dThickness = 0.3
        End If

        If TypeOf oConnPart Is IJProfilePart Then
            Set oPartSupport.Part = oConnPart
            Set oPlateWrapper.object = oThisPart
            
            Dim oMfgPlateWrapper As MfgRuleHelpers.PlatePartHlpr
            Set oMfgPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
            Set oMfgPlateWrapper.object = oThisPart
            Set oUpsidePort = oMfgPlateWrapper.GetSurfacePort(UpSide)
            dThickness = oPlateWrapper.PlateThickness + 0.1
        End If

        'Regard EdgeReinforcementProfilePart as T-connected unconditionally.
        If Not TypeOf oThisPart Is IJProfileER Then
            'Check if connection is T type
            oConnType = 0
            bIsBaseOfTee = False
            oConnType = oPlateWrapper.GetGetConnectionTypeForContour(oConnectionData.AppConnection, bIsBaseOfTee)
    
            'If Plate/Profile are not T - connected then exit
            If Not (oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsBaseOfTee = True) Then
                Exit Function
            End If
        End If
        Set oPlateWrapper = Nothing


        Set oProfilePartSupport = oPartSupport

        Set oConnPartPort = oConnectionData.ToConnectedPort
        Set oThisPartPort = oConnectionData.ConnectingPort

        Dim oLandCurveWB As IJWireBody
        Dim oMfgGeomCol3d As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

        'Set oLandCurveWB = oMfgHelper.GetLandingCurve
        Dim oVector As IJDVector
        oProfilePartSupport.GetProfilePartLandingCurve oLandCurveWB, oVector, 1

        'Get positions of LCs two ends
        oLandCurveWB.GetEndPoints oEndPos1, oEndPos2 ', oStartPointDir, oEndPointDir
        Set oCurve = oLandCurveWB

        'find two mark points at 'FittingMarkDistance' from both ends
        Dim oLandingCurvePos1 As IJDPosition, oLandingCurvePos2 As IJDPosition
        
        Set oLandingCurvePos1 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos1, FittingMarkDistance, oEndPos2)
        Dim dummy As IJDVector
        Set oMarkPointPos1 = m_oMfgRuleHelper.ProjectPointOnSurface(oLandingCurvePos1, oUpsidePort.Geometry, dummy)
        Dim dDist As Double
        dDist = oMarkPointPos1.DistPt(oLandingCurvePos1)
        If dDist < dThickness Then
            oMarkPosColl.Add oMarkPointPos1
        End If
        
        Set oLandingCurvePos2 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos2, FittingMarkDistance, oEndPos1)
        Set oMarkPointPos2 = m_oMfgRuleHelper.ProjectPointOnSurface(oLandingCurvePos2, oUpsidePort.Geometry, dummy)
        dDist = oMarkPointPos2.DistPt(oLandingCurvePos2)
        If dDist < dThickness Then
            oMarkPosColl.Add oMarkPointPos2
        End If
        
    End If

    Set GetFittingMarkPositions = oMarkPosColl
    
CleanUp:
    Set oMarkPointPos1 = Nothing
    Set oMarkPointPos2 = Nothing
    Set oCurve = Nothing
    Set oLandCurveWB = Nothing
    Set oEndPos1 = Nothing
    Set oEndPos2 = Nothing
    Set oVector = Nothing
    Set oThisPartPort = Nothing
    Set oConnPartPort = Nothing
    Set oCommonGeom = Nothing
    Set oFaceSuface = Nothing
    Set oLateralSurface = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ViewUnk(oUnkObj, METHOD)
Const sMETHOD = "viewUnk"
On Error GoTo ErrorHandler

    Dim tmp As IUnknown
    If oUnkObj Is Nothing Then
        Dim lErrNumber As Long
        lErrNumber = LogMessage(Err, MODULE, sMETHOD, "Invalid input to Viewer")
        Exit Function
    End If
    If TypeOf oUnkObj Is IJComplexString Then
        Dim elem As IJElements
        Dim oCS As IJComplexString
        Set oCS = oUnkObj
        oCS.GetCurves elem
        Set tmp = elem.Item(1)
    Else
        Set tmp = oUnkObj
    End If
        
' COMMENTING THE BELOW LINES (that use IJHiliter) EFFECTIVELY MAKES THIS
' ENTIRE FUNCTION A DUMMY.  DID A PERFUNCTORY GREP AND DID NOT FIND ANYONE
' CALLING IT (although it has been duplicated in MfgRuleHelpers\Helper.cls).
' NEED TO COMMENT IT IN ORDER TO AVOID CLIENT DEPENDENCIES.

        
'    Dim m_oHiliter As IJHiliter
'    Set m_oHiliter = New IMSHiliteObjs.Hiliter
'    m_oHiliter.Color = vbRed
'    m_oHiliter.Weight = 2
'    m_oHiliter.Elements.Add oUnkObj
'    'MsgBox method
'
'    m_oHiliter.Elements.Clear
    Set tmp = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function DefineMarkingSide(ByVal Part As Object, ByVal UpSide As GSCADMfgRulesDefinitions.enumPlateSide, ByVal ConnectingSide As GSCADMfgRulesDefinitions.enumPlateSide) As GSCADMfgRulesDefinitions.enumPlateSide
' Used for determining the correct marking side of a marking line in the parts monitor
' General concept is:
' If Upside and connectingside are the same then it should be considered MARKING
' If they are not the same it should be set to ANTI-MARKINGSIDE

Const sMETHOD = "DefineMarkingSide"
On Error GoTo ErrorHandler

'Create the SD plate Wrapper and initialize it
Dim oSDPlateWrapper As StructDetailObjects.PlatePart
Set oSDPlateWrapper = New StructDetailObjects.PlatePart
Set oSDPlateWrapper.object = Part

If oSDPlateWrapper.plateType = DeckPlate And _
   oSDPlateWrapper.ThicknessDirection = AboveDir Then
'TR 57446 need to reverse to marking side in order to get the correct marking directions outputted

    If ConnectingSide = UpSide Then
        DefineMarkingSide = OffsetSide
    Else
        DefineMarkingSide = BaseSide
    End If

Else

    If ConnectingSide = UpSide Then
        DefineMarkingSide = BaseSide
    Else
        DefineMarkingSide = OffsetSide
    End If

End If
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Public Function TrimTeeWireForSplitPCs(oPC As Object, oTeeWire As IJWireBody) As IJWireBody
    Const METHOD = "TrimTeeWireForSplitPCs"
    On Error GoTo ErrorHandler
    
    Dim iIndex As Integer
    Dim dDist As Double, dMinDist1 As Double, dMinDist2 As Double
    Dim dX As Double, dY As Double, dZ As Double
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollectionHelper As IMSRelation.DCollectionHelper
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oCS As IJComplexString
    Dim oCSColl As IJElements, oWireColl As IJElements
    Dim oPointColl As New Collection
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition, oPos As IJDPosition, oProjectedPos As IJDPosition
    Dim oClosestPos1 As IJDPosition, oClosestPos2 As IJDPosition
    Dim oToolBox As IJDTopologyToolBox
    Dim oProfile2dHelper As IJDProfile2dHelper
    Dim oSGOWireBodyUtilities As IJSGOWireBodyUtilities
    Dim oPoint As IJPoint
    Dim oPCWire As IJWireBody, oWire As IJWireBody, oBoundedWire As IJWireBody
    
    ' Create Helper classes
    Set oSGOWireBodyUtilities = New SGOWireBodyUtilities
    Set oToolBox = New IMSModelGeomOps.DGeomOpsToolBox
    Set oProfile2dHelper = New DProfile2dHelper
    Set oMfgMGHelper = New MfgMGHelper
    Set oWireColl = New JObjectCollection
    Set oPos = New DPosition
    
    ' Initialize output same as input
    Set TrimTeeWireForSplitPCs = oTeeWire
    
    'Check if this PC is child of Split operation
    Set oRelationHelper = oPC
    
    ' IJStructGeometry = 6034AD40-FA0B-11D1-B2FD-080036024603
    Set oCollectionHelper = oRelationHelper.CollectionRelations("{6034AD40-FA0B-11D1-B2FD-080036024603}", "StructOperation_RSLT1_DEST")
    
    ' There is nothing to do, if the PC is not result of a split operation.
    If oCollectionHelper.Count = 0 Then
        Exit Function
    End If
    
    ' Get the split point operators
    Set oRelationHelper = oCollectionHelper.Item(1)

    ' IJStructSplit = 0A31DCF2-45EB-11D5-8126-00105AE5AAE5
    Set oCollectionHelper = oRelationHelper.CollectionRelations("{0A31DCF2-45EB-11D5-8126-00105AE5AAE5}", "StructSplit_OPER1_ORIG")
    
    ' Error check: Check it has any split points
    If oCollectionHelper.Count = 0 Then
        Exit Function
    End If
     
    ' Convert input TeeWire into set of complex strings. This is done to remove self overlaps.
    oMfgMGHelper.WireBodyToComplexStrings oTeeWire, oCSColl
       
    ' Convert each CS to a wirebody and Create a collection of all wires.
    For iIndex = 1 To oCSColl.Count
        Dim oWireLump As IJWireBody
        oMfgMGHelper.ComplexStringToWireBody oCSColl.Item(iIndex), oWireLump
        oWireColl.Add oWireLump
    Next iIndex
    
    ' Merge all the wire lumps to a single wirebody
    Set oWire = oProfile2dHelper.MergeWireBodies(oWireColl, Nothing)
            
    ' Collect all the split points in a collection.
    For iIndex = 1 To oCollectionHelper.Count
        Set oPoint = oCollectionHelper.Item(iIndex)
        oPoint.GetPoint dX, dY, dZ
        oPos.Set dX, dY, dZ
        oToolBox.GetNearestPointOnWireBodyFromPoint oWire, oPos, Nothing, oProjectedPos
        oPointColl.Add oProjectedPos
    Next iIndex
    
    ' Add wirebody's end points into the collection.
    oWire.GetEndPoints oStartPos, oEndPos
    oPointColl.Add oStartPos
    oPointColl.Add oEndPos
    
    ' Get the two closest split points for the input PC geometry.
    Set oPCWire = oPC
    dMinDist1 = 1E+30
    dMinDist2 = 1E+30
    For iIndex = 1 To oPointColl.Count
        Set oPos = oPointColl.Item(iIndex)
        oToolBox.GetNearestPointOnWireBodyFromPoint oPCWire, oPos, Nothing, oProjectedPos
        dDist = oProjectedPos.DistPt(oPos)
        If dDist <= dMinDist1 Then
            dMinDist2 = dMinDist1
            dMinDist1 = dDist
            Set oClosestPos2 = oClosestPos1
            Set oClosestPos1 = oPos
        ElseIf dDist < dMinDist2 Then
            dMinDist2 = dDist
            Set oClosestPos2 = oPos
        End If
    Next iIndex
        
    ' Bound the TeeWire by the two closest points
    oSGOWireBodyUtilities.BoundWireByTwoPoints oWire, oClosestPos1, oClosestPos2, oBoundedWire
        
    ' Return the splitted wire as output
    Set TrimTeeWireForSplitPCs = oBoundedWire
    
CleanUp:
    Set oRelationHelper = Nothing
    Set oCollectionHelper = Nothing
    Set oMfgMGHelper = Nothing
    Set oCS = Nothing
    Set oCSColl = Nothing
    Set oWireColl = Nothing
    Set oPointColl = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oPos = Nothing
    Set oProjectedPos = Nothing
    Set oClosestPos1 = Nothing
    Set oClosestPos2 = Nothing
    Set oToolBox = Nothing
    Set oProfile2dHelper = Nothing
    Set oSGOWireBodyUtilities = Nothing
    Set oPoint = Nothing
    Set oPCWire = Nothing
    Set oWire = Nothing
    Set oBoundedWire = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function



