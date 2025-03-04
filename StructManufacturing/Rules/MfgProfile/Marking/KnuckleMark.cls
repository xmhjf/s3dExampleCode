VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Knuckle"
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

Private Const MODULE = "MfgProfileMarking.Knuckle"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "Knuckle: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    'Create the 3d collection
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)

    CreateAPSProfileMarkings STRMFG_KNUCKLE_MARK, ReferenceObjColl, oGeomCol3d

    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    If bSelectiveRecompute Then
        Exit Function
    End If

    ' (Try to) create fitting marks for extend knuckles regardless of whether the profile
    ' actually has extend knuckles or not (need fitting marks for extend "curve" cases also).
    '                                   Per Ninad for 213958
    If Not TypeOf Part Is IJProfileER Then
        ' Don't do this for edge reinforcements
        CreateExtendKnuckleFittingMark Part, UpSide, ReferenceObjColl, oGeomCol3d
    End If

    Dim oGeom3d As IJMfgGeom3d
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMarkingInfo As IJMarkingInfo
    Dim oMoniker As IMoniker
    Dim oCS As IJComplexString
    Dim oLines3d As ILines3d
    Dim oLine As Line3d
    Dim oGeomFac As New GeometryFactory
    Dim oCrvElemets As IJElements
    Dim oComplexStrings3d As IComplexStrings3d
    Dim lGeomCount As Long
    Dim eSectionType As ProfileSectionType

    Set oLines3d = oGeomFac.Lines3d
    Set oCrvElemets = New JObjectCollection
    Set oComplexStrings3d = oGeomFac.ComplexStrings3d

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = Part

    Dim oMountingPort As IJPort
    Set oMountingPort = oSDProfileWrapper.MountingFacePort
    Dim oProfSurface As IJSurfaceBody
    Set oProfSurface = oMountingPort.Geometry

    'Create the SD profile Wrapper and initialize it
    Dim oSDProfilePartSupport As IJProfilePartSupport
    Dim oPartSupp As IJPartSupport

    'Create the SD profile Wrapper and initialize it
    If TypeOf Part Is IJStiffenerPart Then
        Set oSDProfilePartSupport = New ProfilePartSupport
        Set oPartSupp = oSDProfilePartSupport
        Set oPartSupp.Part = Part
    ElseIf TypeOf Part Is IJBeamPart Then
        GoTo CleanUp
    End If

    Dim oColKnuckles As Collection, oColKnuckleGeom As Collection, oColKnuckleAngles As Collection
    oSDProfilePartSupport.GetBendProfileKnuckleMfgData UpSide, oColKnuckles, oColKnuckleGeom, oColKnuckleAngles

    '================================================================='
    Dim oSupp2 As IJPartSupport
    Dim oStiffSys As IJSystem

    Set oSupp2 = New PartSupport
    Set oSupp2.Part = Part
    oSupp2.IsSystemDerivedPart oStiffSys, True

    Dim pIJProfileAttributes As IJProfileAttributes
    Dim oProfileUtils As IJProfileDefinition
    Dim oKnuckleCollection As IJElements
    Dim oMfgKnuckle As IJProfileKnuckleMfg
    Dim lCount As Integer, kk As Integer

    Set oProfileUtils = New ProfileUtils
    Set pIJProfileAttributes = oProfileUtils
    Set oKnuckleCollection = pIJProfileAttributes.GetKnucklesOnProfile(oStiffSys)
    lCount = oKnuckleCollection.Count
    If lCount > 0 Then
        For kk = 1 To lCount
            If TypeOf oKnuckleCollection.Item(kk) Is IJProfileBendKnuckle Then
                Set oMfgKnuckle = oKnuckleCollection.Item(kk)
                GoTo BendKnuckle
            End If
        Next kk
    End If

    '================================================================='
BendKnuckle:

    If Not oMfgKnuckle Is Nothing Then

        If oMfgKnuckle.ManufacturingMethod = pkmmBend Then

            Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
            Set oProfileWrapper.object = Part

            '*** IMPLEMENTING BEND KNUCKLE MARKS ON PROFILE ***'

            'Get Stiffened plate system
            Dim oPlateSystem As IJPlateSystem
            Dim oObj As Object
            Dim oPlate As IJPlate
            Set oPlate = oSDProfilePartSupport.StiffenedPlate
            Set oPlateSystem = m_oMfgRuleHelper.GetTopMostParentSystem(oPlate)

            Dim oPlateUtils As IJPlateAttributes
            Set oPlateUtils = New PlateUtils

            'Get Knuckles on Plate
            Dim oKnuckleElements As IJElements
            Set oKnuckleElements = oPlateUtils.GetKnucklesOnPlate(oPlateSystem)

            If oKnuckleElements Is Nothing Or oKnuckleElements.Count = 0 Then
                GoTo CheckKnuckleMarks
            End If

            'Get Profile Landing Curve
            Dim oLandingWB As IJWireBody
            Set oLandingWB = oProfileWrapper.GetLandingCurve
            Dim oLandingCS As IJComplexString
            Set oLandingCS = m_oMfgRuleHelper.WireBodyToComplexString(oLandingWB)

            Dim oMB As IJDModelBody
            Set oMB = oLandingWB
            Set oMB = oProfSurface

            Dim ii As Integer
            Dim dParam As Double
            Dim dX1 As Double, dY1 As Double, dZ1 As Double, vTanX1 As Double, vTanY1 As Double, vTanZ1 As Double, vDummy As Double
            Dim dX2 As Double, dY2 As Double, dZ2 As Double, vTanX2 As Double, vTanY2 As Double, vTanZ2 As Double
            Dim oMarkBasePos As New DPosition

            Dim lFaceId As Long
            Dim oSurfaceBody As IJSurfaceBody
            Dim oMGHelper As New MfgMGHelper
            Dim oProjPos As IJDPosition
            Dim oProjCS As IJComplexString
            Dim oMfgGeomHelper As MfgGeomHelper
            Set oMfgGeomHelper = New MfgGeomHelper

            'Loop through Knuckles & check which is closest to Landing Curve
            '..if within tolerance, create the mark at the point
            For ii = 1 To oKnuckleElements.Count

                Dim oKnuckleWire As IJWireBody
                Set oKnuckleWire = oKnuckleElements.Item(ii)

                Dim oMB1 As IJDModelBody
                Set oMB1 = oKnuckleWire

                Dim oPt1 As IJDPosition
                Dim oPt2 As IJDPosition
                Dim dMinDist As Double

                oMB.GetMinimumDistance oMB1, oPt1, oPt2, dMinDist

                'if distance < thickness, that means the point is valid
                If Abs(dMinDist) < oPlate.thickness Then

                    Dim oVecAtEnd As New DVector
                    Dim oVecAtStart As New DVector
                    Dim Markvector As New DVector
                    Dim oLandingCurve As IJCurve
                    Dim oWBStartPos As New DPosition, oWBEndPos As New DPosition
                    Dim oWBStartVec As IJDVector, oWBEndVec As IJDVector

                    oLandingCS.GetCurve 1, oLandingCurve

                    Set oLandingCurve = oLandingCS

                    'oPt1 is point between intersection of Plate Knucle and Profile Landing curve
                    oLandingCurve.Parameter oPt1.x, oPt1.y, oPt1.z, dParam

                    'Find two points on either side of oPt1 for evaluation (& get vectors)
                    Dim oPointOnLeft As IJDPosition, oPointOnRight As IJDPosition
                    Dim oRefPos As New DPosition

                    oLandingWB.GetEndPoints oWBStartPos, oWBEndPos, oWBStartVec, oWBEndVec

                    'Get the Knuckle Angle
                    Dim oAngleEndVec As IJDVector, oAngleStartVec As IJDVector
                    Set oAngleEndVec = oWBEndPos.Subtract(oPt1)
                    Set oAngleStartVec = oWBStartPos.Subtract(oPt1)
                    Dim dAngle As Double

                    dAngle = oAngleStartVec.Angle(oAngleEndVec, oAngleStartVec.Cross(oAngleEndVec))
                    dAngle = oAngleEndVec.Angle(oAngleStartVec, oAngleStartVec.Cross(oAngleEndVec))

                    If Abs(dAngle) > 3.14159265358979 Then
                        dAngle = 2 * 3.14159265358979 - dAngle
                    End If

                    'For Flat Bar, Half Round Bar, Round Bar & Square Bar

                    Dim dStartParam As Double, dEndParam As Double
                    Set oPointOnLeft = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oLandingWB, oPt1, 0.1, oWBStartPos)
                    Set oPointOnRight = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oLandingWB, oPt1, 0.1, oWBEndPos)

                    oLandingCurve.Parameter oPointOnLeft.x, oPointOnLeft.y, oPointOnLeft.z, dStartParam
                    oLandingCurve.Evaluate dStartParam, dX1, dY1, dZ1, vTanX1, vTanY1, vTanZ1, vDummy, vDummy, vDummy

                    oLandingCurve.Parameter oPointOnRight.x, oPointOnRight.y, oPointOnRight.z, dEndParam
                    oLandingCurve.Evaluate dEndParam, dX2, dY2, dZ2, vTanX2, vTanY2, vTanZ2, vDummy, vDummy, vDummy

                    oVecAtStart.Set -vTanX1, -vTanY1, -vTanZ1
                    oVecAtEnd.Set vTanX2, vTanY2, vTanZ2

                    Markvector.Set (oVecAtEnd.x + oVecAtStart.x) / 2, (oVecAtEnd.y + oVecAtStart.y) / 2, (oVecAtEnd.z + oVecAtStart.z) / 2
                    Markvector.length = oPlate.thickness / 2

                    '*** Construct Surface Normal ***'
                    Dim oLeftPtNormal As IJDVector, oRightPtNormal As IJDVector
                    Dim oMiddleNormal As IJDVector

                    oProfSurface.GetNormalFromPosition oPointOnLeft, oLeftPtNormal
                    oProfSurface.GetNormalFromPosition oPointOnRight, oRightPtNormal

                    Set oMiddleNormal = oLeftPtNormal.Add(oRightPtNormal)
                    'Reverse since it is normal at profile bottom surface
                    oMiddleNormal.length = oMiddleNormal.length * -1

                    Set oMarkBasePos = oPt1

                    'Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x - Markvector.x, oMarkBasePos.y - Markvector.y, oMarkBasePos.z - Markvector.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)

                    oMiddleNormal.length = oSDProfileWrapper.WebLength / 2

                    eSectionType = oSDProfilePartSupport.SectionType

                    Dim lKnUpside As Long

                    If eSectionType = Flat_Bar Or eSectionType = Half_Round Or eSectionType = S_Channel Then
                        'If Dot Product < 0, Mark starts at oPt1, else starts at other end
                        If oMiddleNormal.Dot(Markvector) < 0 Then
                            Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x, oMarkBasePos.y, oMarkBasePos.z, oMarkBasePos.x + oMiddleNormal.x, oMarkBasePos.y + oMiddleNormal.y, oMarkBasePos.z + oMiddleNormal.z)
                        Else
                            Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x + oMiddleNormal.x, oMarkBasePos.y + oMiddleNormal.y, oMarkBasePos.z + oMiddleNormal.z, oMarkBasePos.x + 2 * oMiddleNormal.x, oMarkBasePos.y + 2 * oMiddleNormal.y, oMarkBasePos.z + 2 * oMiddleNormal.z)
                        End If

                        Set oCrvElemets = New JObjectCollection
                        oCrvElemets.Add oLine

                        Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

                        '*** Project Line on Profile Surface ***'
                        Set oSurfaceBody = oSDProfileWrapper.SubPortBeforeTrim(JXSEC_WEB_LEFT).Geometry

                        Set oMB1 = oSurfaceBody
                        oMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS

                        lFaceId = UpSide
                        lKnUpside = UpSide

                    ElseIf eSectionType = Angle Or eSectionType = Bulb_Angle Or eSectionType = Bulb_Tee Or _
                        eSectionType = Bulb_Type Or eSectionType = Fab_Angle Then

                        Dim sKnUpside As String
                        'If Dot Product < 0, Mark starts at oPt1, else starts at other end
                        'For Angle Profiles, If starts @ oPt1 means Anti_Marking
                        If oMiddleNormal.Dot(Markvector) < 0 Then
                            sKnUpside = "anti_marking"
                            lKnUpside = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
                        Else
                            sKnUpside = "marking"
                            lKnUpside = JXSEC_TOP
                        End If

                        Dim oKnStartPos As New DPosition, oKnEndPos As New DPosition
                        Dim oKnStartDir As IJDVector, oKnEndDir As IJDVector

                        oKnuckleWire.GetEndPoints oKnStartPos, oKnEndPos, oKnStartDir, oKnEndDir

                        Set oLine = oLines3d.CreateBy2Points(Nothing, oKnStartPos.x, oKnStartPos.y, oKnStartPos.z, oKnEndPos.x, oKnEndPos.y, oKnEndPos.z)

                        Set oCrvElemets = New JObjectCollection
                        oCrvElemets.Add oLine

                        Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

                        '*** Project Line on Profile Surface ***'
                        Set oSurfaceBody = oSDProfileWrapper.SubPortBeforeTrim(JXSEC_TOP).Geometry

                        Set oMB1 = oSurfaceBody

                        oMGHelper.ProjectComplexStringToSurface oCS, oSurfaceBody, Nothing, oProjCS
                        Set oCS = oProjCS

                        lFaceId = JXSEC_TOP

                    End If
                End If

                '**************************'
                oCrvElemets.Clear
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

                'Set the marking side
                'oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
                'Set the marking side as upside always
                'oSystemMark.SetMarkingSide UpSide
                oSystemMark.SetMarkingSide lKnUpside

                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

                'oMarkingInfo.Name = oSDProfileWrapper.Name
                oMarkingInfo.thickness = oSDProfileWrapper.WebThickness

                oMarkingInfo.FittingAngle = dAngle
                'oMarkingInfo.direction = lKnUpside

                '            oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oVector)
                '            oMarkingInfo.ThicknessDirection = oVector

                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.FaceId = lFaceId
                oGeom3d.TrimToBoundaries = True
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_KNUCKLE_MARK
                oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "KnuckleFit"

                '            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                '            oGeom3d.PutMoniker oMoniker

                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry oGeomCol3d.Getcount + 1, oGeom3d
                        'lGeomCount = lGeomCount + 1
            Next ii
        End If

    End If
    '**************************************************'

CheckKnuckleMarks:
    'The below function gives the information about Ignore Knuckles on the Profiles
    GetIgnoreKnuckleData Part, UpSide, oColKnuckleGeom, oColKnuckleAngles

    '*** IF THE PROFILE HAS KNUCKLES ***'
    If Not (oColKnuckleGeom Is Nothing) Then
        Dim i As Integer
        For i = 1 To oColKnuckleGeom.Count
            If TypeOf oColKnuckleGeom.Item(i) Is IJWireBody Then
                Dim oWire As IJWireBody
                Set oWire = oColKnuckleGeom.Item(i)
                Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWire)
            Else
                Set oCS = oColKnuckleGeom.Item(i)
            End If

            Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_KNUCKLE_MARK

            'Create a SystemMark object to store additional information
            Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
            Set oObjSystemMark = oSystemMark

            'Set the marking side
            oSystemMark.SetMarkingSide UpSide
            oSystemMark.Set3dGeometry oGeom3d

            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark
            oMarkingInfo.name = "KNUCKLE-" & i
            ' convert to radians

            If (oColKnuckleAngles.Item(i) > 180) Then
                oMarkingInfo.direction = "up"
                oMarkingInfo.FittingAngle = (oColKnuckleAngles.Item(i) - 180) * 3.14159265358979 / 180#
                oSystemMark.SetMarkingSide UpSide
                oGeom3d.FaceId = UpSide
            Else
                oMarkingInfo.direction = "down"
                oMarkingInfo.FittingAngle = oColKnuckleAngles.Item(i) * 3.14159265358979 / 180#
                If (UpSide Mod 2) = 1 Then
                    oSystemMark.SetMarkingSide UpSide + 1
                    oGeom3d.FaceId = UpSide + 1
                Else
                    oSystemMark.SetMarkingSide UpSide - 1
                    oGeom3d.FaceId = UpSide - 1
                End If
            End If

            oGeomCol3d.AddGeometry lGeomCount, oGeom3d

            Set oGeom3d = Nothing
            Set oSystemMark = Nothing
            Set oObjSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oMoniker = Nothing
        Next i
    End If
    '***********************************'

CleanUp:
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    Set oSDProfilePartSupport = Nothing
    Set oColKnuckles = Nothing
    Set oColKnuckleGeom = Nothing
    Set oColKnuckleAngles = Nothing
    Set oGeomCol3d = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2006, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function CreateExtendKnuckleFittingMark(ByVal Part As Object, UpSide As Long, ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                                ByVal oGeomCol3d As IJMfgGeomCol3d)

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = Part

    Dim oProfileWLSurface As IJSurfaceBody
    Set oProfileWLSurface = oSDProfileWrapper.SubPort(JXSEC_WEB_LEFT).Geometry

    Dim oWireBodyCS As ComplexString3d
    Dim nCrvCount As Long, nCrvIdx As Long

    Dim oCurve As IJCurve
    Dim oNextCurve As IJCurve, oPrevCurve As IJCurve
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    Dim dStart2X As Double, dStart2Y As Double, dStart2Z As Double
    Dim dEnd2X As Double, dEnd2Y As Double, dEnd2Z As Double
    Dim dStartPar As Double, dEndPar As Double, dOffsetPar As Double
    Dim vTanX As Double, vTanY As Double, vTanZ As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim oLines3d As ILines3d
    Dim oLine As Line3d
    Dim oCrvElemets As IJElements
    Set oCrvElemets = New JObjectCollection
    Dim oComplexStrings3d As IComplexStrings3d
    Dim oCS As ComplexString3d
    Dim oProjCS As IJComplexString

    Dim oResourceManager As Object
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo

    Dim oVector As IJDVector
    Dim oGeom3d As IJMfgGeom3d

    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport.Part = Part

    Dim oConObjsCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection, oConnCol As Collection
    'Set oConObjsCol = oSDPlateWrapper.ConnectedObjects
    'Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjsCol, oConnCol, oThisPortCol, oOtherPortCol

    Dim oStiffenedPlate As Object
    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, False

    Dim ii As Integer

    Dim oConnectionData As ConnectionData

    Dim oGeomFac As New GeometryFactory
    Dim oMGHelper As New MfgMGHelper

    For ii = 1 To oConObjsCol.Count

        If TypeOf oConObjsCol.Item(ii) Is IJPlatePart Then

            Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper.object = oConObjsCol.Item(ii)

            Dim oSys As IJSystem
            Set oSys = oPlateWrapper.GetRootSystem

            'If the connected plate is the base/stiffened plate
            If oSys Is oStiffenedPlate Then

                Dim oPort As IJPort
                Set oPort = oOtherPortCol.Item(ii)

                Dim oCommonWB As IJWireBody
                Set oCommonWB = m_oMfgRuleHelper.GetCommonGeometry(oProfileWLSurface, oPort.Geometry)

                GoTo CreateMark
            Else
                GoTo NextItem
            End If
        Else
            GoTo NextItem
        End If

NextItem:
    Next

CreateMark:

    Set oWireBodyCS = m_oMfgRuleHelper.WireBodyToComplexString(oCommonWB)
    nCrvCount = oWireBodyCS.CurveCount

    For nCrvIdx = 1 To nCrvCount

        On Error GoTo NextERMark

        If nCrvIdx = nCrvCount Then GoTo CleanUp

        oWireBodyCS.GetCurve nCrvIdx, oCurve
        oWireBodyCS.GetCurve nCrvIdx + 1, oNextCurve

        'If this is a Curve, DO NOT place Knuckle Mark
        If TypeOf oCurve Is IJArc Or TypeOf oNextCurve Is IJArc Then
            GoTo NextERMark
        End If

        If Not TypeOf oCurve Is IJLine Or Not TypeOf oNextCurve Is IJLine Then
            GoTo NextERMark
        End If

        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        oNextCurve.EndPoints dStart2X, dStart2Y, dStart2Z, dEnd2X, dEnd2Y, dEnd2Z

        '*** New Implementation ***'
        Dim oEndPtOf1st As New DPosition
        Dim oStartPtOf2nd As New DPosition

        oEndPtOf1st.Set dEndX, dEndY, dEndZ
        oCurve.Parameter dEndX, dEndY, dEndZ, dEndPar
        oCurve.Evaluate dEndPar, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

        Dim oVecAtEnd As New DVector
        Dim oVecAtStart As New DVector
        Dim Markvector As New DVector

        Dim oMarkBasePos As New DPosition

        oVecAtEnd.Set vTanX, vTanY, vTanZ

        oStartPtOf2nd.Set dStart2X, dStart2Y, dStart2Z
        oNextCurve.Parameter dStart2X, dStart2Y, dStart2Z, dEndPar
        oNextCurve.Evaluate dEndPar, dStart2X, dStart2Y, dStart2Z, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

        oVecAtStart.Set -vTanX, -vTanY, -vTanZ

        Markvector.Set (oVecAtEnd.x + oVecAtStart.x) / 2, (oVecAtEnd.y + oVecAtStart.y) / 2, (oVecAtEnd.z + oVecAtStart.z) / 2

        oMarkBasePos.Set dEndX, dEndY, dEndZ

        Markvector.length = ERFittingMarkLength


'                .--------.                  .--------.
'                |         \                /         |
'                |          \   Fitting    /          |
'                |           \   Marks    /           |
'                |            \/________\/            |
'                |            /          \            |
'                |                                    |
'                |                                    |


        Set oLines3d = oGeomFac.Lines3d
        Set oComplexStrings3d = oGeomFac.ComplexStrings3d

        'Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x, oMarkBasePos.y, oMarkBasePos.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)
        Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x - Markvector.x, oMarkBasePos.y - Markvector.y, oMarkBasePos.z - Markvector.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)

        '**************************'
        oCrvElemets.Add oLine

        Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

        oMGHelper.ProjectComplexStringToSurface oCS, oProfileWLSurface, Nothing, oProjCS
        Set oCS = oProjCS

        oCrvElemets.Clear

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

        'Set the marking side
        'oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
        'Set the marking side as upside always
        oSystemMark.SetMarkingSide UpSide

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark

        'oMarkingInfo.Name = oSDProfileWrapper.Name
        oMarkingInfo.thickness = oSDProfileWrapper.WebThickness

        'oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oVector)
        'oMarkingInfo.ThicknessDirection = oVector

        Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
        oGeom3d.FaceId = UpSide
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_FITTING_MARK

        oSystemMark.Set3dGeometry oGeom3d

        oGeomCol3d.AddGeometry oGeomCol3d.Getcount + 1, oGeom3d
        'lGeomCount = lGeomCount + 1
NextERMark:
    Next nCrvIdx

CleanUp:

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Private Sub GetIgnoreKnuckleData
'
' Description:  This SubRoutine gives the information about IgnoreKnuckle Angle &
'               IgnoreKnuckle object as ComplexString
' ***********************************************************************************
Private Sub GetIgnoreKnuckleData(ByVal oPart As Object, ByVal UpSide As Long, ByRef oColKnuckleGeom As Collection, ByRef oColKnuckleAngles As Collection)

    Const METHOD = "GetIgnoreKnuckleData"
    On Error GoTo ErrorHandler

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = oPart

    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport

    Set oProfilePartSupport = New ProfilePartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = oPart

    Dim j As Long

    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Dim oLeafSystem         As Object

    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    oStructDetailHelper.IsPartDerivedFromSystem oPart, oLeafSystem, False

    Dim oSemanticsUtil  As IJProfilePartSemanticsUtil
    Set oSemanticsUtil = New ProfilePartSemanticsUtil

    Dim oIgnoreKnuckles As IJElements
    oSemanticsUtil.GetProfileKnucklesForPart oLeafSystem, pkmmIgnore, oIgnoreKnuckles

    Dim oMGHelper As New MfgMGHelper
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oPart

    'Checking if the Profile has Ignore Knuckles
    If Not oIgnoreKnuckles Is Nothing Then
        'If a Profile doesn't contain Bend Knuckle on it
        If oColKnuckleGeom Is Nothing Then
            Set oColKnuckleGeom = New Collection
        End If

        If oColKnuckleAngles Is Nothing Then
            Set oColKnuckleAngles = New Collection
        End If
        For j = 1 To oIgnoreKnuckles.Count
            Dim oProfKnuckle As IJProfileKnuckle
            Set oProfKnuckle = oIgnoreKnuckles.Item(j)

            'Getting the Knuckle Angle of the Ignore Knuckle
            oColKnuckleAngles.Add oProfKnuckle.Angle * DEGREES_PER_RADIAN

            'Getting the Knuckle Point
            Dim oPoint As IJPoint
            Set oPoint = oProfKnuckle
            Dim x As Double, y As Double, z As Double
            oPoint.GetPoint x, y, z
            Dim oPosition As IJDPosition
            Set oPosition = New DPosition
            oPosition.Set x, y, z

            'Getting the Upside SurfaceBody of the Profile
            Dim oProfileSurfaceBody As IJSurfaceBody
            Set oProfileSurfaceBody = oProfileWrapper.GetSurfacePort(UpSide).Geometry

            'Getting the Orientation Vectors of the Profile at given knuckle point
            Dim oXVector As IJDVector
            Dim oYVector As IJDVector
            Dim oOrigin As IJDPosition
            oProfilePartSupport.GetOrientation oPosition, oXVector, oYVector, oOrigin
            oYVector.length = oSDProfileWrapper.WebLength * 1.05

            'Creating the Mark as ComplexString
            Dim oMarkLine As IJLine
            Set oMarkLine = New Line3d
            oMarkLine.DefineBy2Points oPosition.x - oYVector.x, oPosition.y - oYVector.y, oPosition.z - oYVector.z, oPosition.x + oYVector.x, oPosition.y + oYVector.y, oPosition.z + oYVector.z
            Dim oMarkCS As IJComplexString
            Set oMarkCS = New ComplexString3d
            oMarkCS.AddCurve oMarkLine, False
            Dim oMarkProjCS As IJComplexString

            'Project the line created on the profile part surface
            oMGHelper.ProjectComplexStringToSurface oMarkCS, oProfileSurfaceBody, Nothing, oMarkProjCS
            If oMarkProjCS Is Nothing Then
                Set oMarkProjCS = oMarkCS
            End If

            oColKnuckleGeom.Add oMarkProjCS
            Set oOrigin = Nothing
            Set oXVector = Nothing
            Set oYVector = Nothing
            Set oProfKnuckle = Nothing
            Set oPoint = Nothing
            Set oPosition = Nothing
            Set oMarkCS = Nothing
            Set oMarkProjCS = Nothing
            Set oMarkLine = Nothing
            Set oProfileSurfaceBody = Nothing
        Next j
    End If
    Set oLeafSystem = Nothing
    Set oStructDetailHelper = Nothing
    Set oMGHelper = Nothing
    Set oSemanticsUtil = Nothing
    Set oIgnoreKnuckles = Nothing
    Set oPartSupport = Nothing
    Set oProfilePartSupport = Nothing
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function CreateKnuckleFittingMarks(ByVal Part As Object, UpSide As Long, ByVal ReferenceObjColl As JCmnShp_CollectionAlias, _
                                                ByVal oGeomCol3d As IJMfgGeomCol3d)

    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = Part

    Dim oProfileWLSurface As IJSurfaceBody
    Set oProfileWLSurface = oSDProfileWrapper.SubPort(JXSEC_WEB_LEFT).Geometry

    Dim oWireBodyCS As ComplexString3d
    Dim nCrvCount As Long, nCrvIdx As Long

    Dim oCurve As IJCurve
    Dim oNextCurve As IJCurve, oPrevCurve As IJCurve
    Dim dStartX As Double, dStartY As Double, dStartZ As Double
    Dim dEndX As Double, dEndY As Double, dEndZ As Double
    Dim dStart2X As Double, dStart2Y As Double, dStart2Z As Double
    Dim dEnd2X As Double, dEnd2Y As Double, dEnd2Z As Double
    Dim dStartPar As Double, dEndPar As Double, dOffsetPar As Double
    Dim vTanX As Double, vTanY As Double, vTanZ As Double
    Dim vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim oLines3d As ILines3d
    Dim oLine As Line3d
    Dim oCrvElemets As IJElements
    Set oCrvElemets = New JObjectCollection
    Dim oComplexStrings3d As IComplexStrings3d
    Dim oCS As ComplexString3d
    Dim oProjCS As IJComplexString

    Dim oResourceManager As Object
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo

    Dim oVector As IJDVector
    Dim oGeom3d As IJMfgGeom3d

    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport.Part = Part

    Dim oConObjsCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection, oConnCol As Collection
    'Set oConObjsCol = oSDPlateWrapper.ConnectedObjects
    'Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjsCol, oConnCol, oThisPortCol, oOtherPortCol

    Dim oStiffenedPlate As Object
    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, False

    Dim ii As Integer

    Dim oConnectionData As ConnectionData

    Dim oGeomFac As New GeometryFactory
    Dim oMGHelper As New MfgMGHelper

    For ii = 1 To oConObjsCol.Count

        If TypeOf oConObjsCol.Item(ii) Is IJPlatePart Then

            Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
            Set oPlateWrapper.object = oConObjsCol.Item(ii)

            Dim oSys As IJSystem
            Set oSys = oPlateWrapper.GetRootSystem

            'If the connected plate is the base/stiffened plate
            If oSys Is oStiffenedPlate Then

                Dim oPort As IJPort
                Set oPort = oOtherPortCol.Item(ii)

                Dim oCommonWB As IJWireBody
                Set oCommonWB = m_oMfgRuleHelper.GetCommonGeometry(oProfileWLSurface, oPort.Geometry)

                GoTo CreateMark
            Else
                GoTo NextItem
            End If
        Else
            GoTo NextItem
        End If

NextItem:
    Next

CreateMark:

    Set oWireBodyCS = m_oMfgRuleHelper.WireBodyToComplexString(oCommonWB)
    nCrvCount = oWireBodyCS.CurveCount

    For nCrvIdx = 1 To nCrvCount

        If nCrvIdx = nCrvCount Then GoTo CleanUp

        oWireBodyCS.GetCurve nCrvIdx, oCurve
        oWireBodyCS.GetCurve nCrvIdx + 1, oNextCurve

        'If this is a Curve, DO NOT place Knuckle Mark
        If TypeOf oCurve Is IJArc Or TypeOf oNextCurve Is IJArc Then
            GoTo NextERMark
        End If

        If Not TypeOf oCurve Is IJLine Or Not TypeOf oNextCurve Is IJLine Then
            GoTo NextERMark
        End If

        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        oNextCurve.EndPoints dStart2X, dStart2Y, dStart2Z, dEnd2X, dEnd2Y, dEnd2Z

        '*** New Implementation ***'
        Dim oEndPtOf1st As New DPosition
        Dim oStartPtOf2nd As New DPosition

        oEndPtOf1st.Set dEndX, dEndY, dEndZ
        oCurve.Parameter dEndX, dEndY, dEndZ, dEndPar
        oCurve.Evaluate dEndPar, dEndX, dEndY, dEndZ, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

        Dim oVecAtEnd As New DVector
        Dim oVecAtStart As New DVector
        Dim Markvector As New DVector

        Dim oMarkBasePos As New DPosition

        oVecAtEnd.Set vTanX, vTanY, vTanZ

        oStartPtOf2nd.Set dStart2X, dStart2Y, dStart2Z
        oNextCurve.Parameter dStart2X, dStart2Y, dStart2Z, dEndPar
        oNextCurve.Evaluate dEndPar, dStart2X, dStart2Y, dStart2Z, vTanX, vTanY, vTanZ, vTan2X, vTan2Y, vTan2Z

        oVecAtStart.Set -vTanX, -vTanY, -vTanZ

        Markvector.Set (oVecAtEnd.x + oVecAtStart.x) / 2, (oVecAtEnd.y + oVecAtStart.y) / 2, (oVecAtEnd.z + oVecAtStart.z) / 2

        oMarkBasePos.Set dEndX, dEndY, dEndZ

        Markvector.length = ERFittingMarkLength


'                .--------.                  .--------.
'                |         \                /         |
'                |          \   Fitting    /          |
'                |           \   Marks    /           |
'                |            \/________\/            |
'                |            /          \            |
'                |                                    |
'                |                                    |


        Set oLines3d = oGeomFac.Lines3d
        Set oComplexStrings3d = oGeomFac.ComplexStrings3d

        'Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x, oMarkBasePos.y, oMarkBasePos.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)
        Set oLine = oLines3d.CreateBy2Points(Nothing, oMarkBasePos.x - Markvector.x, oMarkBasePos.y - Markvector.y, oMarkBasePos.z - Markvector.z, oMarkBasePos.x + Markvector.x, oMarkBasePos.y + Markvector.y, oMarkBasePos.z + Markvector.z)

        '**************************'
        oCrvElemets.Add oLine

        Set oCS = oComplexStrings3d.CreateByCurves(Nothing, oCrvElemets)

        oMGHelper.ProjectComplexStringToSurface oCS, oProfileWLSurface, Nothing, oProjCS
        Set oCS = oProjCS

        oCrvElemets.Clear

        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

        'Set the marking side
        'oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
        'Set the marking side as upside always
        oSystemMark.SetMarkingSide UpSide

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark

        'oMarkingInfo.Name = oSDProfileWrapper.Name
        oMarkingInfo.thickness = oSDProfileWrapper.WebThickness

        'oMarkingInfo.direction = m_oMfgRuleHelper.GetDirection(oVector)
        'oMarkingInfo.ThicknessDirection = oVector

        Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
        oGeom3d.FaceId = UpSide
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_FITTING_MARK

        oSystemMark.Set3dGeometry oGeom3d

        oGeomCol3d.AddGeometry oGeomCol3d.Getcount + 1, oGeom3d
        'lGeomCount = lGeomCount + 1
NextERMark:
    Next nCrvIdx

CleanUp:

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function



