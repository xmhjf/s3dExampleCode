VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogConnProfileRemarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: ProfileRemarking
'
' Description:  Example to show the remarking lines due to profiles
'               that are connected to the supported plates, on the pinjig
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.LogConnProfileRemarking"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oReturnCollection As IJElements
    Set oReturnCollection = New JObjectCollection
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionLogical, oPlateColl)

    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PROFILE_TYPE, _
                            ConnectionLogical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl

        Dim i As Long
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            oReturnCollection.Add ConnectedObjColl.Item(i)
        Next
    Next
    
    GetAPSMarkingLines oPinJig, STRMFG_PROFILELOCATION_MARK, oReturnCollection

    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oReturnCollection
    
    Set oReturnCollection = Nothing
    Set oPlateColl = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5018, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsConnectedProfilePart]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport

    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject

    Dim oMfgGeomHelper As New MfgGeomHelper

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionLogical, oPlateColl)

    Dim oRemarkingSurface As IJSurfaceBody
    If (m_dRemarkingSurfaceOffset > 0.001) Then
        Set oRemarkingSurface = oPinJig.RemarkingSurface
    End If
    
    Dim oMfgMGhelper As New MfgMGHelper
    
    ' ******************************* IMPORTANT NOTE ******************************* '
    ' The Pin jig semantic expects Connection Geometry (with proper attribution).    '
    ' So create a MfgGeom3d object (that has both geometry and attribution).         '
    ' The semantic will use these objects to create connection remarking lines and   '
    ' merge them based on MONIKERS and NAMES of these objects.                       '
    ' ******************************* IMPORTANT NOTE ******************************* '

    Dim Iter As Integer
    For Iter = 1 To oPlateColl.Count
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PROFILE_TYPE, ConnectionLogical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl

        Dim i As Long
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            Dim oWireBody As IJWireBody
            If ElemsToRemark.Contains(ConnectedObjColl.Item(i)) And _
               TypeOf ConnectionsColl.Item(i) Is IJWireBody Then
        ' ... get the LC's geometry
                Set oWireBody = ConnectionsColl.Item(i)
            End If

            If Not oWireBody Is Nothing Then
                Dim oNI As IJNamedItem
                Set oNI = ConnectedObjColl.Item(i)

                Dim RemarkingLineName As String
                If Not oNI Is Nothing Then
                    ' Decorate your name here (e.g., prefix with "Deck" etc.)
                    RemarkingLineName = oNI.Name
                    ' Decorate your name here (e.g., suffix with "SideA" etc.)
                End If

                Dim oCScoll As IJElements
                Set oCScoll = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBody)

                Dim oCS As IJComplexString
                For Each oCS In oCScoll

                    Dim oResultCS As IJComplexString
                
                    ' remarking surface is obtained only when the offset is > 0.001
                    If (Not oRemarkingSurface Is Nothing) And m_dRemarkingSurfaceOffset > 0.001 Then
                        ' Need to use the profile orientation(up vector) for computing the projection vector
                        On Error Resume Next
                        Dim oProfileVVec As IJDVector
                        Set oProfileVVec = GetProfileOrientationVector(oCS, ConnectedObjColl.Item(i))
                        
                        oMfgMGHelper.ProjectComplexStringToSurface oCS, oRemarkingSurface, oProfileVVec, oResultCS
                        On Error GoTo ErrorHandler
                        
                        If oResultCS Is Nothing Then
                            StrMfgLogError Err, MODULE, METHOD, "Failed to project complexstring onto remarking surface", , , , "RULES"
                            Set oResultCS = oCS
                        End If
                    Else
                        Set oResultCS = oCS
                    End If
                    
                    CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                          ConnectedObjColl.Item(i), ReturnColl, _
                                          RemarkingLineName, STRMFG_PinJig_Remarking_Profile
    
                    Set oResultCS = Nothing
                    Set oProfileVVec = Nothing
                   
                Next

                Set oNI = Nothing
                Set oWireBody = Nothing
                Set oCS = Nothing
                Set oCScoll = Nothing
            End If ' end check for valid geometry
        Next ' end looping around connected objects
    Next ' end looping around supported plates
    
    Dim ProfileLocationMarks As IJElements
    Set ProfileLocationMarks = New JObjectCollection
     
    GetAPSMarkingLines oPinJig, STRMFG_PROFILELOCATION_MARK, ProfileLocationMarks

    If ProfileLocationMarks.Count > 0 Then
        CreateGeom3dFromAPSMarkingLines oPinJig, ElemsToRemark, ProfileLocationMarks, STRMFG_PinJig_Remarking_Profile, ReturnColl
    End If
    Set ProfileLocationMarks = Nothing
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl

    Set oMfgGeomHelper = Nothing
    Set oPlateColl = Nothing
    Set ReturnColl = Nothing
    Set oPartSupport = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5021, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
