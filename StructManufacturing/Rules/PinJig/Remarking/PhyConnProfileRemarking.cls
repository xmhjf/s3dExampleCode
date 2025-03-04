VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhyConnProfileRemarking"
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
Private Const MODULE As String = "MfgPinJigRule.PhyConnProfileRemarking"

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
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionPhysical, oPlateColl)

    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PROFILE_TYPE, _
                            ConnectionPhysical, _
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
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5020, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsConnectedProfilePart]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject

    Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
    Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
    
    Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.PartSupport

    Dim oMfgGeomHelper As New MfgGeomHelper

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionPhysical, oPlateColl)

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
                            PROFILE_TYPE, _
                            ConnectionPhysical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl

        Dim i As Long
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            If Not ElemsToRemark.Contains(ConnectedObjColl.Item(i)) Then GoTo NextConnObj

        ' ... get the desired portion of the geometry (SideA/SideB for physical)
            Dim oLocalPartSupport As GSCADSDPartSupport.IJPartSupport
            Set oLocalPartSupport = oProfilePartSupport
            Set oLocalPartSupport.Part = ConnectedObjColl.Item(i)

            Dim eTSide As GSCADSDPartSupport.ThicknessSide
            eTSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
            Set oLocalPartSupport = Nothing

        ' ... get the connection contour for the Tee type connection
            Set oPartSupport.Part = oPlateColl.Item(Iter)
            
            Dim oWireBody As IJWireBody
            Dim oThickDir As IJDVector
            oPartSupport.GetConnectionContour_Tee ConnectionsColl.Item(i), oWireBody, oThickDir, eTSide
            Set oThickDir = Nothing

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
                    CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                          ConnectedObjColl.Item(i), ReturnColl, _
                                          RemarkingLineName, STRMFG_PinJig_Remarking_Profile
                Next

                Set oNI = Nothing
                Set oWireBody = Nothing
                Set oCS = Nothing
                Set oCScoll = Nothing
            End If ' end check for valid geometry
NextConnObj:
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
    Set oProfilePartSupport = Nothing
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
