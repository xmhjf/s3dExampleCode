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
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgProfileMarking
'
'  Abstract:    Rule for creating the End Connection Markings in the MfgProfile command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      A.Santosh Ram     June 29th 2011    created
'
'******************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Const MODULE = "MfgProfileMarking.EndConnectionLogical"

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
    Const METHOD = "EndConnectionLogical: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_END_MARK, ReferenceObjColl, oGeomCol3d
    
    Dim oProfilePartSupport As IJPartSupport
    Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oProfilePartSupport.Part = Part
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oProjCS As IJComplexString
    
    Dim oConnobjColl As Collection
    Dim oConnColl As Collection
    Dim oThisPortColl As Collection
    Dim oOtherPortColl As Collection
    Dim iCount As Integer
    
    'Getting the Connection Objects and Physical Connection along with Ports of the Input Profile
    oProfilePartSupport.GetConnectedObjects ConnectionPhysical, oConnobjColl, oConnColl, oThisPortColl, oOtherPortColl
    
    If oConnobjColl Is Nothing Then GoTo CleanUp
    
    Dim oConnPartColl As IJElements
    Set oConnPartColl = New JObjectCollection
    
    For iCount = 1 To oConnobjColl.Count
        If TypeOf oConnobjColl.Item(iCount) Is IJStructProfilePart Then
            Dim oThisPort As IJPort
            Dim oAppConnection As IJAppConnection
            
            Dim oSDProfileWrapper As Object
            If TypeOf oConnobjColl.Item(iCount) Is IJStiffenerPart Then
                Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
            ElseIf TypeOf oConnobjColl.Item(iCount) Is IJBeamPart Then
                Set oSDProfileWrapper = New StructDetailObjects.BeamPart
            ElseIf TypeOf oConnobjColl.Item(iCount) Is ISPSMemberPartPrismatic Then
                Set oSDProfileWrapper = New StructDetailObjects.MemberPart
            End If
            
            Set oSDProfileWrapper.object = oConnobjColl.Item(iCount)
            Set oThisPort = oThisPortColl.Item(iCount)
            Set oAppConnection = oConnColl.Item(iCount)
            
            Dim oThisPortSB As IJSurfaceBody
            Set oThisPortSB = oThisPort.Geometry
            
            Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oAppConnection

            'To Ensure Duplication of marks is Avoided
            If oConnPartColl.Contains(oConnobjColl.Item(iCount)) Then
                GoTo NextItem
            Else
                oConnPartColl.Add oConnobjColl.Item(iCount)
            End If
            
            Dim oCSMarkColl As IJElements
            'Creating the End Connection Marks on Profile
            Set oCSMarkColl = CreateSpecialEndConnMarks(oConnobjColl.Item(iCount), oThisPort)
            
            If oCSMarkColl Is Nothing Then
                GoTo NextItem
            ElseIf oCSMarkColl.Count = 0 Then
                GoTo NextItem
            End If
            
            Dim oMarkCS1 As IJComplexString
            Dim oMarkCS2 As IJComplexString

            Set oMarkCS1 = oCSMarkColl.Item(1)
            Set oMarkCS2 = oCSMarkColl.Item(2)
            
            Dim jCount As Long
            
            For jCount = 1 To 2
            
                Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
                Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
                Set oProfileWrapper.object = Part
                
                Dim oSystemMark As IJMfgSystemMark
                Dim oMoniker As IMoniker
                Dim oMarkingInfo As MarkingInfo
                Dim oGeom3d As IJMfgGeom3d
                
                'Create a SystemMark object to store additional information
                 Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)

                'Set the marking side
                oSystemMark.SetMarkingSide oProfileWrapper.GetSide(oThisPort)

                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

                oMarkingInfo.thickness = oSDProfileWrapper.WebThickness

                oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle

                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                If jCount = 1 Then
                    Set oProjCS = Nothing
                    oMfgMGHelper.ProjectComplexStringToSurface oMarkCS1, oThisPortSB, Nothing, oProjCS
                    If oProjCS Is Nothing Then
                    oGeom3d.PutGeometry oMarkCS1
                Else
                        oGeom3d.PutGeometry oProjCS
                    End If
                Else
                    Set oProjCS = Nothing
                    oMfgMGHelper.ProjectComplexStringToSurface oMarkCS2, oThisPortSB, Nothing, oProjCS
                    If oProjCS Is Nothing Then
                    oGeom3d.PutGeometry oMarkCS2
                    Else
                        oGeom3d.PutGeometry oProjCS
                End If
                End If

                oGeom3d.PutGeometrytype STRMFG_END_MARK
                oGeom3d.FaceId = oProfileWrapper.GetSide(oThisPort)

                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oAppConnection)
                oGeom3d.PutMoniker oMoniker

                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d
                
                Set oProfileWrapper = Nothing
                Set oGeom3d = Nothing
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oMoniker = Nothing
                
            Next jCount
            
            Set oMarkCS1 = Nothing
            Set oMarkCS2 = Nothing
            Set oCSMarkColl = Nothing
            Set oThisPort = Nothing
            Set oAppConnection = Nothing
            Set oSDProfileWrapper = Nothing
        End If
NextItem:
    Next iCount
    
CleanUp:
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    Set oConnobjColl = Nothing
    Set oConnColl = Nothing
    Set oThisPortColl = Nothing
    Set oOtherPortColl = Nothing
    Set oGeomCol3d = Nothing
    Set oResourceManager = Nothing
    Set oProfilePartSupport = Nothing
    Set oConnPartColl = Nothing
    Set oMfgMGHelper = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2003, , "RULES")
    GoTo CleanUp
End Function
