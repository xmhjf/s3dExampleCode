VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfLapConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Profile Lap Connection Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      StructMfg     December1st 2010    created
'
'******************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Const MODULE = "MfgPlateMarking.ProfLapConnection"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "ProfLapConnection: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1065, , "RULES")
    GoTo CleanUp
End Function
Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ProfLapConnection: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_LAP_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure you can leave the marking rule
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    
    
    Dim bContourLap As Boolean
    Dim oWBColl As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D

    Dim oMfgGeomHelper As New MfgGeomHelper
            
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        Dim oSDConWrapper As Object
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDConWrapper = New StructDetailObjects.ProfilePart
        ElseIf TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
             Set oSDConWrapper = New StructDetailObjects.MemberPart
        Else
            GoTo NextItem
        End If
        
        bContourLap = oSDPlateWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBColl)
        
        If ((bContourLap = True) And Not (oWBColl Is Nothing)) Then
            If oWBColl.Count = 0 Then
                GoTo NextItem
            End If
            
            Set oSDConWrapper.object = oConnectionData.ToConnectable
            
            Dim oWireBodyColl As IJElements
            Set oWireBodyColl = New JObjectCollection
            
            For nWBIndex = 1 To oWBColl.Count
                Set oWB = oWBColl.Item(nWBIndex)
                oWireBodyColl.Add oWB
            Next nWBIndex
            
            Dim oCSColl As IJElements
            Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBodyColl)
            
            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
            
            If oCSColl Is Nothing Then
                GoTo NextItem
            End If
            
            If oCSColl.Count = 1 Then
                For Each oCS In oCSColl
                    'Create a SystemMark object to store additional information
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    
                    'Set the marking side
                    oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
                    
                    'QI for the MarkingInfo object on the SystemMark
                    Set oMarkingInfo = oSystemMark
                    
                    oMarkingInfo.Thickness = oSDConWrapper.WebThickness
                    oMarkingInfo.Name = oSDConWrapper.Name
                    
                    ' Set the thickness direction vector
                    Dim oWebDirVec As IJDVector
                    Set oWebDirVec = GetWebDirectionVector(oConnectionData.ToConnectable)
                    
                    ' Reverse the vector as we want to lap direction from the lap mark
                    oWebDirVec.Length = -1
                    oMarkingInfo.ThicknessDirection = oWebDirVec
                    
                    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    oGeom3d.PutGeometry oCS
                    oGeom3d.PutGeometrytype STRMFG_LAP_MARK
                    'oGeom3d.PutSubGeometryType STRMFG_LAP_MARK
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Lap"
            
                    oGeomCol3d.AddGeometry 1, oGeom3d
                    oSystemMark.Set3dGeometry oGeom3d
                    
                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker
                    
                    Set oSystemMark = Nothing
                    Set oMarkingInfo = Nothing
                    Set oGeom3d = Nothing
                    Set oWebDirVec = Nothing
                Next
            End If
        End If
              
NextItem:
        Set oWB = Nothing
        Set oWireBodyColl = Nothing
        Set oCSColl = Nothing
        Set oCS = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oCS = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1065, , "RULES")
    GoTo CleanUp
End Function
