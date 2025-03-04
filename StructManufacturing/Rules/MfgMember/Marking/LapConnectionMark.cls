VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LapConnection"
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

Private Const MODULE = "MfgMemberMarking.LapConnection"

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
     Const METHOD = "LapConnection: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler

    'Create the SD profile Wrapper and initialize it
    Dim oSDMemberWrapper As Object
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    'Get the Profile Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConObjsCol Is Nothing Then
        'Since there is no connectingstructure you can leave the function
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim bContourLap As Boolean
    Dim oWBCollection As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
    Dim lGeomCount As Long
    lGeomCount = 1
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        ' Assuming we only have lapped plateconnections
        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then GoTo NextItem
        
        'Initialize the profile wrapper and the Physical Connection wrapper
        Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDConPlateWrapper.object = oConnectionData.ToConnectable
            
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection
            
        bContourLap = oSDMemberWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBCollection)
                                                   
        If ((bContourLap = True) And Not (oWBCollection Is Nothing)) Then
            If oWBCollection.Count = 0 Then
                'Since elements count is zero
                GoTo NextItem
            End If
            'Convert the IJWireBody to a IJComplexString
'            For nWBIndex = 1 To oWBCollection.Count
            For nWBIndex = 1 To oWBCollection.Count
                Set oWB = oWBCollection.Item(nWBIndex)
                'Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWB)
    
                Dim oCSColl As IJElements
                Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWB)
                
                For Each oCS In oCSColl
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                
                Dim oSubPort As IJStructPort
                Set oSubPort = oProfileWrapper.GetProfileSubPort(oConnectionData.ConnectingPort, oConnectionData.AppConnection)
                
                'Set the marking side
                Dim MarkingFace As Long
                MarkingFace = oProfileWrapper.GetSide(oSubPort)
                oSystemMark.SetMarkingSide MarkingFace
    
                Set oSubPort = Nothing
                
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
                oMarkingInfo.name = oSDConPlateWrapper.name
    
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_LAP_MARK
                oGeom3d.FaceId = MarkingFace
                Set oObjSystemMark = oSystemMark

                oSystemMark.Set3dGeometry oGeom3d
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
            
                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1
                Next oCS
    
                
            Next nWBIndex
        End If
NextItem:
            
        Set oWBCollection = Nothing
        Set oWB = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConPlateWrapper = Nothing
        Set oSDPhysicalConn = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing

CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2007, , "RULES")
    GoTo CleanUp
End Function


 
