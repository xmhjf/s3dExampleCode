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

Private Const MODULE = "MfgProfileMarking.LapConnection"
Private Const NEEDLABEL As Boolean = True
Private Const CHECKFORSINGLEWB As Boolean = False


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
     Const METHOD = "LapConnection: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    'This line is added because the method "Connection_ContourLap" is not implemented for beam part.
    'It can be removed, once it is implemented
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_LAP_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

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
    Dim oSDConProfileWrapper As New StructDetailObjects.ProfilePart
    
    ' Loop thru each Physical Connections
    Dim bContourLap As Boolean
    Dim oWBCollection As Collection
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3d
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
        
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
            Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
            Set oSDConPlateWrapper.object = oConnectionData.ToConnectable
        ElseIf TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDConProfileWrapper = New StructDetailObjects.ProfilePart
            Set oSDConProfileWrapper.object = oConnectionData.ToConnectable
        Else
            GoTo NextItem
        End If
        
        bContourLap = oSDProfileWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBCollection)
        
        If ((bContourLap = True) And Not (oWBCollection Is Nothing)) Then
            'Convert the IJWireBody to a IJComplexString

            If oWBCollection.Count = 0 Then
                GoTo NextItem
            End If
            
            If CHECKFORSINGLEWB = True Then
                If Not oWBCollection.Count = 1 Then
                    GoTo NextItem
                End If
            End If
            
            ' Convert VB collection of Wirebodies to IJElement collection of Wirebodies
            Dim oWBelements As IJElements
            Set oWBelements = New JObjectCollection

            For nWBIndex = 1 To oWBCollection.Count
                oWBelements.Add oWBCollection.Item(nWBIndex)
            Next nWBIndex
            
            ' From IJElement collection of Wirebodies, get minimal set of complex strings
            Dim oCScoll As IJElements
            Set oCScoll = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWBelements)
            
            Set oWBelements = Nothing
           
            For Each oCS In oCScoll
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                
                'Get the sub port having section information
                Dim oSubPort As IJStructPort
                Set oSubPort = oProfileWrapper.GetProfileSubPort(oConnectionData.ConnectingPort, oConnectionData.AppConnection)
                
                'Set the marking side
                Dim MarkingFace As Long
                MarkingFace = oProfileWrapper.GetSide(oSubPort)
                oSystemMark.SetMarkingSide MarkingFace
                
                Set oSubPort = Nothing
                
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
                
                If NEEDLABEL = True Then
                    If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
                        oMarkingInfo.name = oSDConPlateWrapper.name
                    Else
                        oMarkingInfo.name = oSDConProfileWrapper.name
                    End If
                Else
                    oMarkingInfo.name = ""  ' Give null string as label when not required
                End If
                
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_LAP_MARK
                oGeom3d.FaceId = MarkingFace
                Set oObjSystemMark = oSystemMark
                
                oSystemMark.Set3dGeometry oGeom3d
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
            
                oGeomCol3d.AddGeometry 1, oGeom3d
            Next
            
            Set oCS = Nothing
            Set oCScoll = Nothing
        End If
NextItem:
            
        Set oWBCollection = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDConPlateWrapper = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing

CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2007, , "RULES")
    GoTo CleanUp
End Function


 
