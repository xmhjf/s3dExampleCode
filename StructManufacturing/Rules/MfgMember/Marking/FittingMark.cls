VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittingProfile"
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

Private Const MODULE = "MfgMemberMarking.FittingProfile"

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
    Const METHOD = "FittingProfile: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    'Prepare collections
    Dim oSDMemberWrapper As Object
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part

    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgPart As IJMfgProfilePart
    Dim oMfgProfileWrapper As New MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    'Retrieve collection of ConnectionData objects
    Dim oConnectedObjectsColl As Collection
    'Set oConnectedObjectsColl = oSDMemberWrapper.ConnectedObjects
    Set oConnectedObjectsColl = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    If oConnectedObjectsColl Is Nothing Then
        'Since there is no connecting structure there will be no fitting marks required, therefore we can leave the function
        GoTo CleanUp
    End If
    
    If oConnectedObjectsColl.Count = 0 Then
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData
    Dim oConnectionDataTmp As ConnectionData
    Dim i As Integer
    
    'Find mounting plate ConnectionData object
    For i = 1 To oConnectedObjectsColl.Count
        oConnectionDataTmp = oConnectedObjectsColl.Item(i)
        
        If TypeOf oConnectionDataTmp.ToConnectable Is IJCollarPart Then GoTo NextItem
        If TypeOf oConnectionDataTmp.ToConnectable Is IJSmartPlate Then GoTo NextItem
        If TypeOf oConnectionDataTmp.ToConnectable Is IJProfilePart Then GoTo NextItem
        If Not TypeOf oConnectionDataTmp.ToConnectable Is IJPlatePart Then GoTo NextItem
        
        Dim res As Integer
        Dim oStructPort As IJStructPort
        Set oStructPort = oConnectionDataTmp.ConnectingPort
        res = oStructPort.ContextID And IMSStructConnection.CTX_LATERAL
        If res <= 0 Then GoTo NextItem
        
        oConnectionData = oConnectionDataTmp
NextItem:
    Next i
    
    If oConnectionData.AppConnection Is Nothing Then GoTo ErrorHandler
        
    'Get collection of mark positions
    Dim oMarkPosColl As Collection
    On Error Resume Next
    Set oMarkPosColl = GetFittingMarkPositions(Part, oConnectionData, UpSide)
    
    If Not (oMarkPosColl Is Nothing) Then
        Dim oWireBody As IJWireBody
        Dim oCS As IJComplexString
        Dim oGeom3d As IJMfgGeom3d
        For i = 1 To oMarkPosColl.Count
            Set oWireBody = CreateFittingMarkLine(Part, oMarkPosColl.Item(i), oConnectionData, UpSide)
           
            If Not (oWireBody Is Nothing) Then
               Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWireBody)
                If Not (oCS Is Nothing) Then
                    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                    oGeom3d.PutGeometry oCS
                    oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
                    oGeom3d.FaceId = UpSide
                
                    Dim oSystemMark As IJMfgSystemMark
                    Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                    oSystemMark.SetMarkingSide UpSide
                    oSystemMark.Set3dGeometry oGeom3d
                
                    'QI for the MarkingInfo object on the SystemMark
                    Dim oMarkingInfo As IJMarkingInfo
                    Set oMarkingInfo = oSystemMark
                
                    'Set marking information, if any
                    oMarkingInfo.name = "FITTING"
                    
                    Dim oObjSystemMark As IUnknown
                    Dim oMoniker As IMoniker
                   
                    Set oObjSystemMark = oSystemMark
                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker
            
                    oGeomCol3d.AddGeometry 1, oGeom3d
                    
                    Set oStructPort = Nothing
                    Set oWireBody = Nothing
                    Set oCS = Nothing
                    Set oGeom3d = Nothing
                    Set oSystemMark = Nothing
                    Set oMarkingInfo = Nothing
                    Set oObjSystemMark = Nothing
                    Set oMoniker = Nothing
                End If
            End If
        Next i
    End If
    
        
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oMfgPart = Nothing
    Set oMfgProfileWrapper = Nothing
    Set oConnectedObjectsColl = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2004, , "RULES")
    GoTo CleanUp
End Function

 
