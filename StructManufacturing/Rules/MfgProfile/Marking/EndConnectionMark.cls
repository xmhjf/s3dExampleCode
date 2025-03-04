VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndConnection"
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
' 2004.04.22    MJV                         Included correct error handling
' 2004.06.09    Prashant Gupta              Fix for TR 54992
' 2005.01.2005  Sreedhara Venkata Suresh    Fix for TR-CP·63496  In part monitor, each ref_lines' line segement is named and displayed in view
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.EndConnection"

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
    Const METHOD = "EndConnection: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    'This line is added because the method "Connection_ContourProfileEnd" is not implemented for beam part.
    'It can be removed, once it is implemented
    If TypeOf Part Is IJBeamPart Then Exit Function
    
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
     
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_END_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    Dim oSDConProfileWrapper As Object
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim bContourEnd As Boolean
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
    
    If Not (oConObjsCol Is Nothing) Then
        For nIndex = 1 To oConObjsCol.Count
            oConnectionData = oConObjsCol.Item(nIndex)
            
            Dim oStructPort As IJStructPort
            Dim res As Integer
            Set oStructPort = oConnectionData.ConnectingPort
            
            res = oStructPort.ContextID And CTX_LATERAL
                
            If res <= 0 Then GoTo NextItem
            
            'Check if the connected object is a profile
            If Not TypeOf oConnectionData.ToConnectable Is IJStructProfilePart Then
                 GoTo NextItem
            End If
            
            'Initialize the profile wrapper and the Physical Connection wrapper
            'Set oSDConProfileWrapper = New StructDetailObjects.ProfilePart
            If TypeOf oConnectionData.ToConnectable Is IJStiffenerPart Then
                Set oSDConProfileWrapper = New StructDetailObjects.ProfilePart
            ElseIf TypeOf oConnectionData.ToConnectable Is IJBeamPart Then
                Set oSDConProfileWrapper = New StructDetailObjects.BeamPart
            ElseIf TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
                Set oSDConProfileWrapper = New StructDetailObjects.MemberPart
            End If
            Set oSDConProfileWrapper.object = oConnectionData.ToConnectable
                
            Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
            Set oSDPhysicalConn.object = oConnectionData.AppConnection
                
            bContourEnd = oSDProfileWrapper.Connection_ContourProfileEnd(oConnectionData.AppConnection, oWBCollection)
                                                       
            If ((bContourEnd = True) And Not (oWBCollection Is Nothing)) Then
                'Convert the IJWireBody to a IJComplexString
                'Fix for TR 54992
                'iterate through all the elements in the collection
'                For nWBIndex = 1 To oWBCollection.Count
'                    Set oWB = oWBCollection.Item(nWBIndex)

                If oWBCollection.Count = 0 Then
                    GoTo NextItem
                End If

                'If we iterate through each wire body and convert it into complex string then
                'it will result into multiple Marking lines representing same connection which
                'is not wrong but the part monitor will have so many lables around the connection mark
                'showing the Connected part name.
                'Work around is to converting the wirebodies collection into one single complex string
                'and create one Geom3d marking line for the one End connection. This will result into
                'one Marking line and only one lable showing the Part name in Part monitor.
                'Fix for TR 63496.
                    
                    'Merging collection of wirebodies into single wirebody
                    Set oWB = m_oMfgRuleHelper.MergeWireBodyCollection(oWBCollection)
                    'Converting wire body into complex string
                    Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWB)
        
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
        
                    oMarkingInfo.name = oSDConProfileWrapper.name
                    oMarkingInfo.thickness = oSDConProfileWrapper.WebThickness
        
                    oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle
        
                    Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                    oGeom3d.PutGeometry oCS
                    oGeom3d.FaceId = MarkingFace
                    oGeom3d.PutGeometrytype STRMFG_END_MARK
                    Set oObjSystemMark = oSystemMark
                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker
            
                    oSystemMark.Set3dGeometry oGeom3d
                        
                    oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                    lGeomCount = lGeomCount + 1
             '   Next nWBIndex
            End If
NextItem:
                
            Set oWBCollection = Nothing
            Set oWB = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oGeom3d = Nothing
            Set oSDPhysicalConn = Nothing
        Next nIndex
    End If
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oConObjsCol = Nothing
    Set Item = Nothing
    Set oCS = Nothing
    Set oMoniker = Nothing
    Set oGeomCol3d = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2003, , "RULES")
    GoTo CleanUp
End Function


 
