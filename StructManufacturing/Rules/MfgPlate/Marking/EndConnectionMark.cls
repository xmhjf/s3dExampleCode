VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndConnectionMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the (Profile) End Connection Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'       TBH         feb. 6. 2002    created
'       KONI        june 22 2002    Tested OK
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.EndConnectionMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "EndConnectionMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_END_MARK, ReferenceObjColl, oGeomCol3d

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
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart
    Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim bContourEnd As Boolean
    Dim oWBColl As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    Dim oMfgMGHelper As IJMfgMGHelper
       
    'Create an instance of the StrMfg math geom helper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
            
    lGeomCount = 1
            
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        'Check if the connected object is a profile
        If Not TypeOf oConnectionData.ToConnectable Is IJStructProfilePart Then
             GoTo NextItem
        End If
            
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection
            
        bContourEnd = oSDPlateWrapper.Connection_ContourProfileEnd(oConnectionData.AppConnection, oWBColl)
                                                           
        'Initialize the profile wrapper and the Physical Connection wrapper
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
            Set oSDProfileWrapper.object = oConnectionData.ToConnectable
        Else
            Set oSDMemberWrapper = New StructDetailObjects.MemberPart
            Set oSDMemberWrapper.object = oConnectionData.ToConnectable
        End If
        
        If ((bContourEnd = True) And Not (oWBColl Is Nothing)) Then
            If oWBColl.Count = 0 Then
                GoTo NextItem
            End If
            'Merge the outer contour wirebody collection into one wirebody
            Dim oProfileEndContour As IJWireBody
            oMfgMGHelper.MergeWireBodyCollection oWBColl, oProfileEndContour
   
            Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oProfileEndContour)
    
            'Create a SystemMark object to store additional information
             Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
            'Set the marking side
            oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
    
            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark
    
            If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
                oMarkingInfo.Name = oSDProfileWrapper.Name
                oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness
            Else
                oMarkingInfo.Name = oSDMemberWrapper.Name
                oMarkingInfo.Thickness = oSDMemberWrapper.WebThickness
            End If
    
            oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle
    
            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_END_MARK
            Set oObjSystemMark = oSystemMark
            
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker
            
            oSystemMark.Set3dGeometry oGeom3d
            oGeomCol3d.AddGeometry lGeomCount, oGeom3d
            lGeomCount = lGeomCount + 1
            Set oProfileEndContour = Nothing
            Set oCS = Nothing
            Set oGeom3d = Nothing
            Set oSystemMark = Nothing
            Set oObjSystemMark = Nothing
            Set oMarkingInfo = Nothing

        End If
NextItem:
            
        Set oWB = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDProfileWrapper = Nothing
        Set oSDMemberWrapper = Nothing
        Set oSDPhysicalConn = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set oCS = Nothing

CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oMfgMGHelper = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1009, , "RULES")
    GoTo CleanUp
End Function

 
