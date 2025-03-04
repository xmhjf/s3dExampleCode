VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarPlateLocMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Collar Plate Location Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'       TBH     feb. 6. 2002    created
'       KONI    may. 17 2002    implementation
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.CollarPlateLocMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "CollarPlateLocMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_COLLARPLATELOCATION_MARK, ReferenceObjColl, oGeomCol3d

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
        'Since there is no connecting structure we can leave the marking rule
        GoTo CleanUp
    End If

    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long
    Dim oSDCollarWrapper As New StructDetailObjects.Collar
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    
    ' Loop thru each Physical Connections
    Dim bContourLap As Boolean
    Dim oWBColl As Collection
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    lGeomCount = 1
                
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        'Check if the connected object is a Bracket
        If Not TypeOf oConnectionData.ToConnectable Is IJCollarPart Then
             GoTo NextItem
        End If
        
        'Initialize the profile wrapper and the Physical Connection wrapper
        Set oSDCollarWrapper = New StructDetailObjects.Collar
        Set oSDCollarWrapper.object = oConnectionData.ToConnectable
            
        bContourLap = oSDPlateWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBColl)
        
        If bContourLap = True Then
            Dim lWBIndex As Long
            Dim oWB As IJWireBody
            
            Dim oWireBodyColl As IJElements
            Set oWireBodyColl = New JObjectCollection
            
            For lWBIndex = 1 To oWBColl.Count
                Set oWB = oWBColl.Item(lWBIndex)
                oWireBodyColl.Add oWB
            Next lWBIndex
            
            'Merge the Wire Body collection into one complex strings
            Dim oMfgGeomHelper As New MfgGeomHelper
            Dim oCSColl As IJElements
            
            Set oCSColl = oMfgGeomHelper.MergeCollectionToComplexStrings(oWireBodyColl)
            
            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
            
            If oCSColl Is Nothing Then
                GoTo NextItem
            End If
            
            For Each oCS In oCSColl
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
                'Set the marking side
                oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
    
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
                oMarkingInfo.Name = oSDCollarWrapper.Name

                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_COLLARPLATELOCATION_MARK
                
                oSystemMark.Set3dGeometry oGeom3d
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
                
                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oGeom3d = Nothing
            Next
        End If
        
NextItem:
        Set oWB = Nothing
        Set oWBColl = Nothing
        Set oWireBodyColl = Nothing
        Set oSDCollarWrapper = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set oCS = Nothing
    
CleanUp:

    Set oSDPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oPlateWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1007, , "RULES")
    GoTo CleanUp
End Function


 
