VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefPlanesByBlockFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
'  History:
'           09/24/2010   Suma Mallena      Created
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.RefPlanesByBlockFR"

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
    Const METHOD = "RefPlanesByBlockFr: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
   
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_FRAMELINE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Dim oCSColl         As IJElements
    Dim oMfgGeomHelper  As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
   
    Dim oProfileWrapper As StructDetailObjects.ProfilePart
    Set oProfileWrapper = New StructDetailObjects.ProfilePart
    Set oProfileWrapper.object = Part
    
    'Get the web-left Surface
    Dim oWebLeftPort            As IJPort
    Dim oProfileSurfGeom        As Object
    
    Set oWebLeftPort = oProfileWrapper.SubPort(JXSEC_WEB_LEFT)
    Set oProfileSurfGeom = oWebLeftPort.Geometry

    Dim nIndex                  As Long
    Dim oConObjsCol             As Collection
    Dim oRefplanes              As Collection

    ' LEAD_BLOCK_REFERENCE  -- Gets the frames only from the immediate parent block
    m_oMfgRuleHelper.GetBlockFrames Part, oProfileSurfGeom, x, LEAD_BLOCK_REFERENCE, oRefplanes, oConObjsCol
    
    If Not (oConObjsCol Is Nothing) Then
    
        Dim oSPGMiddleHelper    As New SP3dGSMiddleHelper.SPGMiddleHelper
        Dim oFrameSystem        As IJDCoordinateSystem
        Dim oFrameSysMoniker    As IMoniker

        If oConObjsCol.Count > 0 Then
            'Get the Frame System
            oSPGMiddleHelper.GetCoordinateSystemOfPlane oRefplanes(1), oFrameSystem
            Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oFrameSystem)
        End If
        
        On Error Resume Next
        For nIndex = 1 To oConObjsCol.Count
        
            Dim oWB                 As IJWireBody
            Dim oCS                 As IJComplexString
            Dim oSystemMark         As IJMfgSystemMark
            Dim oMarkingInfo        As MarkingInfo
            Dim oGeom3d             As IJMfgGeom3d
            Dim oNamedItem          As IJNamedItem
            
            Set oWB = oConObjsCol.Item(nIndex)
    
        '   Convert the IJWireBody to a IJComplexString
            Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWB)
            
            If Not oCSColl Is Nothing Then
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                End If
            End If
                
            If (oCSColl Is Nothing) Then
                GoTo NextItem
            End If
            
             ' Get type of Frame and set as custom attribute on Marking Info
            Dim oGridData As ISPGGridData
            Set oGridData = oRefplanes.Item(nIndex)
            
            Dim oCodeListMetaData As IJDCodeListMetaData
            Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
            
            Dim strPlaneType As String
            strPlaneType = oCodeListMetaData.LongStringValue("GridPlaneType", oGridData.Type)
            
            For Each oCS In oCSColl
            '   Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
    
            '   Set the marking side
                oSystemMark.SetMarkingSide UpSide
    
            '   QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
    
            '   Get the Frame Name
                Set oNamedItem = oRefplanes(nIndex)
                oMarkingInfo.name = oNamedItem.name
                
                'Set the Frame Type
                If Not strPlaneType = vbNullString Then
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", strPlaneType
                End If
    
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_FRAMELINE_MARK
                oGeom3d.FaceId = UpSide
                
            '   Set Frame System Moniker on the Geom3d object
                oGeom3d.PutMoniker oFrameSysMoniker
                oGeom3d.Put3dModelMoniker m_oMfgRuleHelper.GetMoniker(oRefplanes.Item(nIndex))
              
                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d
                
                Set oSystemMark = Nothing
                Set oGeom3d = Nothing
                Set oMarkingInfo = Nothing
                Set oCS = Nothing
                Set oNamedItem = Nothing
                
            Next oCS
NextItem:
        
            Set oWB = Nothing
            Set oCS = Nothing
            Set oSystemMark = Nothing
            Set oGeom3d = Nothing
            Set oMarkingInfo = Nothing
            Set oGridData = Nothing
            Set oCodeListMetaData = Nothing
        Next nIndex
        
    End If

'    Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:

    Set oConObjsCol = Nothing
    Set oRefplanes = Nothing
    Set oProfileWrapper = Nothing
    Set oWebLeftPort = Nothing
    Set oProfileSurfGeom = Nothing
    Set oFrameSystem = Nothing
    Set oSPGMiddleHelper = Nothing
    Set oFrameSystem = Nothing
    Set oFrameSysMoniker = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2012, , "RULES")
    GoTo CleanUp
End Function

