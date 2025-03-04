VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileHole"
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
'To be desired after cycle 3 fixes

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.ProfileHole"

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

    Const METHOD = "ProfileHole: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    'Prepare the output collection of marking line's geometries
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    'Create the 3d collection
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_HOLE_TRACE_MARK, ReferenceObjColl, oGeomCol3d
    CreateAPSProfileMarkings STRMFG_HOLE_REF_MARK, ReferenceObjColl, oGeomCol3d
     
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    If bSelectiveRecompute Then
        Exit Function
    End If
       
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = Part
        
    Dim oFeatureColl As Collection
    
    oSDPartSupport.GetFeatures oFeatureColl
    
    Dim iCount As Integer
    Dim eFeatureType As StructFeatureTypes
       
    For iCount = 1 To oFeatureColl.Count
        
        If Not TypeOf oFeatureColl.Item(iCount) Is IJSketchFeature Then
            GoTo NextFeature
        End If
        
        Dim oContourColl As Collection
        Dim oPortMonikerColl As Collection
        
        oSDPartSupport.GetFeatureInfo oFeatureColl.Item(iCount), SideA, oContourColl, oPortMonikerColl
                        
        Dim lFeatureCurves As Long
        
        If Not oContourColl Is Nothing Then
           lFeatureCurves = oContourColl.Count
        End If
    
        Dim oFeatureWireBody As IJWireBody
        
        If lFeatureCurves > 1 Then
            Dim oMfgMGHelper As IJMfgMGHelper
            Set oMfgMGHelper = New MfgMGHelper
            
            Dim oMergedWireBody As IJWireBody
            
            oMfgMGHelper.MergeWireBodyCollection oContourColl, oMergedWireBody
            
            Set oFeatureWireBody = oMergedWireBody
        Else
            Set oFeatureWireBody = oContourColl.Item(1)
        End If
        
        'Create Cross Marks and add to oGeomCol3d
        Dim oGeom3dCustom As IJMfgGeom3d ' Initialized in the case of APS Marks
        CreateProfileHoleMarks Part, oFeatureWireBody, UpSide, oFeatureColl.Item(iCount), oGeom3dCustom, oGeomCol3d
       
NextFeature:
    Next iCount

    'Return collection of ML's geometry
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:

    Set oSDPartSupport = Nothing

Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES")
End Function
