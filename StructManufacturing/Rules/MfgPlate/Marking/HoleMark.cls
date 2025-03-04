VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HoleMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Opening Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      StructMfg     December1st 2010    created
'      karthikeyan     Dec 20 2010    Initial implementation
'
'******************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Const MODULE = "MfgPlateMarking.HoleMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "HoleMark: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1067, , "RULES")
    GoTo CleanUp
End Function
Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "HoleMark: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
        
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
 
    CreateAPSMarkings STRMFG_HOLE_TRACE_MARK, ReferenceObjColl, oGeomCol3d
    CreateAPSMarkings STRMFG_HOLE_REF_MARK, ReferenceObjColl, oGeomCol3d
    
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    
    If bSelectiveRecompute Then
        Exit Function
    End If
 
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = Part
        
    Dim oFeatureColl As Collection
    
    oSDPartSupport.GetFeatures oFeatureColl
    
    Dim icount As Integer
    Dim eFeatureType As StructFeatureTypes
       
    For icount = 1 To oFeatureColl.Count
    
        oSDPartSupport.GetFeatureType oFeatureColl.Item(icount), eFeatureType
        
        Dim oContourColl As Collection
        Dim oPortMonikerColl As Collection
        
        oSDPartSupport.GetFeatureInfo oFeatureColl.Item(icount), SideA, oContourColl, oPortMonikerColl
        
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
        Dim oGeom3dCustom As IJMfgGeom3D ' Initialized in the case of APS Marks
        CreateHoleCrossMarks Part, oFeatureWireBody, UpSide, oFeatureColl.Item(icount), oGeom3dCustom, oGeomCol3d

    Next icount

CleanUp:

    Set oGeomCol3d = Nothing
    Set oSDPartSupport = Nothing
    Set oFeatureWireBody = Nothing
    Set oMergedWireBody = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1067, , "RULES")
    GoTo CleanUp
End Function


