VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocationWithFitMark"
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
'           09/27/2010      Ninad       CR-CP-184322 Ref: Asia: Mfg Ref Data + 13.9 Stiffener End Fitting Mark
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgProfileMarking.Location"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    '*** Call the function to create the Thickness Direction "Bubble" Mark on Profile Location Marking Line ***'
    'Set IJDMfgSystemMarkingRule_CreateAfterUnfold = CreateThicknessDirectionMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, STRMFG_PLATELOCATION_MARK, "Circle")

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "LocationWithFitMark"
    On Error GoTo ErrorHandler
    
    '*** CreateLocationMark with End Fitting Marks ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreateLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, True)
    'Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreatePenetrationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, False)
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2008, , "RULES")

End Function



 


