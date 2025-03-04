VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateLocEndFitMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Plate Location Mark with Fitting Marks
'
'  History:
'           09/17/2010   Ninad      Created
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const PI As Double = 3.14159265358979
Private Const MODULE = "MfgPlateMarking.PlateLocEndFitMark"

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Err GoTo ErrorHandler
    
    PlMrkHelpers.Initialize
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES")
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    '*** Call the function to create the Thickness Direction "Bubble" Mark on Profile Location Marking Line ***'
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = CreateThicknessDirectionMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, STRMFG_PLATELOCATION_MARK, "Circle")
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "PlateLocEndFitMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Err GoTo ErrorHandler
    
    '*** PlateLocEndFitMark with End Fitting Marks and REND Fitting Marks ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreatePlateLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, True, True, False)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES")
End Function

