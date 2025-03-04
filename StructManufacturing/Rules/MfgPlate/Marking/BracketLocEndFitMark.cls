VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketLocEndFitMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Bracket Location Markings with Fitting Marks
'
'  History:
'       09/17/2010      Ninad       CR-CP-184245  Ref: Asia: Mfg Ref Data + 9.11 Bracket End Fitting Mark
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.BracketLocEndFitMark"

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    PlMrkHelpers.Initialize
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1004, , "RULES")
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

    '*** Call the function to create the Thickness Direction "Bubble" Mark on Profile Location Marking Line ***'
    'Set IJDMfgSystemMarkingRule_CreateAfterUnfold = CreateThicknessDirectionMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, STRMFG_BRACKETLOCATION_MARK, "Circle")

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "BracketLocationMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    '*** CreateProfileLocationMark with End Fitting Marks and REND Fitting Marks ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreateBracketLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, True, False, False, , False)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1004, , "RULES")
End Function


