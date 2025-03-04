VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileLocEndFitMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Profile Location Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'            08/26/2010     Ninad      CR-CP-184271  Ref: Asia: Mfg Ref Data + 9.10 Stiffener End Fitting mark
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.ProfileLocEndFitMark"
Private Const PI As Double = 3.14159265358979 'required for TR 53039
Private Const ER_OFFSET_DIST As Double = 0.1
'Private Const IID_IJMfgSystemMarkInfo As Variant = "{CE13B4E9-C6B0-45D6-99F4-0981D870274E}"
    
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Err GoTo ErrorHandler
    
    PlMrkHelpers.Initialize
    Exit Sub
ErrorHandler:
Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    '*** Call the function to create the Thickness Direction "Bubble" Mark on Profile Location Marking Line ***'
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = CreateThicknessDirectionMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, STRMFG_PROFILELOCATION_MARK, "Circle")
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ProfileLocEndFitMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
        
    '*** CreateProfileLocationMark with End Fitting Marks and REND Fitting Marks ***'
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreateProfileLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, True, True, True, False, , 0.01)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
End Function






