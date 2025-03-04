VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileEndFitName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Create Profile Location Marks with End Fitting Mark & Assembly Name
'
'  History:
'      Ninad        feb. 9. 2011    created
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.ProfileLocationMark"
Private Const PI As Double = 3.14159265358979 'required for TR 53039

Private Const IID_IJMfgSystemMarkInfo As Variant = "{CE13B4E9-C6B0-45D6-99F4-0981D870274E}"
    
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

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "ProfileLocationMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
        
    '*** CreateProfileLocationMark without End Fitting Marks and REND Fitting Marks ***'
   Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreateProfileLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, True, True, True, False, 1, , True)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
End Function





