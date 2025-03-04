VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileLocationMark"
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
'      TBH        feb. 6. 2002    created
'       KONI        june 22 2002    Tested OK
'      Prashant Gupta    December 05, 2003    Fix for TR 50701
'      Prashant Gupta    January 09, 2004     Fix for TR 53788
'      Prashant Gupta    January 12, 2004     Fix for TR 53039
'      Sundar Kukkala    January 14, 2004     Fix for TR TR-CP·51923  PLFittingPL2PrPro marking on profile end connection
'      Prashant Gupta    February 19, 2004    Fix for TR 54174 (regression)
'      Sundar Kukkal     April 20, 2004       Fix for TR 56587 Made modifications in the CreateProfileTopLocationMarkingLinemethod to know the bounded ness of the profile.
'                                             Changed the approach in calculating direction vector for profile location marking line in the method CreateComplexStringAtEndLocation
'      MJV        2004.04.23      Included correct error handling
'       Koushik   2008.09.30      TR-CP-148665  ErrorLog message should not have been localized
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
   Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreateProfileLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, False, False, False, , 0.01)
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES")
End Function




