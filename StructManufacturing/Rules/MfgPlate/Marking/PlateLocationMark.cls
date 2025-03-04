VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateLocationMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Plate Edge check Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'       KONI        june 22 2002    Tested OK
'       sgorthi     May 27 2003 Fix for TR#39584
'      MJV        2004.04.23      Included correct error handling
'       Koushik   2008.09.30      TR-CP-148665  ErrorLog message should not have been localized
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const PI As Double = 3.14159265358979
Private Const MODULE = "MfgPlateMarking.PlateLocationMark"

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

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = CreatePlateLocationMark(Part, UpSide, bSelectiveRecompute, ReferenceObjColl, False, False, False, False)
End Function
