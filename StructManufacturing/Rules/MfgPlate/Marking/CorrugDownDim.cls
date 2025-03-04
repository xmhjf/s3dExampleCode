VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorrugDownDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Corrugated Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const MODULE = "MfgPlateMarking.CorrugDownDim :: "

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As JCmnShp_CollectionAlias) As IJMfgGeomCol2d
    Const METHOD = "IJDMfgSystemMarkingRule_CreateAfterUnfold"
    On Error GoTo ErrorHandler
        
     'For Creating Both Corrugate Marks and Dimension Marks
     CreateCorrugateMarksAndDimensionMarks Part, UpSide, False, True, True, True, True
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1027, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As JCmnShp_CollectionAlias) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1027, , "RULES")
    GoTo CleanUp
End Function



