VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Knuckle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.Knuckle"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "Knuckle: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
        
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2006, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

 
