VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnfoldYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: UnFoldYes
'
' Description:  Provides UnFold value for the mfg member
'
' Author:
'
' Comments:
' 02/26/2008    Teja    Created
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileProcessUnfoldRule
Const MODULE = "MfgMemberProcess"

Private Function IJDMfgProfileProcessUnfoldRule_GetUnfold(ByVal ProfilePart As Object) As Long
    Const METHOD = "UnfoldNo:IJDMfgProfileProcessUnfoldRule_GetUnfold"
    On Error GoTo ErrorHandler
    'Set the value for unfold
    IJDMfgProfileProcessUnfoldRule_GetUnfold = 1
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2044, , "RULES")
End Function


