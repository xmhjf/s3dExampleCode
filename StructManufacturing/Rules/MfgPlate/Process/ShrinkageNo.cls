VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShrinkageNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module:
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2010.07.14    StructMfg    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule2
Const Module = "MfgPlateProcess.ShrinkageNo"
Private Sub IJDMfgProcessRule2_ExecuteBeforeUnfoldRule(ByVal pMfgPart As Object)
    Const METHOD = "ShrinkageNo: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1080, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJDMfgProcessRule2_ExecuteAfterUnfoldRule(ByVal oMfgPart As Object)
    Const METHOD = "ShrinkageNo: IJDMfgProcessRule2_ExecuteAfterUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1080, , "RULES")
    GoTo CleanUp
End Sub

