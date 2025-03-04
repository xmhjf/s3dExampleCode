VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldTabNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: WeldTabNo
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProcessRule
Const MODULE = "MfgMemberProcess"

Private Sub IJDMfgProcessRule_ExecuteRule(ByVal Part As Object, MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "WeldTabNo: IJDMfgProcessRule_ExecuteRule"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2039, , "RULES")
End Sub
 
