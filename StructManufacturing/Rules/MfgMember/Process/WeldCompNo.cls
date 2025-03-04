VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompNo
'
' Description:  Remove Weld Compensation from Profile
'
'  Comments:
'
' 8 Feb, 2008    Creation
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkRule

Const MODULE = "MfgMemberProcess.WeldCompNo"


Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    

    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2036, , "RULES")
End Function

 
