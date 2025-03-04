VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideWebLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module:UpsideWebLeft
'
' Description:  Determines the process settings for the mfg profile
'
' Author:
'
' Comments:
' 8 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileUpsideRule

Const MODULE = "MfgMemberProcess"

Private Function IJDMfgProfileUpsideRule_GetUpside(ByVal ProfilePart As Object) As Long
    Const METHOD = "UpsideWebLeft: IJDMfgProfileUpsideRule_GetUpside"
    On Error GoTo ErrorHandler
    
    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_LEFT
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2032, , "RULES")
End Function
 
