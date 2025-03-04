VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpsideWebRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module:UpsideWebRight
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
    Const METHOD = "UpsideWebRight: IJDMfgProfileUpsideRule_GetUpside"
    On Error GoTo ErrorHandler

    IJDMfgProfileUpsideRule_GetUpside = JXSEC_WEB_RIGHT
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2033, , "RULES")
End Function
 
