VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RollLinesFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: AntiMoldedSide
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgPlateProcess.RollLinesFixed"

Implements IJDMfgProcessRollLines

Private Function IJDMfgProcessRollLines_ExecuteRule(ByVal pPart As Object) As Long
    Const METHOD = "RollLinesFixed: IJDMfgProcessRollLines_ExecuteRule"
    On Error GoTo ErrorHandler

    IJDMfgProcessRollLines_ExecuteRule = 3
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1054, , "RULES")
End Function
 
