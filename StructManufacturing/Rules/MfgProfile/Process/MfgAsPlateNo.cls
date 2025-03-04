VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgAsPlateNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module:MfgAsProfileNo
'
' Description:  Provides the MfgAsPlate value for the mfg profile
'
' Author:
'
' Comments:
' 28 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileProcessMfgAsPlateRule
Const MODULE = "MfgProfileProcess"

Private Function IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate(ByVal ProfilePart As Object) As Long
    Const METHOD = "MfgAsPlate: IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate"
    On Error GoTo ErrorHandler
        IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate = 0
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2045, , "RULES")
End Function
 



