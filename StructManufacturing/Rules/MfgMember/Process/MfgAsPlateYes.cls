VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgAsPlateYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*******************************************************************************
' Copyright (C) 2008, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberProcess
' Module: MfgAsProfileYes
'
' Description:  Provides the MfgAsPlate value for the mfg member
'
' Author:
'
' Comments:
' 28 Feb, 2008    Creation
'*******************************************************************************
Option Explicit

Implements IJDMfgProfileProcessMfgAsPlateRule
Const MODULE = "MfgMemberProcess"

Private Function IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate(ByVal ProfilePart As Object) As Long

    Const METHOD = "MfgAsPlate: IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate"
    On Error GoTo ErrorHandler
        SetBuiltUpNames ProfilePart, True
    
    IJDMfgProfileProcessMfgAsPlateRule_GetMfgAsPlate = 0        ' 1
    'Note: This will not be supported for initial version. Best solution is to have the rule return '0'.
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2046, , "RULES")
End Function
 

