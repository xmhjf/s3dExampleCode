VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: BevelNo
'
' Description:  Provides the bevel deviation angle for the mfg profile
'
' Author:
'
' Comments:
' 26/2/2008    Teja     Created
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgProfileProcess.BevelNo"

Implements IJDMfgProfileBevelRule

Private Function IJDMfgProfileBevelRule_GetBevelDeviationAngle(ByVal ProfilePart As Object) As Double
    Const METHOD = "BevelNo: IJDMfgProfileBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    'Set the deviation angle at which the physical connection should be splitted
    
    ' 90 degrees: Effectively disables varying bevel.
    ' (angles are measured in radinas in GSCAD)
    IJDMfgProfileBevelRule_GetBevelDeviationAngle = (90# * 3.141529) / 180#

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2042, , "RULES")
End Function

 
