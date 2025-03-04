VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NAOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: NAOne
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
'
'*******************************************************************************
Option Explicit

Private Const Module = "MfgPlateProcess.NAOne"

Implements IJDMfgPlateNeutralAxisRule

Private Function IJDMfgPlateNeutralAxisRule_GetNeutralAxis(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "NeutralAxisFixed: IJDMfgPlateNeutralAxisRule_GetNeutralAxis"
    On Error GoTo ErrorHandler
    
    'Set the value to 1.0
    IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 1#
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1049, , "RULES")
End Function
 


