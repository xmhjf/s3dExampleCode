VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fixed"
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

Private Const MODULE = "MfgPlateProcess.Fixed"

Implements IJDMfgPlateNeutralAxisRule

Private Function IJDMfgPlateNeutralAxisRule_GetNeutralAxis(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "NeutralAxisFixed: IJDMfgPlateNeutralAxisRule_GetNeutralAxis"
    On Error GoTo ErrorHandler
    
    'Set the value to the middle of the plate thickness
    IJDMfgPlateNeutralAxisRule_GetNeutralAxis = 0.5
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1049, , "RULES")
End Function
 
