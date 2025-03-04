VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess
' Module: BevelFixed
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Private Const Module = "MfgPlateProcess.BevelFixed"

Implements IJDMfgPlateBevelRule

Private Function IJDMfgPlateBevelRule_GetBevelDeviationAngle(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "BevelFixed: IJDMfgPlateBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    ' Set the deviation angle when the physical connection should be splitted

    ' 3 degrees: GNEST hard-coded standard max. deviation-angle
    ' (angles are measured in radians in GSCAD)
    IJDMfgPlateBevelRule_GetBevelDeviationAngle = (3# * 3.141529) / 180#

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Function
