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

Private Const MODULE = "MfgPlateProcess.BevelNo"

Implements IJDMfgPlateBevelRule

Private Function IJDMfgPlateBevelRule_GetBevelDeviationAngle(ByVal PlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As Double
    Const METHOD = "BevelNo: IJDMfgPlateBevelRule_GetBevelDeviationAngle"
    On Error GoTo ErrorHandler

    'Set the deviation angle at which the physical connection should be splitted
    
    ' 90 degrees: Effectively disables varying bevel.
    ' (angles are measured in radinas in GSCAD)
    IJDMfgPlateBevelRule_GetBevelDeviationAngle = (90# * 3.141529) / 180#

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1041, , "RULES")
End Function

 
