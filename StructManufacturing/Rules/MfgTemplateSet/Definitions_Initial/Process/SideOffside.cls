VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideOffside"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module: SideOffside
'
' Description:  Determines the process settings for the mfg template
'
' Author:
'
' Comments:
' 2006.01.11    Ram K.        creation
'*******************************************************************************
Option Explicit

Private Const MODULE = "MfgTemplateProcess.SideOffside"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    IJDMfgPlateUpSideRule_GetPlateUpSide = OffsetSide
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
End Function

 
