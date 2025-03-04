VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SideBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Base Side Process rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Private Const MODULE = "MfgTemplateProcess.SideBase"

Implements IJDMfgPlateUpSideRule

Private Function IJDMfgPlateUpSideRule_GetPlateUpSide(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart) As GSCADMfgRulesDefinitions.enumPlateSide
    Const METHOD = "IJDMfgPlateUpSideRule_GetPlateUpSide"
    On Error GoTo ErrorHandler
    
    IJDMfgPlateUpSideRule_GetPlateUpSide = BaseSide
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3005, , "RULES")
End Function
 
